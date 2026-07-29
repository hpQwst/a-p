# Deploy — AWS App Runner

O QWST Auto PPT roda como site interno, protegido pelo Microsoft Entra. Cada
usuário comum pertence a uma das squads `squad1`–`squad5` e só enxerga dados
dessa squad. Administradores podem selecionar qualquer squad e gerenciar
usuários em `/admin/users`.

## Arquitetura

```text
Azure DevOps  ──► deploy.ps1 ──► S3 (build) ──► CodeBuild ──► ECR
                                                                 │
Equipe ──HTTPS / Microsoft Entra──► App Runner squad4e5-auto-ppt │
                                      1 vCPU / 2 GB, máx. 1       │
                                                   └──────────────┘
                                                   └──► S3 de estado
```

O código vive no Azure DevOps. A AWS não se conecta ao repositório:
`deploy.ps1` empacota o commit atual com `git archive`, envia o zip ao S3 e
dispara o CodeBuild. Não é necessário Docker local.

Um único container recebe uploads, analisa, chama IA quando necessário e gera o
PPT. O estado durável fica no S3. Entradas imutáveis são persistidas uma vez;
salvamentos manuais e automáticos posteriores gravam apenas estado e caches
pequenos. Se a instância reiniciar durante um preview, o trabalho é restaurado e
a análise pendente recomeça.

A geração roda sem Microsoft Office, por edição cirúrgica do pacote OOXML. O
Excel embutido e o cache visual do gráfico são atualizados sem quebrar
“Editar dados”.

O limite de uma instância é deliberado: duas instâncias gravando o mesmo objeto
JSON simultaneamente poderiam perder uma atualização. Antes de aumentar
`MaxSize`, implemente escrita condicional por ETag em `project_store.py`.

## Pré-requisitos

- AWS CLI autenticado na conta correta, região `us-east-1`;
- chave OpenAI;
- aplicativo Microsoft Entra single-tenant;
- URI de redirecionamento exatamente
  `https://<url-do-app>/auth/callback`;
- e-mail do administrador inicial.

Só crie ou altere recursos iniciados por `squad4` ou `squad5`. Os recursos deste
projeto usam `squad4e5-auto-ppt`.

## Publicar

O deploy de produção exige Entra e desativa a senha compartilhada:

```powershell
.\infra\aws\deploy.ps1 `
  -EntraTenantId "..." `
  -EntraClientId "..." `
  -EntraClientSecret "..." `
  -EntraRedirectUri "https://.../auth/callback" `
  -BootstrapAdminEmails "hugo.rocha@qwst.co"
```

O script:

1. cria ou atualiza build bucket, ECR e CodeBuild;
2. empacota o último commit e envia o zip ao S3;
3. constrói e publica a imagem;
4. grava OpenAI, Entra e sessão no Secrets Manager;
5. cria ou atualiza S3 de estado e App Runner;
6. imprime a URL.

O script publica o último commit, não alterações soltas no disco. Faça commit e
push para Azure DevOps e GitHub antes do deploy.

## Variáveis de produção

Definidas pela stack:

| Variável | Valor |
| --- | --- |
| `AUTO_PPT_STORAGE_BACKEND` | `s3` |
| `AUTO_PPT_S3_BUCKET` | `squad4e5-auto-ppt-<conta>` |
| `AUTO_PPT_RUNTIME_ROOT` | `/tmp/auto-ppt-jobs` |
| `AUTO_PPT_TEAM_PASSWORD_ENABLED` | `0` |
| `AUTO_PPT_BOOTSTRAP_ADMINS` | lista de e-mails |
| `ENTRA_TENANT_ID` | diretório corporativo |
| `ENTRA_CLIENT_ID` | aplicativo corporativo |
| `ENTRA_REDIRECT_URI` | callback HTTPS |
| `OPENAI_API_KEY` | secret |
| `ENTRA_CLIENT_SECRET` | secret |
| `AUTO_PPT_SESSION_SECRET` | secret |

A senha antiga pode continuar existindo no Secrets Manager por histórico, mas
não é injetada no App Runner e não autentica ninguém.

## Usuários e isolamento

No primeiro login:

- e-mails em `AUTO_PPT_BOOTSTRAP_ADMINS` viram administradores;
- demais usuários escolhem uma única squad entre `squad1` e `squad5`;
- usuários comuns não podem trocar a própria squad;
- administradores podem alterar squad, ativar/desativar e promover/rebaixar;
- alterações administrativas são auditadas no S3.

O middleware aplica isolamento tanto nas listagens quanto em URLs diretas de
projetos e jobs.

## Verificar

```powershell
& "C:\Program Files\Amazon\AWSCLIV2\aws.exe" apprunner list-services `
  --profile default --region us-east-1 `
  --query "ServiceSummaryList[?ServiceName=='squad4e5-auto-ppt']"
```

Depois do deploy:

1. confirme tela com apenas login Microsoft;
2. entre como administrador e revise `/admin/users`;
3. entre como usuário comum e confirme o isolamento por squad;
4. suba PPTX e planilhas, gere preview, salve e baixe o PPT;
5. confirme progresso por objeto durante preview e geração;
6. retome o projeto depois de uma nova sessão.

Logs ficam no CloudWatch em `/aws/apprunner/squad4e5-auto-ppt`.

## Memória e tamanho de upload

Benchmarks locais com cópias, sem alterar os originais:

| Caso | Slides | Arquivos | Pico observado |
| --- | ---: | ---: | ---: |
| TIM (análise, 178 abas) | 89 | 203,4 MB + 0,8 MB | 352,1 MiB |
| Natura CB | 40 | 26,7 MB + 4,6 MB | 487,6 MB |
| Natura CB (geração, 4 abas) | 118 | 80,6 MB + 4,6 MB | 681,9 MiB |

Resultado: 2 GB continuam adequados; não aumentar memória. A interface mostra
o tamanho combinado apenas como informação, sem fingir que ele prevê o pico de
RAM. O limite individual padrão subiu para 350 MB
(`AUTO_PPT_MAX_UPLOAD_MB`), sem mudança de custo de infraestrutura.

Para repetir uma medição:

```powershell
.\.venv\Scripts\python.exe -m scripts.measure_memory_case `
  --pptx "caminho\modelo.pptx" `
  --xlsx "caminho\dados.xlsx"
```

## Matching em planilhas com várias abas

Cada aba do XLSX é uma fonte independente. No caso TIM real são 639 objetos
contra 178 abas. O `cProfile` mediu 11,6 s para apenas um objeto antes da
otimização; a projeção para o deck completo era de 23 a 48 minutos.

As características normalizadas de cada aba agora são calculadas uma vez por
análise, e o Hungarian trabalha na matriz retangular real, sem criar 178 linhas
artificiais para um slide com um único objeto. Resultado medido:

- TIM completo: 66,5–69,7 s, 148 matches e 352,1 MiB de pico;
- Natura CB 118 slides: 53,3 s, com a mesma assinatura dos 3 matches anteriores;
- geração Natura CB: 61,5 s e 681,9 MiB de pico;
- nenhum candidato é cortado e a atribuição 1:1 continua ativa.

## Custos

`configure-cost-controls.ps1` mantém um orçamento mensal de US$ 20, filtrado
pela tag `Name=squad4e5-auto-ppt`, com alerta real em 80% e previsto em 100%:

```powershell
.\infra\aws\configure-cost-controls.ps1 `
  -AlertEmail "hugo.rocha@qwst.co"
```

A ativação da tag `Name` como cost allocation tag exige a conta pagadora. Se a
conta vinculada receber `AccessDenied`, o administrador de Billing deve ativar
essa tag; o orçamento já pode existir, mas o filtro só passa a contabilizar
custos depois da ativação.

O deploy também cria:

- alarme `squad4e5-auto-ppt-5xx`, que envia e-mail quando houver ao menos uma
  resposta 5xx em cinco minutos (a assinatura SNS precisa ser confirmada);
- lifecycle S3 de 180 dias apenas para artefatos de `runs`, 30 dias para versões
  antigas e 7 dias para uploads multipart incompletos. O checkpoint atual do
  projeto não expira.

App Runner cobra memória provisionada continuamente e CPU durante requisições.
Para pausar sem destruir estado:

```powershell
& "C:\Program Files\Amazon\AWSCLIV2\aws.exe" apprunner pause-service `
  --service-arn <arn-do-servico> --profile default --region us-east-1
```

## Recuperar dados

O bucket tem versionamento. Para localizar versões anteriores:

```powershell
& "C:\Program Files\Amazon\AWSCLIV2\aws.exe" s3api list-object-versions `
  --bucket squad4e5-auto-ppt-<conta> `
  --prefix auto-ppt/squads/ `
  --profile default --region us-east-1
```

## Próximos degraus

- escrita condicional por ETag antes de permitir mais de uma instância;
- fila externa apenas se o volume simultâneo justificar o custo;
- `AUTO_PPT_ASYNC_GENERATION=1` se algum deck exceder o limite da requisição.
