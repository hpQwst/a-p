# Deploy — AWS App Runner

O QWST Auto PPT roda como um site interno: a equipe abre uma URL no navegador,
digita a senha compartilhada e usa. Não há programa para instalar em máquina
nenhuma, e todo mundo enxerga os mesmos projetos e mapeamentos.

## Arquitetura

```
Azure DevOps  ──► (repositório: fonte da verdade do código)
      │
      │  deploy.ps1 empacota o commit e envia
      ▼
S3 (build) ──► CodeBuild ──► ECR ──┐
                                   │
Equipe (navegador) ──HTTPS──► App Runner  squad4e5-auto-ppt
                                   │       1 vCPU / 2 GB, no máx. 1 instância
                                   └──► S3  squad4e5-auto-ppt-<conta>
                                            (estado compartilhado)
```

O código vive no Azure DevOps. A AWS **não se conecta ao repositório**: o
`deploy.ps1` empacota o commit atual com `git archive`, envia o zip para o S3 e o
CodeBuild constrói a partir dele. Isso evita credencial cruzada entre nuvens e
funciona com qualquer serviço de repositório — o CodeBuild, aliás, não consegue
ler Azure Repos nativamente.

Um único container faz tudo: recebe o upload, analisa, chama a IA quando
necessário e gera o PPT. Não existem workers separados nem fila.

A geração do PPT roda sem Microsoft Office: o pacote Office Open XML é editado de
forma cirúrgica (atualiza o `.xlsx` embutido, preserva a estrutura ZIP/OPC e
atualiza o cache visual do gráfico). É isso que permite rodar em container Linux
mantendo o "Editar dados" funcional.

**Por que no máximo uma instância:** o estado compartilhado é gravado como JSON no
S3. Cada gravação isolada é atômica, mas duas instâncias gravando o mesmo arquivo
ao mesmo tempo poderiam perder uma atualização. Com o uso atual (poucas vezes por
mês) uma instância sobra. Antes de aumentar `MaxSize`, é preciso implementar
escrita condicional por ETag no `project_store.py`.

## Pré-requisitos

- AWS CLI logado com permissão na conta (`aws sts get-caller-identity`)
- Região `us-east-1` — **App Runner não existe em `sa-east-1`**
- `OPENAI_API_KEY` no `.env` local (ou passado por parâmetro)
- Uma senha para a equipe

Não é preciso ter Docker nem conectar a AWS ao repositório: a imagem é
construída no CodeBuild a partir de um zip.

> **Limite de recursos:** só criar/alterar recursos que comecem com `squad4`/`squad5`.
> Tudo o mais na conta pertence a outras pessoas.

## Publicar

Não há configuração manual no console: o próprio script cria tudo.

```powershell
.\infra\aws\deploy.ps1 -TeamPassword "a-senha-da-equipe"
```

O que ele faz, em ordem:

1. Sobe a stack de build (`build.yaml`): repositório ECR, bucket de build e projeto CodeBuild
2. Empacota o **último commit** com `git archive` e envia o zip para o S3
3. Dispara o CodeBuild e espera a imagem ficar pronta
4. Guarda a chave da OpenAI e a senha da equipe no Secrets Manager
5. Sobe a stack da aplicação (`apprunner.yaml`) e imprime a URL

Envie a URL para a equipe — a primeira visita pede a senha.

> O deploy publica o **último commit**, não o que está no disco. Se houver
> alterações não commitadas, o script avisa. Commite e envie para o Azure DevOps
> antes de publicar.

## Publicar uma versão nova

O mesmo comando. Ele constrói a imagem do commit atual e manda o App Runner
trocar. Ninguém precisa atualizar nada na própria máquina.

```powershell
.\infra\aws\deploy.ps1 -TeamPassword "a-senha-da-equipe"
```

Para trocar apenas a senha da equipe, sem reconstruir a imagem:

```powershell
.\infra\aws\deploy.ps1 -TeamPassword "nova-senha" -SkipBuild
```

## Variáveis no ambiente publicado

Definidas pela stack — não altere à mão no console:

| Variável | Valor |
| --- | --- |
| `AUTO_PPT_STORAGE_BACKEND` | `s3` |
| `AUTO_PPT_S3_BUCKET` | `squad4e5-auto-ppt-<conta>` |
| `AUTO_PPT_RUNTIME_ROOT` | `/tmp/auto-ppt-jobs` |
| `OPENAI_API_KEY` | secret |
| `AUTO_PPT_TEAM_PASSWORD` | secret |

`AUTO_PPT_TEAM_PASSWORD` é o que protege a URL. **Se ficar vazia e o login
Microsoft não estiver configurado, o app fica aberto para qualquer pessoa que
tenha o link.**

## Login com a conta Microsoft (opcional)

Além da senha da equipe, o app aceita login corporativo via Microsoft Entra. As
duas formas convivem: a senha continua funcionando como reserva.

Para ligar, o registro do aplicativo no Entra precisa ter como URI de redirecionamento
exatamente `https://<url-do-app>/auth/callback`, e estas variáveis precisam chegar
ao container:

| Variável | Observação |
| --- | --- |
| `ENTRA_TENANT_ID` | id do diretório |
| `ENTRA_CLIENT_ID` | id do aplicativo |
| `ENTRA_CLIENT_SECRET` | guardar no Secrets Manager, nunca no código |
| `ENTRA_REDIRECT_URI` | idêntico ao cadastrado, com `/auth/callback` no fim |
| `AUTO_PPT_SESSION_SECRET` | assina o cookie de sessão |

O aplicativo é single-tenant: contas de outro diretório são recusadas mesmo que a
Microsoft autentique com sucesso. A tela de login avisa se a configuração estiver
pela metade ou com o endereço de retorno malformado.

## Verificar

```powershell
aws apprunner list-services --region us-east-1 --query "ServiceSummaryList[?ServiceName=='squad4e5-auto-ppt']"
```

Logs da aplicação ficam no CloudWatch, em `/aws/apprunner/squad4e5-auto-ppt`.

Depois de subir, teste o ciclo inteiro: entrar com a senha, subir um `.pptx` com
as planilhas, gerar o preview e baixar o PPT. Depois abra de outra máquina e
confirme que o projeto e o mapeamento aparecem — é isso que prova que o estado
compartilhado está funcionando.

## Custo

O App Runner cobra a memória provisionada continuamente e a CPU só durante as
requisições; some ECR e S3, ambos baratos no volume deste projeto. A conta já tem
outro serviço App Runner igual (1 vCPU / 2 GB), então a fatura atual é a melhor
referência de custo real.

Para pausar sem destruir nada:

```powershell
aws apprunner pause-service --service-arn <arn-do-servico> --region us-east-1
```

## Recuperar um mapeamento sobrescrito

O bucket tem versionamento ligado. Se um mapeamento for salvo errado, dá para
listar e restaurar a versão anterior:

```powershell
aws s3api list-object-versions --bucket squad4e5-auto-ppt-<conta> --prefix auto-ppt/squads/
```

## Próximos degraus (ainda não construídos)

Se o volume crescer ou surgir necessidade de mais controle:

- Login corporativo (OIDC) no lugar da senha compartilhada
- Escrita condicional por ETag no `project_store.py`, para permitir mais de uma instância
- `AUTO_PPT_ASYNC_GENERATION=1` se algum deck grande estourar o tempo limite da requisição
