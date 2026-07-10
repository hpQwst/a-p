# Deploy e operacao

## Arquitetura inicial na AWS

O deploy atual usa uma arquitetura simples e suficiente para validar a v1:

- ECS Fargate Linux roda o app FastAPI continuamente para upload, preview, status e download.
- Preview e geracao sobem workers Fargate sob demanda via `RunTask`.
- CodeBuild monta a imagem Docker na nuvem, sem Docker local.
- ECR armazena a imagem.
- S3 guarda jobs, checkpoints, projetos, execucoes, inputs, outputs e relatorios.
- Secrets Manager guarda `OPENAI_API_KEY`.
- CloudWatch Logs guarda logs separados de web e worker.
- ALB com HTTPS publica a aplicacao; o acesso fica restrito ao CIDR configurado enquanto nao houver autenticacao.

Importante: a etapa de geracao final de PPT com graficos editaveis agora roda sem Microsoft Office/COM. O caminho validado edita o pacote Office Open XML de forma cirurgica: atualiza o workbook `.xlsx` embutido, preserva a estrutura ZIP/OPC original e atualiza o cache visual do grafico no `ppt/charts/chartX.xml`. Isso evita o problema do `python-pptx chart.replace_data()` e permite rodar em ECS Fargate Linux.

Portanto, a arquitetura de producao pode manter o fluxo inteiro em containers Linux:

- `web/api`: FastAPI em ECS Fargate Linux para upload, preview, IA, geracao e download.
- `storage`: S3 para inputs, outputs, checkpoints, mapeamentos e estado dos jobs.
- `jobs`: tasks Fargate sob demanda para `preview` e `generate`, sem EFS e sem Office.

Para a v1, nao e obrigatorio adicionar mais servicos. O proximo degrau natural, se houver aumento de volume ou necessidade de controle de concorrencia, e:

- Cognito ou IdP corporativo para autenticacao.
- DynamoDB para status/metadados de jobs, projetos e execucoes.
- SQS para fila de analise/geracao.
- Um worker Fargate Linux separado consumindo a fila, se o volume ou o tempo de IA por slide exigir orquestracao extra.

O container nao instala nem chama Microsoft Office, COM ou LibreOffice. Formulas de XLSX passam pelo avaliador interno deterministico; preview e IA usam apenas estrutura OpenXML e texto extraido. A geracao final continua sendo feita pelo writer OpenXML preservador.

O core ja esta separado em `ppt_automator/`, a UI em `web/` e o ponto de worker em `worker/processor.py`, para permitir essa troca sem reescrever a logica de PowerPoint.

## Comandos

Deploy ou atualizacao:

```powershell
.\infra\aws\deploy_v1.ps1 `
  -ImageUri "123456789012.dkr.ecr.us-east-1.amazonaws.com/qwst-auto-ppt@sha256:..." `
  -VpcId "vpc-..." `
  -PublicSubnetIds "subnet-a,subnet-b" `
  -CertificateArn "arn:aws:acm:us-east-1:123456789012:certificate/..." `
  -AllowedCidr "203.0.113.0/24" `
  -OpenAISecretArn "arn:aws:secretsmanager:us-east-1:123456789012:secret:..."
```

O script legado `infra/aws/deploy_fargate.ps1` foi movido para `infra/aws/legacy/` e nao deve ser usado como caminho principal de producao.

Todos os recursos criados pelo script devem receber a tag `Name=qwst-auto-ppt` para acompanhamento de custos.

## Git Azure DevOps

Quando quiser versionar este projeto no Azure DevOps:

```powershell
git init
git add .
git commit -m "Initial auto-ppt app"
git remote add origin https://qwst-equipe-tecnica@dev.azure.com/qwst-equipe-tecnica/qwst-equipe-tecnica/_git/qwst-auto-ppt
git push -u origin main
```

Antes do push, confira que `.env`, `workspace_data/`, `outputs/`, `.venv/` e arquivos grandes/sensiveis nao entraram no commit.
