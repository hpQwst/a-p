# Deploy e operacao

## Arquitetura inicial na AWS

O deploy atual usa uma arquitetura simples e suficiente para validar o produto:

- ECS Fargate roda o app FastAPI.
- CodeBuild monta a imagem Docker na nuvem, sem Docker local.
- ECR armazena a imagem.
- S3 fica preparado para guardar squads, projetos, execucoes, inputs, outputs e relatorios.
- Secrets Manager guarda `OPENAI_API_KEY`.
- CloudWatch Logs guarda logs do container.
- Security Group libera a porta `8501` para o CIDR configurado.

Para producao corporativa, a evolucao natural e:

- ALB com HTTPS na frente do FastAPI.
- Cognito ou IdP corporativo para autenticacao.
- S3 como storage definitivo de uploads e outputs.
- DynamoDB para status/metadados de jobs, projetos e execucoes.
- SQS para fila de analise/geracao.
- Um servico worker Fargate consumindo a fila.
- Security Group fechado para rede/VPN corporativa.

O core ja esta separado em `ppt_automator/`, a UI em `web/` e o ponto de worker em `worker/processor.py`, para permitir essa troca sem reescrever a logica de PowerPoint.

## Comandos

Deploy ou atualizacao:

```powershell
.\infra\aws\deploy_fargate.ps1 -AppName qwst-auto-ppt -Region us-east-1 -AllowedCidr 0.0.0.0/0
```

Pausar para economizar:

```powershell
.\infra\aws\stop_fargate.ps1 -AppName qwst-auto-ppt -Region us-east-1
```

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
