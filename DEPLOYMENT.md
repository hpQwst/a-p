# Deploy e operacao

## Arquitetura inicial na AWS

O primeiro deploy usa uma arquitetura simples e suficiente para uso pontual:

- ECS Fargate roda o app Streamlit.
- CodeBuild monta a imagem Docker na nuvem, sem Docker local.
- ECR armazena a imagem.
- S3 guarda squads, projetos, execucoes, inputs, outputs e relatorios.
- Secrets Manager guarda `OPENAI_API_KEY`.
- CloudWatch Logs guarda logs do container.
- Security Group libera a porta `8501` para o CIDR configurado.

Para producao corporativa, a evolucao natural e colocar um Application Load Balancer com HTTPS, autenticacao via Cognito/IdP corporativo e limitar acesso por rede/VPN.

## Comandos

Deploy ou atualizacao:

```powershell
.\infra\aws\deploy_fargate.ps1 -AppName qwst-auto-ppt -Region sa-east-1 -AllowedCidr 0.0.0.0/0
```

Pausar para economizar:

```powershell
.\infra\aws\stop_fargate.ps1 -AppName qwst-auto-ppt -Region sa-east-1
```

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
