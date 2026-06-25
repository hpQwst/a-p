# Deploy e operacao

## Arquitetura inicial na AWS

O deploy atual usa uma arquitetura simples e suficiente para validar o produto:

- ECS Fargate Linux roda o app FastAPI e pode fazer upload, preview, IA, status e download de artefatos prontos.
- CodeBuild monta a imagem Docker na nuvem, sem Docker local.
- ECR armazena a imagem.
- S3 fica preparado para guardar squads, projetos, execucoes, inputs, outputs e relatorios.
- Secrets Manager guarda `OPENAI_API_KEY`.
- CloudWatch Logs guarda logs do container.
- Security Group libera a porta `8501` para o CIDR configurado.

Importante: a etapa de geracao final de PPT com graficos editaveis nao deve rodar em Linux/Fargate usando escrita XML manual do workbook embutido. O PowerPoint abre o grafico, mas o comando `Editar dados` pode quebrar o vinculo interno do Excel. Para preservar `Editar dados`, o workbook embutido precisa ser salvo por um motor compativel com Office. Hoje a via validada e Microsoft Excel via COM no Windows.

Portanto, a arquitetura de producao deve separar:

- `web/api`: FastAPI em ECS Fargate Linux, barato e sempre disponivel.
- `generation-worker`: Windows com Microsoft Excel instalado/licenciado, consumindo jobs de uma fila e salvando o PPT final no S3.

Para producao corporativa, a evolucao natural e:

- ALB com HTTPS na frente do FastAPI.
- Cognito ou IdP corporativo para autenticacao.
- S3 como storage definitivo de uploads e outputs.
- DynamoDB para status/metadados de jobs, projetos e execucoes.
- SQS para fila de analise/geracao.
- Um worker Windows/Office consumindo a fila. Pode ser EC2 Windows iniciado sob demanda, Auto Scaling Group Windows com desired count 0/1, ou uma alternativa corporativa de RPA/Office. O Fargate Linux atual nao atende a essa etapa.
- Security Group fechado para rede/VPN corporativa.

AWS Fargate tambem suporta containers Windows em ECS, mas isso nao resolve sozinho a exigencia de Excel COM: o container precisaria ter Office instalado, licenciado e operando de forma estavel. A propria Microsoft alerta que automacao de Office em ambiente server-side/unattended e uma abordagem com riscos operacionais; se for usada, deve ser isolada, monitorada e reiniciavel. Para o nosso volume baixo de uso, o desenho mais controlavel e um worker Windows dedicado e ligado apenas quando houver jobs.

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
