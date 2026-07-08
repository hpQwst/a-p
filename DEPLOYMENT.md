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

Importante: a etapa de geracao final de PPT com graficos editaveis agora roda sem Microsoft Office/COM. O caminho validado edita o pacote Office Open XML de forma cirurgica: atualiza o workbook `.xlsx` embutido, preserva a estrutura ZIP/OPC original e atualiza o cache visual do grafico no `ppt/charts/chartX.xml`. Isso evita o problema do `python-pptx chart.replace_data()` e permite rodar em ECS Fargate Linux.

Portanto, a arquitetura de producao pode manter o fluxo inteiro no container Linux:

- `web/api`: FastAPI em ECS Fargate Linux para upload, preview, IA, geracao e download.
- `storage`: S3 para inputs, outputs, checkpoints e mapeamentos.
- `jobs`: execucao sincrona para arquivos pequenos ou fila SQS/worker Fargate para jobs grandes.

Para producao corporativa, a evolucao natural e:

- ALB com HTTPS na frente do FastAPI.
- Cognito ou IdP corporativo para autenticacao.
- S3 como storage definitivo de uploads e outputs.
- DynamoDB para status/metadados de jobs, projetos e execucoes.
- SQS para fila de analise/geracao.
- Um worker Fargate Linux separado consumindo a fila, se o volume ou o tempo de IA por slide exigir processamento assincrono.
- Security Group fechado para rede/VPN corporativa.

O container instala LibreOffice para dois usos auxiliares: recalcular datasources com formulas quando necessario e renderizar slides para preview/IA visual. A geracao final do PPT editavel nao depende do LibreOffice salvar o PPTX; ela continua sendo feita pelo writer OpenXML do projeto.

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
