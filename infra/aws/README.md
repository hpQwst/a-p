# Infraestrutura AWS

Dois arquivos, um caminho só:

| Arquivo | O que faz |
| --- | --- |
| `apprunner.yaml` | CloudFormation: bucket S3 do estado compartilhado, roles IAM, escala e o serviço App Runner |
| `deploy.ps1` | Publica: garante o ECR, constrói no CodeBuild, guarda os segredos e sobe a stack |

```powershell
.\infra\aws\deploy.ps1 -TeamPassword "a-senha-da-equipe"
```

O passo a passo completo, incluindo a configuração única do projeto CodeBuild,
está em [`DEPLOYMENT.md`](../../DEPLOYMENT.md).

## Regras que valem aqui

- **Só criar ou alterar recursos que comecem com `squad4`/`squad5`.** Todo o resto
  da conta pertence a outras pessoas. O `deploy.ps1` recusa outros nomes.
- **Região `us-east-1`.** App Runner não existe em `sa-east-1`.
- **No máximo uma instância.** O estado compartilhado é JSON no S3; duas
  instâncias gravando ao mesmo tempo poderiam perder uma atualização. Aumentar
  `MaxSize` exige antes escrita condicional por ETag no `project_store.py`.

## Histórico

A arquitetura anterior (ECS Fargate + ALB + workers sob demanda) foi removida:
nunca chegou a ser publicada e trazia complexidade que o volume de uso não
justifica. O código dos workers Fargate saiu junto. Está tudo no histórico do git
se algum dia for preciso consultar.
