# Infraestrutura AWS

Três arquivos, um caminho só:

| Arquivo | O que faz |
| --- | --- |
| `build.yaml` | CloudFormation: ECR, bucket de build e projeto CodeBuild |
| `apprunner.yaml` | CloudFormation: bucket do estado compartilhado, roles IAM, escala e o serviço App Runner |
| `deploy.ps1` | Publica: empacota o commit, envia ao S3, constrói, guarda segredos e sobe as stacks |

```powershell
.\infra\aws\deploy.ps1 -TeamPassword "a-senha-da-equipe"
```

Nada precisa ser configurado no console. O passo a passo está em
[`DEPLOYMENT.md`](../../DEPLOYMENT.md).

## Por que o build vem de um zip no S3

O código mora no Azure DevOps, e o CodeBuild **não lê Azure Repos** (as origens
aceitas são CodeCommit, GitHub, GitLab, Bitbucket, S3 e CodePipeline). Em vez de
espelhar o repositório ou guardar credencial de uma nuvem na outra, o
`deploy.ps1` empacota o commit atual com `git archive` e envia o zip.

Isso mantém o Azure DevOps como fonte da verdade, não depende de nenhum serviço
descontinuado e não exige integração entre as nuvens.

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
