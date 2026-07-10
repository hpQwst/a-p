# Infraestrutura AWS

## Caminho recomendado para a v1

`v1-fargate.yaml` cria um ECS Fargate Service para a interface e uma task definition separada para workers sob demanda, atras de um ALB HTTPS, com S3 privado, IAM de menor privilegio, Secrets Manager e CloudWatch Logs.

Essa topologia mantem `DesiredCount=1` apenas para a interface server-rendered. Preview e geração são isolados em workers Fargate e sincronizados no S3; não usam disco compartilhado nem aplicativos Office.

Pre-requisitos:

1. Imagem publicada no ECR com tag de commit ou digest.
2. VPC com pelo menos duas subnets publicas.
3. Certificado ACM na mesma regiao.
4. Secret do OpenAI no Secrets Manager, se a IA for usada.
5. CIDR corporativo/VPN. O script recusa `0.0.0.0/0` enquanto nao houver autenticacao.

Deploy:

```powershell
.\infra\aws\deploy_v1.ps1 `
  -ImageUri "123456789012.dkr.ecr.us-east-1.amazonaws.com/qwst-auto-ppt@sha256:..." `
  -VpcId "vpc-..." `
  -PublicSubnetIds "subnet-a,subnet-b" `
  -CertificateArn "arn:aws:acm:us-east-1:123456789012:certificate/..." `
  -AllowedCidr "203.0.113.0/24" `
  -OpenAISecretArn "arn:aws:secretsmanager:us-east-1:123456789012:secret:..."
```

Os scripts em `legacy/` preservam o prototipo anterior, que publica a porta 8501 diretamente por IP e nao deve ser usado como deploy de producao.

## Crescimento posterior

SQS e DynamoDB ficam para uma fase posterior, se houver necessidade de controle de concorrência, agendamento ou consultas de jobs em escala. A arquitetura atual já isola o processamento em tasks Fargate e mantém o estado compartilhado no S3.
