param(
    [Parameter(Mandatory = $true)][string]$ImageUri,
    [Parameter(Mandatory = $true)][string]$VpcId,
    [Parameter(Mandatory = $true)][string]$PublicSubnetIds,
    [Parameter(Mandatory = $true)][string]$CertificateArn,
    [Parameter(Mandatory = $true)][string]$AllowedCidr,
    [string]$StackName = "qwst-auto-ppt-v1",
    [string]$AppName = "qwst-auto-ppt",
    [string]$Region = "us-east-1",
    [string]$OpenAISecretArn = "",
    [string]$ExistingBucketName = ""
)

$ErrorActionPreference = "Stop"
$Template = Join-Path $PSScriptRoot "v1-fargate.yaml"

if (-not (Get-Command aws -ErrorAction SilentlyContinue)) {
    throw "AWS CLI nao encontrado."
}
if ($AllowedCidr -eq "0.0.0.0/0") {
    throw "AllowedCidr=0.0.0.0/0 foi bloqueado. Use o CIDR corporativo/VPN ate habilitar autenticacao."
}
if ($ImageUri -match ":latest$") {
    Write-Warning "Prefira uma imagem ECR imutavel por digest ou tag de commit, nao :latest."
}

aws cloudformation validate-template --template-body "file://$Template" --region $Region | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Template CloudFormation invalido." }

$Parameters = @(
    "AppName=$AppName",
    "ImageUri=$ImageUri",
    "VpcId=$VpcId",
    "PublicSubnetIds=$PublicSubnetIds",
    "CertificateArn=$CertificateArn",
    "AllowedCidr=$AllowedCidr",
    "OpenAISecretArn=$OpenAISecretArn",
    "ExistingBucketName=$ExistingBucketName"
)

aws cloudformation deploy `
    --stack-name $StackName `
    --template-file $Template `
    --capabilities CAPABILITY_NAMED_IAM `
    --parameter-overrides $Parameters `
    --no-fail-on-empty-changeset `
    --tags "Name=$AppName" "Environment=v1" `
    --region $Region
if ($LASTEXITCODE -ne 0) { throw "Falha no deploy do stack $StackName." }

aws cloudformation describe-stacks `
    --stack-name $StackName `
    --query "Stacks[0].Outputs" `
    --output table `
    --region $Region
