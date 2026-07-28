<#
.SYNOPSIS
    Publica o QWST Auto PPT no AWS App Runner.

.DESCRIPTION
    Faz, em ordem: garante o repositorio ECR, constroi a imagem no CodeBuild (nao
    precisa de Docker na sua maquina), guarda os segredos no Secrets Manager e
    cria/atualiza a stack do App Runner.

    Rodar de novo publica uma versao nova: reconstroi a imagem e manda o App
    Runner trocar. Nenhuma maquina da equipe precisa ser atualizada.

.EXAMPLE
    .\infra\aws\deploy.ps1 -TeamPassword "senha-da-equipe"
#>
param(
    [string]$AppName = "squad4e5-auto-ppt",
    [string]$Region = "us-east-1",
    [string]$TeamPassword = "",
    [string]$OpenAIKey = "",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$Template = Join-Path $PSScriptRoot "apprunner.yaml"
$StackName = "$AppName-stack"

if (-not $AppName.StartsWith("squad4") -and -not $AppName.StartsWith("squad5")) {
    throw "AppName precisa comecar com squad4/squad5. Recursos fora disso pertencem a outras pessoas."
}
if (-not (Get-Command aws -ErrorAction SilentlyContinue)) {
    throw "AWS CLI nao encontrado."
}

$AccountId = (aws sts get-caller-identity --query Account --output text).Trim()
$RepositoryUri = "$AccountId.dkr.ecr.$Region.amazonaws.com/$AppName"
$BucketName = "$AppName-$AccountId"
Write-Host "Conta $AccountId | regiao $Region | app $AppName" -ForegroundColor Cyan

# --- segredos -------------------------------------------------------------
if (-not $OpenAIKey) {
    $envFile = Join-Path $RepoRoot ".env"
    if (Test-Path $envFile) {
        $match = Select-String -Path $envFile -Pattern '^\s*OPENAI_API_KEY\s*=\s*(.+)$' | Select-Object -First 1
        if ($match) { $OpenAIKey = $match.Matches[0].Groups[1].Value.Trim().Trim('"') }
    }
}
if (-not $OpenAIKey) { throw "Informe -OpenAIKey ou deixe OPENAI_API_KEY no .env." }
if (-not $TeamPassword) { throw "Informe -TeamPassword: e a senha que a equipe vai digitar para entrar." }

function Set-Secret([string]$Name, [string]$Value) {
    $arn = aws secretsmanager create-secret --name $Name --secret-string $Value --region $Region --query ARN --output text 2>$null
    if (-not $arn) {
        aws secretsmanager put-secret-value --secret-id $Name --secret-string $Value --region $Region --query ARN --output text | Out-Null
        $arn = aws secretsmanager describe-secret --secret-id $Name --region $Region --query ARN --output text
    }
    return $arn.Trim()
}

Write-Host "Guardando segredos no Secrets Manager..." -ForegroundColor Cyan
$OpenAISecretArn = Set-Secret "$AppName/openai-api-key" $OpenAIKey
$TeamPasswordSecretArn = Set-Secret "$AppName/team-password" $TeamPassword

# --- repositorio de imagens ----------------------------------------------
aws ecr describe-repositories --repository-names $AppName --region $Region 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Criando repositorio ECR $AppName..." -ForegroundColor Cyan
    aws ecr create-repository --repository-name $AppName --image-scanning-configuration scanOnPush=true --region $Region | Out-Null
}

# --- build da imagem no CodeBuild ----------------------------------------
$ImageTag = (git -C $RepoRoot rev-parse --short=12 HEAD 2>$null)
if (-not $ImageTag) { $ImageTag = Get-Date -Format "yyyyMMddHHmmss" }
$ImageUri = "${RepositoryUri}:${ImageTag}"

if (-not $SkipBuild) {
    Write-Host "Construindo a imagem no CodeBuild (sem Docker local)..." -ForegroundColor Cyan
    Write-Host "  Projeto CodeBuild esperado: $AppName-build" -ForegroundColor DarkGray
    Write-Host "  Se ainda nao existir, veja DEPLOYMENT.md secao 'Build da imagem'." -ForegroundColor DarkGray
    $buildId = aws codebuild start-build `
        --project-name "$AppName-build" `
        --environment-variables-override "name=IMAGE_REPO_NAME,value=$AppName,type=PLAINTEXT" "name=IMAGE_TAG,value=$ImageTag,type=PLAINTEXT" `
        --region $Region --query "build.id" --output text
    Write-Host "  build: $buildId" -ForegroundColor DarkGray
    do {
        Start-Sleep -Seconds 10
        $status = aws codebuild batch-get-builds --ids $buildId --region $Region --query "builds[0].buildStatus" --output text
        Write-Host "  status: $status" -ForegroundColor DarkGray
    } while ($status -eq "IN_PROGRESS")
    if ($status -ne "SUCCEEDED") { throw "Build falhou: $status" }
}

# --- stack ----------------------------------------------------------------
Write-Host "Publicando a stack $StackName..." -ForegroundColor Cyan
aws cloudformation deploy `
    --template-file $Template `
    --stack-name $StackName `
    --capabilities CAPABILITY_NAMED_IAM `
    --region $Region `
    --parameter-overrides `
        "AppName=$AppName" `
        "ImageUri=$ImageUri" `
        "OpenAISecretArn=$OpenAISecretArn" `
        "TeamPasswordSecretArn=$TeamPasswordSecretArn" `
        "BucketName=$BucketName"

$ServiceUrl = aws cloudformation describe-stacks --stack-name $StackName --region $Region `
    --query "Stacks[0].Outputs[?OutputKey=='ServiceUrl'].OutputValue" --output text

Write-Host ""
Write-Host "Pronto. Envie este endereco para a equipe:" -ForegroundColor Green
Write-Host "  $ServiceUrl" -ForegroundColor Green
Write-Host "A primeira visita pede a senha da equipe." -ForegroundColor Green
