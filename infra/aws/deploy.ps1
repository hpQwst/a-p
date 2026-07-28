<#
.SYNOPSIS
    Publica o QWST Auto PPT no AWS App Runner.

.DESCRIPTION
    O codigo e enviado como um zip para o S3 e construido no CodeBuild, entao a
    AWS nao precisa de conexao com o repositorio. O Azure DevOps continua sendo a
    fonte da verdade do codigo; a AWS so recebe o conteudo do commit atual.

    Nao e preciso ter Docker na maquina, nem configurar nada no console.

    Rodar de novo publica uma versao nova.

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
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$SourceKey = "source/source.zip"

if (-not $AppName.StartsWith("squad4") -and -not $AppName.StartsWith("squad5")) {
    throw "AppName precisa comecar com squad4/squad5. Recursos fora disso pertencem a outras pessoas."
}
foreach ($tool in @("aws", "git")) {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) { throw "$tool nao encontrado." }
}

$AccountId = (aws sts get-caller-identity --query Account --output text).Trim()
Write-Host "Conta $AccountId | regiao $Region | app $AppName" -ForegroundColor Cyan

# --- 1. infraestrutura de build ------------------------------------------
Write-Host "[1/5] Infraestrutura de build (ECR, bucket, CodeBuild)..." -ForegroundColor Cyan
aws cloudformation deploy `
    --template-file (Join-Path $PSScriptRoot "build.yaml") `
    --stack-name "$AppName-build" `
    --capabilities CAPABILITY_NAMED_IAM `
    --region $Region `
    --parameter-overrides "AppName=$AppName" "SourceObjectKey=$SourceKey" | Out-Null

$BuildBucket = aws cloudformation describe-stacks --stack-name "$AppName-build" --region $Region `
    --query "Stacks[0].Outputs[?OutputKey=='BuildBucketName'].OutputValue" --output text
$RepositoryUri = aws cloudformation describe-stacks --stack-name "$AppName-build" --region $Region `
    --query "Stacks[0].Outputs[?OutputKey=='RepositoryUri'].OutputValue" --output text

# --- 2. empacota o commit atual ------------------------------------------
$ImageTag = (git -C $RepoRoot rev-parse --short=12 HEAD).Trim()
$ImageUri = "${RepositoryUri}:${ImageTag}"

if (-not $SkipBuild) {
    $dirty = git -C $RepoRoot status --porcelain
    if ($dirty) {
        Write-Warning "Ha alteracoes nao commitadas. O deploy publica o ULTIMO COMMIT ($ImageTag), nao o que esta no disco."
    }

    Write-Host "[2/5] Empacotando o commit $ImageTag..." -ForegroundColor Cyan
    # git archive respeita o .gitignore e nao leva .git, workspace_data nem .env.
    $zipPath = Join-Path ([System.IO.Path]::GetTempPath()) "$AppName-$ImageTag.zip"
    git -C $RepoRoot archive --format=zip --output $zipPath HEAD
    if (-not (Test-Path $zipPath)) { throw "Falha ao empacotar o codigo." }

    Write-Host "[3/5] Enviando para s3://$BuildBucket/$SourceKey..." -ForegroundColor Cyan
    aws s3 cp $zipPath "s3://$BuildBucket/$SourceKey" --region $Region | Out-Null
    Remove-Item $zipPath -Force

    # --- 3. build --------------------------------------------------------
    Write-Host "[4/5] Construindo a imagem no CodeBuild..." -ForegroundColor Cyan
    $buildId = aws codebuild start-build `
        --project-name "$AppName-build" `
        --environment-variables-override "name=IMAGE_TAG,value=$ImageTag,type=PLAINTEXT" `
        --region $Region --query "build.id" --output text
    do {
        Start-Sleep -Seconds 10
        $status = (aws codebuild batch-get-builds --ids $buildId --region $Region --query "builds[0].buildStatus" --output text).Trim()
        Write-Host "      $status" -ForegroundColor DarkGray
    } while ($status -eq "IN_PROGRESS")
    if ($status -ne "SUCCEEDED") {
        throw "Build falhou ($status). Logs: aws codebuild batch-get-builds --ids $buildId --region $Region"
    }
}

# --- 4. segredos ----------------------------------------------------------
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
        aws secretsmanager put-secret-value --secret-id $Name --secret-string $Value --region $Region | Out-Null
        $arn = aws secretsmanager describe-secret --secret-id $Name --region $Region --query ARN --output text
    }
    return $arn.Trim()
}

$OpenAISecretArn = Set-Secret "$AppName/openai-api-key" $OpenAIKey
$TeamPasswordSecretArn = Set-Secret "$AppName/team-password" $TeamPassword

# --- 5. aplicacao ---------------------------------------------------------
Write-Host "[5/5] Publicando a aplicacao..." -ForegroundColor Cyan
aws cloudformation deploy `
    --template-file (Join-Path $PSScriptRoot "apprunner.yaml") `
    --stack-name "$AppName-stack" `
    --capabilities CAPABILITY_NAMED_IAM `
    --region $Region `
    --parameter-overrides `
        "AppName=$AppName" `
        "ImageUri=$ImageUri" `
        "OpenAISecretArn=$OpenAISecretArn" `
        "TeamPasswordSecretArn=$TeamPasswordSecretArn" `
        "BucketName=$AppName-$AccountId" | Out-Null

$ServiceUrl = aws cloudformation describe-stacks --stack-name "$AppName-stack" --region $Region `
    --query "Stacks[0].Outputs[?OutputKey=='ServiceUrl'].OutputValue" --output text

Write-Host ""
Write-Host "Pronto. Envie este endereco para a equipe:" -ForegroundColor Green
Write-Host "  $ServiceUrl" -ForegroundColor Green
Write-Host "A primeira visita pede a senha da equipe." -ForegroundColor Green
