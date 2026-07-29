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
    .\infra\aws\deploy.ps1 -EntraTenantId "..." -EntraClientId "..." -EntraClientSecret "..." -EntraRedirectUri "https://.../auth/callback"
#>
param(
    [string]$AppName = "squad4e5-auto-ppt",
    [string]$ExpectedAccountId = "134164930693",
    [string]$Region = "us-east-1",
    [string]$OpenAIKey = "",
    [string]$EntraTenantId = "",
    [string]$EntraClientId = "",
    [string]$EntraClientSecret = "",
    [string]$EntraRedirectUri = "",
    [string]$BootstrapAdminEmails = "hugo.rocha@qwst.co",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$SourceKey = "source/source.zip"
$AwsCli = "C:\Program Files\Amazon\AWSCLIV2\aws.exe"

if (-not $AppName.StartsWith("squad4") -and -not $AppName.StartsWith("squad5")) {
    throw "AppName precisa comecar com squad4/squad5. Recursos fora disso pertencem a outras pessoas."
}
if (-not (Test-Path -LiteralPath $AwsCli)) { throw "AWS CLI nao encontrada em $AwsCli." }
if (-not (Get-Command "git" -ErrorAction SilentlyContinue)) { throw "git nao encontrado." }

$AccountId = (& $AwsCli sts get-caller-identity --profile default --region $Region --query Account --output text).Trim()
if ($AccountId -ne $ExpectedAccountId) {
    throw "Conta AWS inesperada: $AccountId."
}
Write-Host "Conta $AccountId | regiao $Region | app $AppName" -ForegroundColor Cyan

# --- 1. infraestrutura de build ------------------------------------------
Write-Host "[1/5] Infraestrutura de build (ECR, bucket, CodeBuild)..." -ForegroundColor Cyan
& $AwsCli cloudformation deploy `
    --template-file (Join-Path $PSScriptRoot "build.yaml") `
    --stack-name "$AppName-build" `
    --capabilities CAPABILITY_NAMED_IAM `
    --profile default `
    --region $Region `
    --parameter-overrides "AppName=$AppName" "SourceObjectKey=$SourceKey" | Out-Null

$BuildBucket = & $AwsCli cloudformation describe-stacks --stack-name "$AppName-build" --profile default --region $Region `
    --query "Stacks[0].Outputs[?OutputKey=='BuildBucketName'].OutputValue" --output text
$RepositoryUri = & $AwsCli cloudformation describe-stacks --stack-name "$AppName-build" --profile default --region $Region `
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
    & $AwsCli s3 cp $zipPath "s3://$BuildBucket/$SourceKey" --profile default --region $Region | Out-Null
    Remove-Item $zipPath -Force

    # --- 3. build --------------------------------------------------------
    Write-Host "[4/5] Construindo a imagem no CodeBuild..." -ForegroundColor Cyan
    $buildId = & $AwsCli codebuild start-build `
        --project-name "$AppName-build" `
        --environment-variables-override "name=IMAGE_TAG,value=$ImageTag,type=PLAINTEXT" `
        --profile default --region $Region --query "build.id" --output text
    do {
        Start-Sleep -Seconds 10
        $status = (& $AwsCli codebuild batch-get-builds --ids $buildId --profile default --region $Region --query "builds[0].buildStatus" --output text).Trim()
        Write-Host "      $status" -ForegroundColor DarkGray
    } while ($status -eq "IN_PROGRESS")
    if ($status -ne "SUCCEEDED") {
        throw "Build falhou ($status). Consulte o build $buildId no CodeBuild."
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

function Set-Secret([string]$Name, [string]$Value) {
    # Sem redirecionar stderr: no PowerShell 5.1 isso vira NativeCommandError e,
    # com ErrorActionPreference=Stop, aborta o script mesmo quando o comando so
    # avisou que o secret ja existe. Decidimos pelo codigo de saida.
    $previous = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        # list-secrets devolve vazio quando nao existe. describe-secret serviria,
        # mas escreve no stderr e polui a saida do deploy na primeira execucao.
        $existing = & $AwsCli secretsmanager list-secrets --profile default --region $Region `
            --filters "Key=name,Values=$Name" --query "SecretList[?Name=='$Name'].ARN | [0]" --output text
        if ($LASTEXITCODE -eq 0 -and $existing -and $existing -ne "None") {
            & $AwsCli secretsmanager put-secret-value --secret-id $Name --secret-string $Value --profile default --region $Region | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "Nao consegui atualizar o secret $Name." }
            & $AwsCli secretsmanager tag-resource --secret-id ($existing.Trim()) --tags "Key=Name,Value=$AppName" --profile default --region $Region | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "Nao consegui etiquetar o secret $Name." }
            return $existing.Trim()
        }
        $created = & $AwsCli secretsmanager create-secret --name $Name --secret-string $Value `
            --tags "Key=Name,Value=$AppName" --profile default --region $Region --query ARN --output text
        if ($LASTEXITCODE -ne 0 -or -not $created) { throw "Nao consegui criar o secret $Name." }
        return $created.Trim()
    }
    finally {
        $ErrorActionPreference = $previous
    }
}

$OpenAISecretArn = Set-Secret "$AppName/openai-api-key" $OpenAIKey

# Chave de assinatura do cookie. Gerada uma vez e reaproveitada: trocar a cada
# deploy derrubaria a sessao de quem estivesse usando.
$sessionArn = & $AwsCli secretsmanager list-secrets --profile default --region $Region `
    --filters "Key=name,Values=$AppName/session-secret" `
    --query "SecretList[?Name=='$AppName/session-secret'].ARN | [0]" --output text
if (-not $sessionArn -or $sessionArn -eq "None") {
    $bytes = New-Object byte[] 32
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
    $sessionArn = Set-Secret "$AppName/session-secret" ([Convert]::ToBase64String($bytes))
}
$SessionSecretArn = $sessionArn.Trim()
& $AwsCli secretsmanager tag-resource --secret-id $SessionSecretArn `
    --tags "Key=Name,Value=$AppName" --profile default --region $Region | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Nao consegui etiquetar o secret de sessao." }

# Login Microsoft e obrigatorio em producao. A senha compartilhada foi
# desativada para preservar identidade nominal e isolamento por squad.
$EntraSecretArn = ""
if (-not $EntraClientSecret -or -not $EntraTenantId -or -not $EntraClientId -or -not $EntraRedirectUri) {
    throw "Informe EntraTenantId, EntraClientId, EntraClientSecret e EntraRedirectUri. Login por senha esta desativado."
}
if ($EntraRedirectUri -notmatch '^https://[^/]+/auth/callback$') {
    throw "EntraRedirectUri deve ser https://<host>/auth/callback (recebido: $EntraRedirectUri)."
}
$EntraSecretArn = Set-Secret "$AppName/entra-client-secret" $EntraClientSecret

# --- 5. aplicacao ---------------------------------------------------------
Write-Host "[5/5] Publicando a aplicacao..." -ForegroundColor Cyan
& $AwsCli cloudformation deploy `
    --template-file (Join-Path $PSScriptRoot "apprunner.yaml") `
    --stack-name "$AppName-stack" `
    --capabilities CAPABILITY_NAMED_IAM `
    --profile default `
    --region $Region `
    --parameter-overrides `
        "AppName=$AppName" `
        "ImageUri=$ImageUri" `
        "OpenAISecretArn=$OpenAISecretArn" `
        "SessionSecretArn=$SessionSecretArn" `
        "EntraTenantId=$EntraTenantId" `
        "EntraClientId=$EntraClientId" `
        "EntraRedirectUri=$EntraRedirectUri" `
        "EntraClientSecretArn=$EntraSecretArn" `
        "BootstrapAdminEmails=$BootstrapAdminEmails" `
        "BucketName=$AppName-$AccountId" | Out-Null

$ServiceUrl = & $AwsCli cloudformation describe-stacks --stack-name "$AppName-stack" --profile default --region $Region `
    --query "Stacks[0].Outputs[?OutputKey=='ServiceUrl'].OutputValue" --output text

Write-Host ""
Write-Host "Pronto. Envie este endereco para a equipe:" -ForegroundColor Green
Write-Host "  $ServiceUrl" -ForegroundColor Green
Write-Host "A primeira visita usa a conta Microsoft e pede a escolha do squad." -ForegroundColor Green
