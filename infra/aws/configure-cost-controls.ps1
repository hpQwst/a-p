<#
.SYNOPSIS
    Ativa rateio por Name e cria um budget mensal somente para squad4e5-auto-ppt.
#>
param(
    [string]$AppName = "squad4e5-auto-ppt",
    [string]$AccountId = "134164930693",
    [decimal]$MonthlyLimitUsd = 20,
    [string]$AlertEmail = "hugo.rocha@qwst.co",
    [string]$Region = "us-east-1"
)

$ErrorActionPreference = "Stop"
$AwsCli = "C:\Program Files\Amazon\AWSCLIV2\aws.exe"
if (-not $AppName.StartsWith("squad4") -and -not $AppName.StartsWith("squad5")) {
    throw "AppName precisa comecar com squad4/squad5."
}

$actualAccount = (& $AwsCli sts get-caller-identity --profile default --region $Region --query Account --output text).Trim()
if ($actualAccount -ne $AccountId) {
    throw "Conta AWS inesperada: $actualAccount."
}

# Em AWS Organizations, somente a conta pagadora pode ativar tags de rateio.
# A falha aqui nao cria um budget amplo: o filtro continua estritamente no Name
# deste app e o script avisa que a conta pagadora precisa concluir a ativacao.
$previousErrorPreference = $ErrorActionPreference
$ErrorActionPreference = "Continue"
$tagActivation = & $AwsCli ce update-cost-allocation-tags-status `
    --profile default `
    --region $Region `
    --cost-allocation-tags-status "TagKey=Name,Status=Active" 2>&1
$tagActivationExit = $LASTEXITCODE
$ErrorActionPreference = $previousErrorPreference
if ($tagActivationExit -ne 0) {
    Write-Warning "Nao foi possivel ativar a tag Name nesta conta vinculada. A conta pagadora precisa ativa-la no Billing."
}

$budgetName = "$AppName-monthly"
$budget = @{
    BudgetName = $budgetName
    BudgetLimit = @{ Amount = [string]$MonthlyLimitUsd; Unit = "USD" }
    CostFilters = @{ TagKeyValue = @("user:Name`$$AppName") }
    TimeUnit = "MONTHLY"
    BudgetType = "COST"
} | ConvertTo-Json -Depth 6 -Compress
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$budgetPath = Join-Path ([System.IO.Path]::GetTempPath()) "$budgetName-$([guid]::NewGuid().ToString('N')).json"
[System.IO.File]::WriteAllText($budgetPath, $budget, $utf8NoBom)

$ErrorActionPreference = "Continue"
$existing = & $AwsCli budgets describe-budget `
    --account-id $AccountId `
    --budget-name $budgetName `
    --profile default `
    --region $Region `
    --query "Budget.BudgetName" `
    --output text 2>$null
$describeExit = $LASTEXITCODE
$ErrorActionPreference = $previousErrorPreference

try {
    if ($describeExit -eq 0 -and $existing -eq $budgetName) {
        & $AwsCli budgets update-budget `
            --account-id $AccountId `
            --new-budget "file://$budgetPath" `
            --profile default `
            --region $Region | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Falha ao atualizar budget $budgetName." }
    } else {
        $notifications = @(
            @{
                Notification = @{
                    NotificationType = "ACTUAL"
                    ComparisonOperator = "GREATER_THAN"
                    Threshold = 80
                    ThresholdType = "PERCENTAGE"
                }
                Subscribers = @(@{ SubscriptionType = "EMAIL"; Address = $AlertEmail })
            },
            @{
                Notification = @{
                    NotificationType = "FORECASTED"
                    ComparisonOperator = "GREATER_THAN"
                    Threshold = 100
                    ThresholdType = "PERCENTAGE"
                }
                Subscribers = @(@{ SubscriptionType = "EMAIL"; Address = $AlertEmail })
            }
        ) | ConvertTo-Json -Depth 7 -Compress
        $notificationsPath = Join-Path ([System.IO.Path]::GetTempPath()) "$budgetName-notifications-$([guid]::NewGuid().ToString('N')).json"
        [System.IO.File]::WriteAllText($notificationsPath, $notifications, $utf8NoBom)
        & $AwsCli budgets create-budget `
            --account-id $AccountId `
            --budget "file://$budgetPath" `
            --notifications-with-subscribers "file://$notificationsPath" `
            --profile default `
            --region $Region | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Falha ao criar budget $budgetName." }
    }
}
finally {
    Remove-Item -LiteralPath $budgetPath -Force -ErrorAction SilentlyContinue
    if ($notificationsPath) {
        Remove-Item -LiteralPath $notificationsPath -Force -ErrorAction SilentlyContinue
    }
}

& $AwsCli budgets describe-budget `
    --account-id $AccountId `
    --budget-name $budgetName `
    --profile default `
    --region $Region `
    --query "Budget.{Name:BudgetName,Limit:BudgetLimit,Filters:CostFilters}" `
    --output json
