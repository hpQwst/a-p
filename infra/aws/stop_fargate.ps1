param(
    [string]$AppName = "qwst-auto-ppt",
    [string]$Region = ""
)

$ErrorActionPreference = "Stop"

if (-not $Region) {
    $Region = (aws configure get region).Trim()
    if (-not $Region) {
        $Region = "sa-east-1"
    }
}

$SafeAppName = ($AppName.ToLower() -replace "[^a-z0-9-]", "-").Trim("-")
$ClusterName = $SafeAppName
$ServiceName = "$SafeAppName-service"

aws ecs update-service `
    --cluster $ClusterName `
    --service $ServiceName `
    --desired-count 0 `
    --region $Region | Out-Null

aws ecs wait services-stable --cluster $ClusterName --services $ServiceName --region $Region

Write-Host "Fargate service stopped: $ServiceName"
