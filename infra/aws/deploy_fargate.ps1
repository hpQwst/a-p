param(
    [string]$AppName = "qwst-auto-ppt",
    [string]$Region = "",
    [string]$AllowedCidr = "0.0.0.0/0",
    [int]$DesiredCount = 1,
    [string]$Cpu = "1024",
    [string]$Memory = "3072"
)

$ErrorActionPreference = "Stop"

function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message"
}

function Assert-AwsOk() {
    if ($LASTEXITCODE -ne 0) {
        throw "AWS CLI command failed."
    }
}

function Save-JsonFile($Object, [string]$Path) {
    $Object | ConvertTo-Json -Depth 30 | Set-Content -Encoding UTF8 -Path $Path
}

function Get-RoleArn([string]$RoleName) {
    $arn = aws iam get-role --role-name $RoleName --query "Role.Arn" --output text 2>$null
    if ($LASTEXITCODE -eq 0 -and $arn -and $arn -ne "None") {
        return $arn.Trim()
    }
    return $null
}

function Ensure-Role([string]$RoleName, [string]$ServicePrincipal) {
    $existing = Get-RoleArn $RoleName
    if ($existing) {
        return $existing
    }

    $trustPath = Join-Path $TempDir "$RoleName-trust.json"
    Save-JsonFile @{
        Version = "2012-10-17"
        Statement = @(
            @{
                Effect = "Allow"
                Principal = @{ Service = $ServicePrincipal }
                Action = "sts:AssumeRole"
            }
        )
    } $trustPath

    aws iam create-role --role-name $RoleName --assume-role-policy-document "file://$trustPath" | Out-Null
    Assert-AwsOk
    return (Get-RoleArn $RoleName)
}

function Put-InlinePolicy([string]$RoleName, [string]$PolicyName, $PolicyObject) {
    $policyPath = Join-Path $TempDir "$RoleName-$PolicyName.json"
    Save-JsonFile $PolicyObject $policyPath
    aws iam put-role-policy --role-name $RoleName --policy-name $PolicyName --policy-document "file://$policyPath" | Out-Null
    Assert-AwsOk
}

function Get-EnvValue([string]$Path, [string]$Name) {
    if (-not (Test-Path $Path)) {
        return ""
    }
    foreach ($line in Get-Content -Path $Path) {
        if ($line -match "^\s*$Name\s*=\s*(.+?)\s*$") {
            return $Matches[1].Trim().Trim('"').Trim("'")
        }
    }
    return ""
}

$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$TempDir = Join-Path ([System.IO.Path]::GetTempPath()) "$AppName-deploy"
if (Test-Path $TempDir) {
    $resolvedTemp = (Resolve-Path $TempDir).Path
    if (-not $resolvedTemp.StartsWith([System.IO.Path]::GetTempPath(), [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean unexpected temp path: $resolvedTemp"
    }
    Remove-Item -LiteralPath $resolvedTemp -Recurse -Force
}
New-Item -ItemType Directory -Path $TempDir | Out-Null

if (-not $Region) {
    $Region = (aws configure get region).Trim()
    if (-not $Region) {
        $Region = "sa-east-1"
    }
}

$SafeAppName = ($AppName.ToLower() -replace "[^a-z0-9-]", "-").Trim("-")
$ImageTag = (Get-Date -Format "yyyyMMddHHmmss")

Write-Step "Reading AWS account"
$AccountId = (aws sts get-caller-identity --query Account --output text --region $Region).Trim()
Assert-AwsOk

$EcrRepo = $SafeAppName
$ClusterName = $SafeAppName
$ServiceName = "$SafeAppName-service"
$BuildProjectName = "$SafeAppName-build"
$BucketName = "$SafeAppName-$AccountId-$($Region.ToLower())"
$SourceKey = "source/$ImageTag.zip"
$LogGroup = "/ecs/$SafeAppName"
$OpenAiSecretName = "$SafeAppName/openai-api-key"
$OpenAiModel = Get-EnvValue (Join-Path $RepoRoot ".env") "OPENAI_MODEL"
if (-not $OpenAiModel) {
    $OpenAiModel = "gpt-5.5"
}

Write-Step "Ensuring S3 bucket $BucketName"
aws s3api head-bucket --bucket $BucketName --region $Region 2>$null
if ($LASTEXITCODE -ne 0) {
    if ($Region -eq "us-east-1") {
        aws s3api create-bucket --bucket $BucketName --region $Region | Out-Null
    } else {
        aws s3api create-bucket --bucket $BucketName --region $Region --create-bucket-configuration "LocationConstraint=$Region" | Out-Null
    }
    Assert-AwsOk
}
aws s3api put-public-access-block --bucket $BucketName --public-access-block-configuration "BlockPublicAcls=true,IgnorePublicAcls=true,BlockPublicPolicy=true,RestrictPublicBuckets=true" --region $Region | Out-Null
Assert-AwsOk
aws s3api put-bucket-encryption --bucket $BucketName --server-side-encryption-configuration '{"Rules":[{"ApplyServerSideEncryptionByDefault":{"SSEAlgorithm":"AES256"}}]}' --region $Region | Out-Null
Assert-AwsOk

Write-Step "Ensuring ECR repository $EcrRepo"
aws ecr describe-repositories --repository-names $EcrRepo --region $Region 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    aws ecr create-repository --repository-name $EcrRepo --image-scanning-configuration scanOnPush=true --region $Region | Out-Null
    Assert-AwsOk
}
$RepoArn = (aws ecr describe-repositories --repository-names $EcrRepo --query "repositories[0].repositoryArn" --output text --region $Region).Trim()
$ImageUri = "$AccountId.dkr.ecr.$Region.amazonaws.com/${EcrRepo}:$ImageTag"

Write-Step "Syncing OpenAI key to Secrets Manager if .env is present"
$OpenAiKey = Get-EnvValue (Join-Path $RepoRoot ".env") "OPENAI_API_KEY"
$SecretArn = ""
if ($OpenAiKey -and $OpenAiKey -notmatch "^coloque_") {
    $secretFile = Join-Path $TempDir "openai_secret.txt"
    Set-Content -NoNewline -Encoding UTF8 -Path $secretFile -Value $OpenAiKey
    aws secretsmanager describe-secret --secret-id $OpenAiSecretName --region $Region 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) {
        aws secretsmanager update-secret --secret-id $OpenAiSecretName --secret-string "file://$secretFile" --region $Region | Out-Null
    } else {
        aws secretsmanager create-secret --name $OpenAiSecretName --secret-string "file://$secretFile" --region $Region | Out-Null
    }
    Assert-AwsOk
    Remove-Item -LiteralPath $secretFile -Force
    $SecretArn = (aws secretsmanager describe-secret --secret-id $OpenAiSecretName --query ARN --output text --region $Region).Trim()
} else {
    Write-Warning "OPENAI_API_KEY not found in .env. The app will run without AI mapping until the secret is configured."
}

Write-Step "Ensuring IAM roles"
$ExecutionRoleName = "$SafeAppName-ecs-exec-role"
$TaskRoleName = "$SafeAppName-task-role"
$CodeBuildRoleName = "$SafeAppName-codebuild-role"
$ExecutionRoleArn = Ensure-Role $ExecutionRoleName "ecs-tasks.amazonaws.com"
$TaskRoleArn = Ensure-Role $TaskRoleName "ecs-tasks.amazonaws.com"
$CodeBuildRoleArn = Ensure-Role $CodeBuildRoleName "codebuild.amazonaws.com"

aws iam attach-role-policy --role-name $ExecutionRoleName --policy-arn "arn:aws:iam::aws:policy/service-role/AmazonECSTaskExecutionRolePolicy" | Out-Null
Assert-AwsOk

if ($SecretArn) {
    Put-InlinePolicy $ExecutionRoleName "read-openai-secret" @{
        Version = "2012-10-17"
        Statement = @(
            @{
                Effect = "Allow"
                Action = @("secretsmanager:GetSecretValue")
                Resource = $SecretArn
            }
        )
    }
}

Put-InlinePolicy $TaskRoleName "project-storage" @{
    Version = "2012-10-17"
    Statement = @(
        @{
            Effect = "Allow"
            Action = @("s3:ListBucket")
            Resource = "arn:aws:s3:::$BucketName"
        },
        @{
            Effect = "Allow"
            Action = @("s3:GetObject", "s3:PutObject", "s3:DeleteObject")
            Resource = "arn:aws:s3:::$BucketName/auto-ppt/*"
        }
    )
}

Put-InlinePolicy $CodeBuildRoleName "build-and-push" @{
    Version = "2012-10-17"
    Statement = @(
        @{
            Effect = "Allow"
            Action = @("ecr:GetAuthorizationToken")
            Resource = "*"
        },
        @{
            Effect = "Allow"
            Action = @(
                "ecr:BatchCheckLayerAvailability",
                "ecr:CompleteLayerUpload",
                "ecr:DescribeRepositories",
                "ecr:InitiateLayerUpload",
                "ecr:PutImage",
                "ecr:UploadLayerPart"
            )
            Resource = $RepoArn
        },
        @{
            Effect = "Allow"
            Action = @("logs:CreateLogGroup", "logs:CreateLogStream", "logs:PutLogEvents")
            Resource = "*"
        },
        @{
            Effect = "Allow"
            Action = @("s3:GetObject", "s3:GetObjectVersion", "s3:PutObject")
            Resource = @("arn:aws:s3:::$BucketName/source/*", "arn:aws:s3:::$BucketName/codebuild/*")
        }
    )
}

Start-Sleep -Seconds 12

Write-Step "Packaging source for CodeBuild"
$StageDir = Join-Path $TempDir "source"
New-Item -ItemType Directory -Path $StageDir | Out-Null
Copy-Item -LiteralPath (Join-Path $RepoRoot "app.py") -Destination $StageDir
Copy-Item -LiteralPath (Join-Path $RepoRoot "requirements.txt") -Destination $StageDir
Copy-Item -LiteralPath (Join-Path $RepoRoot "Dockerfile") -Destination $StageDir
Copy-Item -LiteralPath (Join-Path $RepoRoot "buildspec.yml") -Destination $StageDir
Copy-Item -LiteralPath (Join-Path $RepoRoot "ppt_automator") -Destination (Join-Path $StageDir "ppt_automator") -Recurse
$ZipPath = Join-Path $TempDir "source.zip"
Compress-Archive -Path (Join-Path $StageDir "*") -DestinationPath $ZipPath -Force
aws s3 cp $ZipPath "s3://$BucketName/$SourceKey" --region $Region | Out-Null
Assert-AwsOk

Write-Step "Creating or updating CodeBuild project"
$CodeBuildConfigPath = Join-Path $TempDir "codebuild-project.json"
Save-JsonFile @{
    name = $BuildProjectName
    serviceRole = $CodeBuildRoleArn
    artifacts = @{ type = "NO_ARTIFACTS" }
    source = @{
        type = "S3"
        location = "$BucketName/$SourceKey"
        buildspec = "buildspec.yml"
    }
    environment = @{
        type = "LINUX_CONTAINER"
        image = "aws/codebuild/standard:7.0"
        computeType = "BUILD_GENERAL1_MEDIUM"
        privilegedMode = $true
        environmentVariables = @(
            @{ name = "IMAGE_REPO_NAME"; value = $EcrRepo; type = "PLAINTEXT" },
            @{ name = "IMAGE_TAG"; value = $ImageTag; type = "PLAINTEXT" }
        )
    }
} $CodeBuildConfigPath

aws codebuild batch-get-projects --names $BuildProjectName --query "projects[0].name" --output text --region $Region 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) {
    aws codebuild update-project --cli-input-json "file://$CodeBuildConfigPath" --region $Region | Out-Null
} else {
    aws codebuild create-project --cli-input-json "file://$CodeBuildConfigPath" --region $Region | Out-Null
}
Assert-AwsOk

Write-Step "Starting cloud Docker build"
$BuildId = (aws codebuild start-build --project-name $BuildProjectName --region $Region --query "build.id" --output text).Trim()
Assert-AwsOk
do {
    Start-Sleep -Seconds 20
    $BuildStatus = (aws codebuild batch-get-builds --ids $BuildId --region $Region --query "builds[0].buildStatus" --output text).Trim()
    Write-Host "CodeBuild status: $BuildStatus"
} while ($BuildStatus -in @("IN_PROGRESS", "QUEUED"))
if ($BuildStatus -ne "SUCCEEDED") {
    throw "CodeBuild failed with status $BuildStatus. Check the CodeBuild logs for $BuildId."
}

Write-Step "Ensuring ECS cluster and networking"
aws logs create-log-group --log-group-name $LogGroup --region $Region 2>$null | Out-Null
aws ecs create-cluster --cluster-name $ClusterName --region $Region 2>$null | Out-Null

$VpcId = (aws ec2 describe-vpcs --filters "Name=isDefault,Values=true" --query "Vpcs[0].VpcId" --output text --region $Region).Trim()
Assert-AwsOk
if (-not $VpcId -or $VpcId -eq "None") {
    throw "No default VPC found in region $Region. Create/provide a VPC before running this script."
}
$SubnetsText = (aws ec2 describe-subnets --filters "Name=vpc-id,Values=$VpcId" --query "Subnets[].SubnetId" --output text --region $Region).Trim()
Assert-AwsOk
$SubnetIds = $SubnetsText -split "\s+" | Where-Object { $_ }
if ($SubnetIds.Count -eq 0) {
    throw "No subnets found in VPC $VpcId."
}

$SecurityGroupName = "$SafeAppName-sg"
$SgId = (aws ec2 describe-security-groups --filters "Name=vpc-id,Values=$VpcId" "Name=group-name,Values=$SecurityGroupName" --query "SecurityGroups[0].GroupId" --output text --region $Region 2>$null).Trim()
if (-not $SgId -or $SgId -eq "None") {
    $SgId = (aws ec2 create-security-group --group-name $SecurityGroupName --description "Streamlit access for $SafeAppName" --vpc-id $VpcId --query "GroupId" --output text --region $Region).Trim()
    Assert-AwsOk
}
aws ec2 authorize-security-group-ingress --group-id $SgId --protocol tcp --port 8501 --cidr $AllowedCidr --region $Region 2>$null | Out-Null

Write-Step "Registering ECS task definition"
$Environment = @(
    @{ name = "AUTO_PPT_STORAGE_BACKEND"; value = "s3" },
    @{ name = "AUTO_PPT_S3_BUCKET"; value = $BucketName },
    @{ name = "AUTO_PPT_S3_PREFIX"; value = "auto-ppt" },
    @{ name = "OPENAI_MODEL"; value = $OpenAiModel }
)
$Secrets = @()
if ($SecretArn) {
    $Secrets += @{ name = "OPENAI_API_KEY"; valueFrom = $SecretArn }
}
$ContainerDefinition = @{
    name = "auto-ppt"
    image = $ImageUri
    essential = $true
    portMappings = @(
        @{ containerPort = 8501; hostPort = 8501; protocol = "tcp" }
    )
    environment = $Environment
    logConfiguration = @{
        logDriver = "awslogs"
        options = @{
            "awslogs-group" = $LogGroup
            "awslogs-region" = $Region
            "awslogs-stream-prefix" = "ecs"
        }
    }
}
if ($Secrets.Count -gt 0) {
    $ContainerDefinition.secrets = $Secrets
}
$TaskDefPath = Join-Path $TempDir "task-definition.json"
Save-JsonFile @{
    family = $SafeAppName
    requiresCompatibilities = @("FARGATE")
    networkMode = "awsvpc"
    cpu = "$Cpu"
    memory = "$Memory"
    executionRoleArn = $ExecutionRoleArn
    taskRoleArn = $TaskRoleArn
    runtimePlatform = @{
        operatingSystemFamily = "LINUX"
        cpuArchitecture = "X86_64"
    }
    containerDefinitions = @($ContainerDefinition)
} $TaskDefPath
$TaskDefinitionArn = (aws ecs register-task-definition --cli-input-json "file://$TaskDefPath" --query "taskDefinition.taskDefinitionArn" --output text --region $Region).Trim()
Assert-AwsOk

Write-Step "Creating or updating Fargate service"
$SubnetConfig = ($SubnetIds -join ",")
$NetworkConfig = "awsvpcConfiguration={subnets=[$SubnetConfig],securityGroups=[$SgId],assignPublicIp=ENABLED}"
$ExistingService = (aws ecs describe-services --cluster $ClusterName --services $ServiceName --query "services[0].status" --output text --region $Region 2>$null).Trim()
if ($ExistingService -eq "ACTIVE") {
    aws ecs update-service --cluster $ClusterName --service $ServiceName --task-definition $TaskDefinitionArn --desired-count $DesiredCount --region $Region | Out-Null
} else {
    aws ecs create-service `
        --cluster $ClusterName `
        --service-name $ServiceName `
        --task-definition $TaskDefinitionArn `
        --desired-count $DesiredCount `
        --launch-type FARGATE `
        --network-configuration $NetworkConfig `
        --region $Region | Out-Null
}
Assert-AwsOk

aws ecs wait services-stable --cluster $ClusterName --services $ServiceName --region $Region
Assert-AwsOk

$TaskArn = (aws ecs list-tasks --cluster $ClusterName --service-name $ServiceName --desired-status RUNNING --query "taskArns[0]" --output text --region $Region).Trim()
if ($TaskArn -and $TaskArn -ne "None") {
    $EniId = (aws ecs describe-tasks --cluster $ClusterName --tasks $TaskArn --query "tasks[0].attachments[0].details[?name=='networkInterfaceId'].value | [0]" --output text --region $Region).Trim()
    $PublicIp = (aws ec2 describe-network-interfaces --network-interface-ids $EniId --query "NetworkInterfaces[0].Association.PublicIp" --output text --region $Region).Trim()
    Write-Host ""
    Write-Host "Deploy complete."
    Write-Host "URL: http://$PublicIp`:8501"
    Write-Host "S3 bucket: $BucketName"
    Write-Host "ECR image: $ImageUri"
    Write-Host "Allowed CIDR: $AllowedCidr"
} else {
    Write-Warning "Deploy finished, but no running task was found yet."
}
