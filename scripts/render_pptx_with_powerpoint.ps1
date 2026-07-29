param(
    [Parameter(Mandatory = $true)]
    [string]$InputPptx,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [int]$Width = 1920
)

$ErrorActionPreference = "Stop"
$inputPath = (Resolve-Path -LiteralPath $InputPptx).Path
$outputPath = [System.IO.Path]::GetFullPath($OutputDir)
[System.IO.Directory]::CreateDirectory($outputPath) | Out-Null
$statusPath = Join-Path $outputPath ".render-status.txt"
[System.IO.File]::WriteAllText($statusPath, "iniciando")

$presentation = $null
$powerPoint = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.DisplayAlerts = 1
    $powerPoint.AutomationSecurity = 3
    [System.IO.File]::WriteAllText($statusPath, "abrindo")
    $presentation = $powerPoint.Presentations.Open($inputPath, $true, $false, $false)
    [System.IO.File]::WriteAllText($statusPath, "aberto")
    $slideWidth = [double]$presentation.PageSetup.SlideWidth
    $slideHeight = [double]$presentation.PageSetup.SlideHeight
    $height = [Math]::Max(1, [Math]::Round($Width * $slideHeight / $slideWidth))
    $files = @()

    foreach ($slide in $presentation.Slides) {
        $fileName = "slide-$($slide.SlideIndex).png"
        $filePath = Join-Path $outputPath $fileName
        [System.IO.File]::WriteAllText($statusPath, "exportando $($slide.SlideIndex)")
        $slide.Export($filePath, "PNG", $Width, $height)
        $files += $fileName
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($slide)
    }
    [System.IO.File]::WriteAllText($statusPath, "concluido")

    [ordered]@{
        input = $inputPath
        output = $outputPath
        slides = $files.Count
        width = $Width
        height = $height
        files = $files
    } | ConvertTo-Json -Compress
}
finally {
    if ($null -ne $presentation) {
        $presentation.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($null -ne $powerPoint) {
        $powerPoint.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
