# PowerShell script to deploy main-js static files to gh-pages/development or gh-pages/release

param(
    [string]$TargetEnv = "development"  # default to development, can be set to 'release'
)

$ErrorActionPreference = 'Stop'

$sourceRoot = Join-Path $PSScriptRoot ".."  # relative to scripts folder
$targetRoot = Join-Path $PSScriptRoot "../../gh-pages/$TargetEnv"

# Ensure target directories exist
$srcTarget = Join-Path $targetRoot "src"
$assetsTarget = Join-Path $targetRoot "assets"

if (!(Test-Path $srcTarget)) { New-Item -ItemType Directory -Path $srcTarget -Force | Out-Null }
if (!(Test-Path $assetsTarget)) { New-Item -ItemType Directory -Path $assetsTarget -Force | Out-Null }

# Extract and update build number from README.md
$readmePath = Join-Path $sourceRoot 'README.md'
$readmeContent = Get-Content $readmePath -Raw
$todayDate = Get-Date -Format "MMMM d, yyyy"
$buildNumber = 1  # Default to 1 if no build number found

if ($readmeContent -match '\*\*Build #(\d+)\*\*') {
    $buildNumber = [int]$Matches[1] + 1
}
Write-Host "Build number: $buildNumber"

# Update README.md with new build number
$newVersionLine = "**Build #$buildNumber** ($todayDate)"

if ($readmeContent -match '## Current Version\s+\*\*Build #\d+\*\*') {
    $readmeContent = $readmeContent -replace '## Current Version\s+\*\*Build #\d+\*\* \([^)]+\)', "## Current Version`n`n$newVersionLine"
} else {
    $readmeContent = $readmeContent -replace '# Shape Master JS\s+', "# Shape Master JS`n`n## Current Version`n`n$newVersionLine`n`n"
}
Set-Content -Path $readmePath -Value $readmeContent

# Copy src and assets (overwrite existing)
Write-Host "Copying src..."
Copy-Item -Path (Join-Path $sourceRoot 'src\*') -Destination $srcTarget -Recurse -Force
Write-Host "Copying assets..."
Copy-Item -Path (Join-Path $sourceRoot 'assets\*') -Destination $assetsTarget -Recurse -Force

# Copy README.md to the target environment
Write-Host "Copying README.md..."
Copy-Item -Path (Join-Path $sourceRoot 'README.md') -Destination $targetRoot -Force

# Copy manifest.xml for development sideloading
if ($TargetEnv -eq "development") {
    Write-Host "Copying manifest.xml..."
    Copy-Item -Path (Join-Path $sourceRoot 'manifest.xml') -Destination $targetRoot -Force
}

# Commit and push changes
Set-Location $targetRoot
$commitMsg = "Deploy build $buildNumber to $TargetEnv"
git add .
git commit -m $commitMsg
# Optionally push (uncomment next line if you want auto-push)
git push
Set-Location $PSScriptRoot

Write-Host "Deployment to gh-pages/$TargetEnv complete. Build number: $buildNumber"
