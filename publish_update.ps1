<#
Publishes the test build in dist\AutoYouTubeUploader as a GitHub release.

Workflow: edit code, run .\build_exe.ps1 (fast, folder only), test the EXE, repeat.
When satisfied, commit and push your changes, then run this script. It zips the
folder, writes latest.json, tags vX.Y.Z, and creates the public release with both files.

  .\publish_update.ps1 -DryRun   # build the zip and manifest only; no tag or release
  .\publish_update.ps1           # tag, push the tag, and create the GitHub release
#>
param(
    [switch]$DryRun,
    [string]$Notes = ""
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$AppDir = Join-Path $ProjectRoot "dist\AutoYouTubeUploader"
$Executable = Join-Path $AppDir "AutoYouTubeUploader.exe"
$Source = Join-Path $ProjectRoot "auto_yt.py"

if ((Get-Content -LiteralPath $Source -Raw) -notmatch 'APP_VERSION = "([^"]+)"') {
    throw "Could not determine APP_VERSION from auto_yt.py."
}
$Version = $Matches[1]
$Tag = "v$Version"

if (-not (Test-Path -LiteralPath $Executable -PathType Leaf)) {
    throw "Build the test folder first with .\build_exe.ps1"
}
foreach ($secret in @("config.json", "client.json", "token.json")) {
    if (Test-Path -LiteralPath (Join-Path $AppDir $secret)) {
        throw "Remove $secret from $AppDir before publishing. User settings and secrets must not be in the update."
    }
}

$BuildTime = (Get-Item -LiteralPath $Executable).LastWriteTime
foreach ($file in @("auto_yt.py", "client_app.py", "updater.py", "requirements.txt")) {
    $path = Join-Path $ProjectRoot $file
    if ((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).LastWriteTime -gt $BuildTime)) {
        throw "$file changed after the last build. Run .\build_exe.ps1 first."
    }
}

Push-Location $ProjectRoot
try {
    if (git status --porcelain) {
        throw "Working tree has uncommitted changes. Commit and push before publishing."
    }
    if (-not $DryRun -and (git rev-parse -q --verify "refs/tags/$Tag")) {
        throw "Tag $Tag already exists. Bump APP_VERSION in auto_yt.py and rebuild."
    }
    if (-not (git branch --remotes --contains HEAD)) {
        throw "HEAD is not pushed to origin. Push your commits before publishing."
    }
    $Remote = git remote get-url origin
}
finally {
    Pop-Location
}

if ($Remote -notmatch 'github\.com[:/]([^/]+)/([^/]+?)(\.git)?$') {
    throw "Could not determine the GitHub repository from origin: $Remote"
}
$Repo = "$($Matches[1])/$($Matches[2])"
$DownloadUrl = "https://github.com/$Repo/releases/download/$Tag/AutoYouTubeUploader-$Version.zip"

$ZipPath = Join-Path $ProjectRoot "dist\AutoYouTubeUploader-$Version.zip"
$ManifestPath = Join-Path $ProjectRoot "dist\latest.json"
Remove-Item -LiteralPath $ZipPath, $ManifestPath -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($AppDir, $ZipPath)

$Hash = (Get-FileHash -LiteralPath $ZipPath -Algorithm SHA256).Hash.ToLowerInvariant()
$Manifest = @{
    version = $Version
    download_url = $DownloadUrl
    sha256 = $Hash
    notes = $Notes
} | ConvertTo-Json
[System.IO.File]::WriteAllText($ManifestPath, $Manifest, (New-Object System.Text.UTF8Encoding($false)))

Write-Host "Built $ZipPath"
Write-Host "Built $ManifestPath"

if ($DryRun) {
    Write-Host "Dry run: no tag or release created. Download URL would be $DownloadUrl"
    return
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI (gh) is required to create the release. Install it from https://cli.github.com and run 'gh auth login'."
}

Push-Location $ProjectRoot
try {
    git tag -a $Tag -m "Release $Tag"
    if ($LASTEXITCODE -ne 0) { throw "Creating tag $Tag failed." }
    git push origin $Tag
    if ($LASTEXITCODE -ne 0) { throw "Pushing tag $Tag failed." }
    $ReleaseNotes = if ($Notes) { $Notes } else { "Release $Tag" }
    gh release create $Tag $ZipPath $ManifestPath --repo $Repo --verify-tag --title $Tag --notes $ReleaseNotes
    if ($LASTEXITCODE -ne 0) { throw "Creating the GitHub release failed." }
}
finally {
    Pop-Location
}

Write-Host "Published $Tag to $Repo. Clients pick it up on their next update check."
