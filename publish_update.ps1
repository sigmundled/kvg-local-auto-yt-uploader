param(
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [Parameter(Mandatory = $true)]
    [string]$DownloadUrl,
    [string]$Notes = ""
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$AppDir = Join-Path $ProjectRoot "dist\AutoYouTubeUploader"
$Executable = Join-Path $AppDir "AutoYouTubeUploader.exe"
$Source = Join-Path $ProjectRoot "auto_yt.py"
if (-not (Test-Path -LiteralPath $Executable -PathType Leaf)) {
    throw "Build the client first with .\build_exe.ps1"
}
foreach ($secret in @("config.json", "client.json", "token.json")) {
    if (Test-Path -LiteralPath (Join-Path $AppDir $secret)) {
        throw "Remove $secret from $AppDir before publishing. User settings and secrets must not be in the update."
    }
}
if ((Get-Content -LiteralPath $Source -Raw) -notmatch 'APP_VERSION = "([^"]+)"') {
    throw "Could not determine APP_VERSION from auto_yt.py."
}
if ($Matches[1] -ne $Version) {
    throw "Requested version $Version does not match APP_VERSION $($Matches[1]). Rebuild after updating APP_VERSION."
}
if (-not $DownloadUrl.StartsWith("https://", [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "DownloadUrl must use HTTPS."
}

$ZipPath = Join-Path $ProjectRoot "dist\AutoYouTubeUploader-$Version.zip"
Remove-Item -LiteralPath $ZipPath -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $AppDir "*") -DestinationPath $ZipPath

$Hash = (Get-FileHash -LiteralPath $ZipPath -Algorithm SHA256).Hash.ToLowerInvariant()
$Manifest = @{
    version = $Version
    download_url = $DownloadUrl
    sha256 = $Hash
    notes = $Notes
} | ConvertTo-Json
$ManifestPath = Join-Path $ProjectRoot "dist\latest.json"
$Manifest | Set-Content -LiteralPath $ManifestPath -Encoding UTF8

Write-Host "Publish these files:"
Write-Host "  $ZipPath  (upload to the URL given as -DownloadUrl)"
Write-Host "  $ManifestPath"
Write-Host "Ensure latest.json's public URL is configured as the client update manifest URL."
