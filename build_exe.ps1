$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $ProjectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $Python -PathType Leaf)) {
    throw "Virtual environment not found. Run: py -3.12 -m venv .venv"
}

& $Python -m pip install -r (Join-Path $ProjectRoot "requirements-build.txt")
if ($LASTEXITCODE -ne 0) {
    throw "Build dependency installation failed."
}

Push-Location $ProjectRoot
try {
    & $Python -m PyInstaller `
        --noconfirm `
        --clean `
        --onedir `
        --windowed `
        --name "AutoYouTubeUploader" `
        --collect-all googleapiclient `
        --collect-all google_auth_oauthlib `
        --collect-all msal `
        --collect-all customtkinter `
        --collect-all pystray `
        --add-data "CHANGELOG.md;." `
        "auto_yt.py"
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller build failed."
    }
}
finally {
    Pop-Location
}

$Executable = Join-Path $ProjectRoot "dist\AutoYouTubeUploader\AutoYouTubeUploader.exe"
if (-not (Test-Path -LiteralPath $Executable -PathType Leaf)) {
    throw "Expected executable was not created: $Executable"
}

Write-Host "Built $Executable"

$SignPfx = $env:AUTOYT_SIGN_PFX
if ($SignPfx) {
    $SignTool = Get-ChildItem -Path "${env:ProgramFiles(x86)}\Windows Kits\10\bin" -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match '\\x64\\signtool\.exe$' } |
        Sort-Object FullName -Descending |
        Select-Object -First 1
    if (-not $SignTool) {
        throw "signtool.exe was not found. Install the Windows SDK to sign the EXE."
    }
    $TimestampUrl = if ($env:AUTOYT_SIGN_TIMESTAMP_URL) { $env:AUTOYT_SIGN_TIMESTAMP_URL } else { "http://timestamp.digicert.com" }
    $SignArgs = @("sign", "/fd", "SHA256", "/f", $SignPfx, "/tr", $TimestampUrl, "/td", "SHA256")
    if ($env:AUTOYT_SIGN_PFX_PASSWORD) {
        $SignArgs += @("/p", $env:AUTOYT_SIGN_PFX_PASSWORD)
    }
    & $SignTool.FullName @SignArgs $Executable
    if ($LASTEXITCODE -ne 0) {
        throw "Code signing failed."
    }
    Write-Host "Signed $Executable"
}
else {
    Write-Host "AUTOYT_SIGN_PFX is not set, so the EXE is unsigned."
}
