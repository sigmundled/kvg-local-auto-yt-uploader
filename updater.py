"""Minimal signed-by-transport updater for the packaged client."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import requests


def _version_parts(value: str) -> tuple[int, ...]:
    parts = []
    for part in value.strip().split("."):
        if not part.isdigit():
            raise ValueError(f"Version is not numeric dotted format: {value}")
        parts.append(int(part))
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts)


def newer_version(current: str, available: str) -> bool:
    return _version_parts(available) > _version_parts(current)


def check_update(manifest_url: str, current_version: str) -> dict[str, str] | None:
    if not manifest_url:
        return None
    if not manifest_url.lower().startswith("https://"):
        raise ValueError("The update manifest URL must use HTTPS.")
    with requests.get(manifest_url, timeout=(15, 60)) as response:
        response.raise_for_status()
        manifest = response.json()
    if not isinstance(manifest, dict):
        raise ValueError("The update manifest must be a JSON object.")
    required = ("version", "download_url", "sha256")
    if any(not isinstance(manifest.get(key), str) or not manifest[key].strip() for key in required):
        raise ValueError("The update manifest requires version, download_url, and sha256.")
    download_url = manifest["download_url"].strip()
    if not download_url.lower().startswith("https://"):
        raise ValueError("The update download URL must use HTTPS.")
    if not all(character in "0123456789abcdefABCDEF" for character in manifest["sha256"].strip()):
        raise ValueError("The update manifest SHA-256 is not hexadecimal.")
    if len(manifest["sha256"].strip()) != 64:
        raise ValueError("The update manifest SHA-256 must contain 64 hexadecimal characters.")
    if not newer_version(current_version, manifest["version"].strip()):
        return None
    return {
        "version": manifest["version"].strip(),
        "download_url": download_url,
        "sha256": manifest["sha256"].strip().lower(),
        "notes": str(manifest.get("notes") or ""),
    }


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def download_update(update: dict[str, str]) -> Path:
    """Download and verify an update ZIP, returning the extracted application folder."""
    target_dir = (
        Path(os.environ.get("LOCALAPPDATA", Path.home())) / "AutoYouTubeUploader" / "updates"
        if os.name == "nt"
        else Path.home() / "AutoYouTubeUploader" / "updates"
    )
    target_dir.mkdir(parents=True, exist_ok=True)
    version = update["version"]
    downloaded = target_dir / f"AutoYouTubeUploader-{version}.zip"
    temporary = target_dir / f".AutoYouTubeUploader-{version}.download"
    with requests.get(update["download_url"], stream=True, timeout=(15, 120)) as response:
        response.raise_for_status()
        with temporary.open("wb") as stream:
            for block in response.iter_content(chunk_size=1024 * 1024):
                if block:
                    stream.write(block)
    if _sha256(temporary) != update["sha256"]:
        temporary.unlink(missing_ok=True)
        raise ValueError("The downloaded update failed its SHA-256 verification.")
    if not zipfile.is_zipfile(temporary):
        temporary.unlink(missing_ok=True)
        raise ValueError("The downloaded update is not a ZIP archive.")
    temporary.replace(downloaded)

    staged = target_dir / f"AutoYouTubeUploader-{version}"
    shutil.rmtree(staged, ignore_errors=True)
    with zipfile.ZipFile(downloaded) as archive:
        if "AutoYouTubeUploader.exe" not in archive.namelist():
            raise ValueError("The downloaded update does not contain AutoYouTubeUploader.exe.")
        archive.extractall(staged)
    downloaded.unlink(missing_ok=True)
    return staged


def install_update(staged: Path) -> None:
    """Replace the installed application folder once this process exits, then relaunch it.

    Only the application files are replaced. config.json, which lives beside the EXE,
    is not part of the update archive and is left untouched.
    """
    if not getattr(sys, "frozen", False):
        raise RuntimeError("Automatic replacement is available only in the packaged EXE.")
    executable = Path(sys.executable).resolve()
    app_dir = executable.parent
    script = Path(tempfile.gettempdir()) / f"auto-yt-update-{os.getpid()}.ps1"

    def ps_quote(path: Path) -> str:
        return "'" + str(path).replace("'", "''") + "'"

    script.write_text(
        (
            f"$process = Get-Process -Id {os.getpid()} -ErrorAction SilentlyContinue\n"
            "if ($process) { $process.WaitForExit() }\n"
            f"Remove-Item -LiteralPath {ps_quote(app_dir / '_internal')} -Recurse -Force -ErrorAction SilentlyContinue\n"
            f"Copy-Item -Path {ps_quote(staged / '*')} -Destination {ps_quote(app_dir)} -Recurse -Force\n"
            f"Start-Process -FilePath {ps_quote(executable)}\n"
            f"Remove-Item -LiteralPath {ps_quote(staged)} -Recurse -Force -ErrorAction SilentlyContinue\n"
            f"Remove-Item -LiteralPath {ps_quote(script)} -Force -ErrorAction SilentlyContinue\n"
        ),
        encoding="utf-8",
    )
    subprocess.Popen(
        [
            "powershell.exe",
            "-NoProfile",
            "-WindowStyle",
            "Hidden",
            "-File",
            str(script),
        ],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
    )
