# Local YouTube uploader

This script reads a video request from Excel, downloads the submitted file, verifies the completed download, and only then asks for permission to upload. Uploads are **private** by default. It does not use Power Automate.

It maps the screenshot's `Id`, `Title`, `Description`, `Tags`, `Playlist`, `Language`, `Video File`, `on SharePoint / OneDrive`, and `YouTube Link` fields (header matching ignores spaces and punctuation). Existing YouTube links and local upload records are skipped. After a successful upload, the script writes the YouTube URL into the local workbook's `YouTube Link` cell.

## Credentials

`client.json` is a placeholder for a Google Cloud OAuth **Desktop app** client JSON. Replace the `REPLACE_WITH_...` values using the downloaded Google client JSON, or replace this file with the downloaded JSON. Do not put an access or refresh token here; the script obtains a token through local sign-in and stores it in the ignored `token.json`.

The previously exposed Google OAuth should be revoked/rotated, not reused. Never send OAuth credentials or tokens in chat.

For private SharePoint files and workbook URLs, create/configure a Microsoft Entra public-client app with delegated `Files.Read.All` and `Sites.Read.All` permissions, then set its client ID for the current PowerShell session:

```powershell
$env:MS_GRAPH_CLIENT_ID = "YOUR_ENTRA_PUBLIC_CLIENT_APP_ID"
```

The app uses Microsoft's device sign-in for Graph. Your organization may require administrator approval. This is separate from Google OAuth.

## Install (Windows PowerShell)

From this directory:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Run

The script is preconfigured for the Kverneland OneDrive workbook and video folder, so
running it without arguments checks the latest recognizable request row. A blank
`YouTube Link` on that row approves it for upload; an existing link stops the run.
After a successful upload, the private YouTube URL is written back to that row.

```powershell
python .\auto_yt.py
```

Keep the workbook closed in Excel while the script writes the completed YouTube URL.
You can add `--dry-run` to perform only the local copy and verification.

To use another workbook, pass a local/synced `.xlsx` path or a SharePoint workbook URL:

```powershell
python .\auto_yt.py --workbook "C:\path\to\Request.xlsx"
```

For a OneDrive-synced video folder, provide `--local-video-dir`. The uploader searches
that folder and its subfolders for an exact, case-insensitive filename match from the
request, copies the match into `downloads`, and verifies it before continuing. Missing
or duplicate matches are reported instead of silently choosing a file. If you choose a
local video folder, make sure the workbook's `Video File` filename matches the synced
video's filename.

```powershell
python .\auto_yt.py `
  --workbook "C:\Users\kk_sle\OneDrive - Kverneland Group\Group Marketing and Communications - Documents\Request to upload video to Kverneland Group YouTube channel.xlsx" `
  --local-video-dir "C:\Users\kk_sle\OneDrive - Kverneland Group\Group Marketing and Communications - Documents\Apps\Microsoft Forms\KvG YouTube Music License\KvG YouTube files\Video File" `
  --dry-run
```

The local video folder removes the need for Graph access to the video itself. If the
workbook is still provided as a SharePoint URL rather than a local/synced `.xlsx` file,
Graph authentication is still needed to read the workbook.

The supplied Excel `Doc.aspx?sourcedoc=...&file=...` edit URL is supported. For the Kverneland workbook use:

```powershell
python .\auto_yt.py `
  --workbook "PASTE_THE_EXCEL_URL_HERE" `
  --site-url "https://kvgrp.sharepoint.com/sites/KvGGroupMarketing"
```

Rows that already have a YouTube link are skipped. Select a pending row from the menu, or specify the worksheet row:

```powershell
python .\auto_yt.py --workbook "C:\path\to\Request.xlsx" --row 12 --dry-run
```

A dry run downloads and verifies but never signs in to YouTube or uploads. A normal run stores the video in `downloads`, checks nonzero length and known file size, rejects HTML sign-in pages, checks the MP4 signature for `.mp4`, and prints byte size and SHA-256. Only after these checks does it ask whether to upload. Answer `y` to continue; any other answer cancels. A successful upload is then written into the local workbook's `YouTube Link` cell.

YouTube sign-in and upload happen only after download verification and confirmation. Uploads use the resumable API and remain private. Playlist insertion, thumbnails, captions, and music-license metadata are not handled yet.

## Tests

```powershell
python -m unittest discover -s tests -v
```

## Build and distribute the Windows EXE

The build uses PyInstaller to bundle Python and the libraries into one graphical
desktop client:

```powershell
.\build_exe.ps1
```

The output is the folder `dist\AutoYouTubeUploader\`, which contains
`AutoYouTubeUploader.exe` and its supporting `_internal` folder. Start the EXE from
inside that folder. The EXE performs a first-run setup:

1. Select the local/synced request workbook.
2. Select the synced `Video File` folder.
3. Select a Google OAuth **Desktop app** client JSON.

When open, the client checks for the latest blank-`YouTube Link` request immediately
and then every hour. It verifies the matching local video, uploads it as private, and
writes the resulting URL into the workbook. The window provides status/log output,
`Check now`, `Settings`, and `Quit` controls. The installed client version and update
status are displayed beneath the main status line.

Settings are saved in `config.json` in the same folder as the app (the EXE folder, or the
project folder when run from source), so setup is not repeated on each launch. The folder
must be writable. Google tokens and upload state stay per Windows user in
`%LOCALAPPDATA%\AutoYouTubeUploader`, and are not embedded in the EXE.
Use the `Settings` button in the window to change paths later.

Running `python auto_yt.py` or the EXE with no arguments opens the client window. Passing
arguments runs the command-line interface instead.

Distribute the release ZIP from the GitHub Releases page of this repository (see
[Client updates](#client-updates)), together with the authorized Google Desktop OAuth
`client.json`. Keep that JSON outside the application folder and do not publish it.
Each colleague unzips the folder to a writable location,
starts `AutoYouTubeUploader.exe` from there, selects the client JSON during first-run
setup, and then authenticates with their own Google/YouTube profile.
The selected client JSON is reused, but access and refresh tokens remain separate for
each Windows user.

Optional code signing: set `AUTOYT_SIGN_PFX` (and optionally `AUTOYT_SIGN_PFX_PASSWORD`
and `AUTOYT_SIGN_TIMESTAMP_URL`) before building, and `build_exe.ps1` signs the EXE with
`signtool` (Windows SDK required). Never commit the `.pfx` file or its password. Without
these variables the build is unsigned, which is the current setup.

## Client updates

The default update manifest is
`https://github.com/sigmundled/kvg-local-auto-yt-uploader/releases/latest/download/latest.json`.
It is used unless `update_manifest_url` is set in `config.json` or the setup form.
The client checks that manifest during each hourly cycle and shows an update notice with an
`Update` button. Updates are accepted only over HTTPS, require a matching 64-character
SHA-256 hash, and are verified before the application folder is replaced. The update
replaces the program files only; `config.json` and per-user data are kept.

Development loop (fast, no zip):

1. Edit the code and run `python -m unittest discover -s tests`.
2. Run `.\build_exe.ps1`. This rebuilds only the test folder `dist\AutoYouTubeUploader\`.
3. Test `dist\AutoYouTubeUploader\AutoYouTubeUploader.exe`. Repeat as needed.

Publishing a release (only when you are done editing):

1. Increment `APP_VERSION` in `auto_yt.py`, commit, run `.\build_exe.ps1`, and push.
2. Dry run to create the zip and manifest without publishing:

   ```powershell
   .\publish_update.ps1 -DryRun
   ```

3. Publish. This tags `vX.Y.Z`, pushes the tag, and creates a public GitHub release with
   the ZIP and `latest.json`. It needs the GitHub CLI (`gh`) and `gh auth login`:

   ```powershell
   .\publish_update.ps1 -Notes "Describe this release"
   ```

The script refuses to run if the test folder is older than the sources, the working tree
has uncommitted changes, HEAD is not pushed, or the tag already exists. It also refuses
if `config.json`, `client.json`, or `token.json` are in the folder.

This is poll-based update delivery: open clients notice a published update during their
next hourly check.

Clients built before this folder-based layout (single `AutoYouTubeUploader.exe` file)
cannot apply folder updates automatically. Each colleague needs one manual reinstall
from the new folder ZIP.
