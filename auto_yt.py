"""Local Excel-to-YouTube uploader with a verified-download gate."""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Callable
from urllib.parse import parse_qs, quote, unquote, urlsplit

import requests

GRAPH_SCOPES = ["Files.Read.All", "Sites.Read.All"]
YOUTUBE_SCOPE = ["https://www.googleapis.com/auth/youtube.upload"]
MP4_SIGNATURE = b"ftyp"
APP_VERSION = "1.3.0"
DEFAULT_UPDATE_MANIFEST_URL = (
    "https://github.com/sigmundled/kvg-local-auto-yt-uploader/releases/latest/download/latest.json"
)
IS_FROZEN = bool(getattr(sys, "frozen", False))
APP_DIR = (
    Path(sys.executable).resolve().parent
    if IS_FROZEN
    else Path(__file__).resolve().parent
)
USER_DATA_DIR = (
    Path(os.environ.get("LOCALAPPDATA", Path.home())) / "AutoYouTubeUploader"
    if os.name == "nt"
    else Path.home() / "AutoYouTubeUploader"
)
DEFAULT_WORKBOOK = (
    Path.home()
    / "OneDrive - Kverneland Group"
    / "Group Marketing and Communications - Documents"
    / "Request to upload video to Kverneland Group YouTube channel.xlsx"
)
DEFAULT_LOCAL_VIDEO_DIR = (
    Path.home()
    / "OneDrive - Kverneland Group"
    / "Group Marketing and Communications - Documents"
    / "Apps"
    / "Microsoft Forms"
    / "KvG YouTube Music License"
    / "KvG YouTube files"
    / "Video File"
)
ALIASES = {
    "id": {"id", "responseid"},
    "title": {"title", "videotitle", "youtubetitle"},
    "description": {"description", "videodescription"},
    "tags": {"tags", "videotags"},
    "playlist": {"playlist", "playlistid", "youtubeplaylist"},
    "video": {
        "video", "videofile", "videofilelink", "videourl", "linktovideo",
        "linktovideofile", "sharepointonedrive", "onsharepointonedrive",
        "sharepointonedrivevideo", "videoonsharepointonedrive",
        "videofileonsharepointonedrive",
    },
    "youtube": {"youtubelink", "youtubeurl", "uploadedvideolink"},
    "language": {"language", "videolanguage", "defaultlanguage"},
    "size": {"videosize", "filesize", "size"},
}


@dataclass(frozen=True)
class VideoRequest:
    row_number: int
    row_id: str
    title: str
    description: str
    tags: list[str]
    playlist: str
    language: str
    video_url: str
    video_name: str
    expected_size: int | None
    youtube_link: str


def _norm(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _column(headers: list[object], field: str) -> str | None:
    for header in headers:
        normalized = _norm(header)
        if normalized in ALIASES[field]:
            return str(header)
    if field == "video":
        for header in headers:
            normalized = _norm(header)
            if "video" in normalized and any(
                marker in normalized for marker in ("file", "link", "sharepoint", "onedrive")
            ):
                return str(header)
    return None


def _source(value: object, hyperlink: str | None = None) -> tuple[str, str, int | None]:
    if not isinstance(value, str) and not hyperlink:
        return "", "", None
    candidates = [str(hyperlink).strip()] if hyperlink else []
    if value is not None and str(value).strip() not in candidates:
        candidates.append(str(value).strip())
    for text in candidates:
        if text.startswith(("[", "{")):
            try:
                decoded = json.loads(text)
            except json.JSONDecodeError:
                decoded = None
            entries = decoded if isinstance(decoded, list) else [decoded]
            for entry in entries:
                if not isinstance(entry, dict):
                    continue
                name = str(entry.get("name") or "")
                try:
                    size = int(entry["size"]) if entry.get("size") is not None else None
                except (TypeError, ValueError):
                    size = None
                link = entry.get("link") or entry.get("url") or entry.get("downloadUrl")
                if link:
                    return str(link).strip(), name, size
        match = re.search(r"https?://[^\s\"<>]+", text, flags=re.IGNORECASE)
        if match:
            return match.group(0).rstrip(".,);"), "", None
        if Path(text).is_file():
            return str(Path(text).resolve()), Path(text).name, None
    return "", "", None


def _parse_tags(value: object) -> list[str]:
    if not value:
        return []
    if isinstance(value, (list, tuple)):
        return [str(tag).strip().lstrip("#") for tag in value if str(tag).strip()]
    text = str(value).strip()
    if text.startswith("["):
        try:
            decoded = json.loads(text)
        except json.JSONDecodeError:
            decoded = None
        if isinstance(decoded, list):
            return _parse_tags(decoded)
    return [tag.strip().lstrip("#") for tag in re.split(r"[,;\n]", text) if tag.strip()]


def _make_request(
    headers: list[object],
    values: dict[str, object],
    hyperlinks: dict[str, str],
    row_number: int,
) -> VideoRequest | None:
    title_col, video_col = _column(headers, "title"), _column(headers, "video")
    if not title_col or not video_col:
        return None
    title = str(values.get(title_col) or "").strip()
    url, name, embedded_size = _source(values.get(video_col), hyperlinks.get(video_col))
    if not url:
        for header, value in values.items():
            if any(part in _norm(header) for part in ("video", "sharepoint", "onedrive")):
                url, name, embedded_size = _source(value, hyperlinks.get(header))
                if url:
                    break
    if not title or not url:
        return None

    def field(name: str) -> str:
        column = _column(headers, name)
        return str(values.get(column) or "").strip() if column else ""

    try:
        size = int(field("size")) if field("size") else embedded_size
    except ValueError:
        size = embedded_size
    tag_col = _column(headers, "tags")
    return VideoRequest(
        row_number=row_number,
        row_id=field("id") or str(row_number),
        title=title,
        description=field("description"),
        tags=_parse_tags(values.get(tag_col)) if tag_col else [],
        playlist=field("playlist"),
        language=field("language"),
        video_url=url,
        video_name=name,
        expected_size=size,
        youtube_link=field("youtube"),
    )


def load_requests(workbook_path: str | Path, sheet_name: str | None = None) -> tuple[str, list[VideoRequest]]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError("Install dependencies with: python -m pip install -r requirements.txt") from exc
    workbook = load_workbook(workbook_path, data_only=True)
    sheets = [workbook[sheet_name]] if sheet_name else workbook.worksheets
    for sheet in sheets:
        header_row = None
        for row_number in range(1, min(sheet.max_row, 50) + 1):
            headers = [cell.value for cell in sheet[row_number]]
            if _column(headers, "title") and _column(headers, "video"):
                header_row = row_number
                break
        if header_row is None:
            continue
        headers = [cell.value for cell in sheet[header_row]]
        result = []
        for row_number in range(header_row + 1, sheet.max_row + 1):
            cells = sheet[row_number]
            values = {
                str(headers[i]): cells[i].value
                for i in range(min(len(headers), len(cells)))
                if headers[i] is not None
            }
            links = {
                str(headers[i]): cells[i].hyperlink.target
                for i in range(min(len(headers), len(cells)))
                if headers[i] is not None and cells[i].hyperlink
            }
            request = _make_request(headers, values, links, row_number)
            if request:
                result.append(request)
        if result:
            workbook.close()
            return sheet.title, result
    workbook.close()
    raise ValueError("No worksheet with recognizable Title and Video File columns was found.")


def _safe_filename(name: str, fallback: str = "video.mp4") -> str:
    base = Path(unquote(name)).name
    base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", base).strip(" .")
    if not base:
        base = fallback
    if not Path(base).suffix:
        base += ".mp4"
    return base


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def verify_download(path: str | Path, expected_size: int | None = None) -> tuple[int, str]:
    file_path = Path(path)
    if not file_path.is_file():
        raise ValueError("Downloaded video is missing.")
    size = file_path.stat().st_size
    if size <= 0:
        raise ValueError("Downloaded video is empty; refusing to upload.")
    if expected_size is not None and size != expected_size:
        raise ValueError(f"Download size mismatch: expected {expected_size} bytes, got {size}.")
    with file_path.open("rb") as stream:
        prefix = stream.read(512)
    if prefix.lstrip().lower().startswith((b"<!doctype html", b"<html", b"<head", b"<body")):
        raise ValueError("Downloaded content is an HTML page, not a video.")
    if file_path.suffix.lower() == ".mp4" and MP4_SIGNATURE not in prefix[:64]:
        raise ValueError("The .mp4 file does not have a valid MP4 ftyp signature.")
    return size, _sha256(file_path)


def _graph_cache_path() -> Path:
    root = Path(os.environ.get("LOCALAPPDATA", Path.home())) if os.name == "nt" else Path.home() / ".cache"
    return root / "AutoYouTubeUploader" / "msal_cache.bin"


class GraphClient:
    """Microsoft Graph delegated device-code client for SharePoint files."""

    def __init__(self, client_id: str, site_url: str | None):
        try:
            import msal
        except ImportError as exc:
            raise RuntimeError("Install dependencies with: python -m pip install -r requirements.txt") from exc
        self.site_url = site_url
        self.cache_path = _graph_cache_path()
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        self.cache = msal.SerializableTokenCache()
        if self.cache_path.exists():
            self.cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
        self.app = msal.PublicClientApplication(
            client_id,
            authority="https://login.microsoftonline.com/organizations",
            token_cache=self.cache,
        )

    def _token(self) -> str:
        accounts = self.app.get_accounts()
        result = self.app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0]) if accounts else None
        if not result:
            flow = self.app.initiate_device_flow(scopes=GRAPH_SCOPES)
            if "user_code" not in flow:
                raise RuntimeError("Could not start Microsoft sign-in; check the Entra app registration.")
            print(flow["message"])
            result = self.app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(
                "Microsoft sign-in failed: "
                + str(result.get("error_description", result.get("error", "unknown error")))
            )
        if self.cache.has_state_changed:
            self.cache_path.write_text(self.cache.serialize(), encoding="utf-8")
        return result["access_token"]

    def _json(self, url: str) -> dict:
        response = requests.get(
            url,
            headers={"Authorization": f"Bearer {self._token()}"},
            timeout=(15, 60),
        )
        response.raise_for_status()
        return response.json()

    @staticmethod
    def _share_id(url: str) -> str:
        encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
        return f"u!{encoded}"

    def _item(self, file_url: str) -> dict:
        share_url = f"https://graph.microsoft.com/v1.0/shares/{self._share_id(file_url)}/driveItem"
        try:
            return self._json(share_url)
        except requests.RequestException:
            pass
        site_url = self.site_url
        if not site_url:
            parsed = urlsplit(file_url)
            parts = [part for part in unquote(parsed.path).split("/") if part]
            for index, part in enumerate(parts[:-1]):
                if part.lower() in {"sites", "teams"}:
                    site_url = f"{parsed.scheme}://{parsed.netloc}/{part}/{parts[index + 1]}"
                    break
        if not site_url:
            raise RuntimeError("Use a SharePoint sharing link or provide --site-url.")
        site = urlsplit(site_url.rstrip("/"))
        site_api = (
            f"https://graph.microsoft.com/v1.0/sites/{site.netloc}:"
            f"{quote(unquote(site.path), safe='/')}"
        )
        try:
            site_item = self._json(site_api)
            drives = self._json(
                f"https://graph.microsoft.com/v1.0/sites/{site_item['id']}/drives"
            ).get("value", [])
            query = parse_qs(urlsplit(file_url).query)
            source_id = query.get("sourcedoc", [None])[0]
            document_name = query.get("file", [None])[0]
            if source_id and document_name:
                wanted_id = re.sub(r"[^a-fA-F0-9]", "", source_id).lower()
                for drive in drives:
                    search = quote(f"'{document_name}'", safe="'")
                    search_url = (
                        f"https://graph.microsoft.com/v1.0/drives/{drive['id']}"
                        f"/root/search(q={search})?$select=id,name,size,sharepointIds,parentReference"
                    )
                    try:
                        matches = self._json(search_url).get("value", [])
                    except requests.RequestException:
                        continue
                    for item in matches:
                        unique_id = item.get("sharepointIds", {}).get("listItemUniqueId", "")
                        normalized_id = re.sub(r"[^a-fA-F0-9]", "", unique_id).lower()
                        if normalized_id == wanted_id:
                            return item
                raise RuntimeError(
                    "Could not locate the workbook by its sourcedoc ID. Check the site URL, "
                    "file access, and Graph permissions."
                )
            target = unquote(urlsplit(file_url).path).rstrip("/").lower()
            matches = [
                drive for drive in drives
                if target.startswith(unquote(urlsplit(drive.get("webUrl", "")).path).rstrip("/").lower())
            ]
            matches.sort(key=lambda drive: len(drive.get("webUrl", "")), reverse=True)
            if not matches:
                raise RuntimeError("No SharePoint document library matched the file link.")
            drive = matches[0]
            library = unquote(urlsplit(drive["webUrl"]).path).rstrip("/")
            relative = unquote(urlsplit(file_url).path)[len(library):].strip("/")
            item_url = (
                f"https://graph.microsoft.com/v1.0/drives/{drive['id']}/root:/"
                f"{quote(relative, safe='/')}:/"
            )
            return self._json(item_url)
        except (requests.RequestException, KeyError) as exc:
            raise RuntimeError(
                "Graph could not resolve the SharePoint file. Check access and delegated "
                "Files.Read.All/Sites.Read.All permissions."
            ) from exc

    def download_url(self, file_url: str) -> tuple[str, int, str]:
        item = self._item(file_url)
        url = item.get("@microsoft.graph.downloadUrl")
        if not url:
            raise RuntimeError("Microsoft Graph returned no file download URL.")
        return url, int(item.get("size") or 0), str(item.get("name") or "")


def _stream_response(
    response: requests.Response,
    destination: Path,
    expected_size: int | None,
) -> None:
    response.raise_for_status()
    if "text/html" in response.headers.get("Content-Type", "").lower():
        raise ValueError("The server returned a sign-in/error page instead of the video.")
    header_size = response.headers.get("Content-Length")
    size_to_check = expected_size
    if size_to_check is None and header_size and header_size.isdigit():
        size_to_check = int(header_size)
    partial = destination.with_suffix(destination.suffix + ".part")
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        with partial.open("wb") as stream:
            for block in response.iter_content(chunk_size=1024 * 1024):
                if block:
                    stream.write(block)
        verify_download(partial, size_to_check)
        if destination.suffix.lower() == ".mp4":
            with partial.open("rb") as stream:
                if MP4_SIGNATURE not in stream.read(64):
                    raise ValueError("Downloaded content has no MP4 signature.")
        partial.replace(destination)
    except Exception:
        partial.unlink(missing_ok=True)
        raise


def _copy_verified_local_video(
    source: Path,
    destination: Path,
    expected_size: int | None,
) -> Path:
    try:
        shutil.copyfile(source, destination)
        verify_download(destination, expected_size)
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    return destination


def _find_local_video(directory: str | Path, name: str) -> Path:
    folder = Path(directory).expanduser()
    if not folder.is_dir():
        raise NotADirectoryError(f"Local video folder does not exist: {folder}")
    matches = [
        path
        for path in folder.rglob("*")
        if path.is_file() and path.name.casefold() == name.casefold()
    ]
    if not matches:
        raise FileNotFoundError(f"No local video named '{name}' found in: {folder}")
    if len(matches) > 1:
        locations = ", ".join(str(path) for path in matches[:5])
        more = " (and more)" if len(matches) > 5 else ""
        raise ValueError(f"Multiple local videos named '{name}' found: {locations}{more}")
    return matches[0]


def download_video(
    request: VideoRequest,
    directory: str | Path,
    graph: GraphClient | None,
    local_video_dir: str | Path | None = None,
) -> Path:
    parsed = urlsplit(request.video_url)
    name = request.video_name or Path(unquote(parsed.path)).name or "video.mp4"
    destination = Path(directory) / f"row-{request.row_number}-{_safe_filename(name)}"
    if destination.exists():
        suffix = hashlib.sha256(request.video_url.encode("utf-8")).hexdigest()[:8]
        destination = destination.with_name(f"{destination.stem}-{suffix}{destination.suffix}")
    destination.parent.mkdir(parents=True, exist_ok=True)

    if local_video_dir:
        source = _find_local_video(local_video_dir, name)
        return _copy_verified_local_video(source, destination, request.expected_size)

    local_file = Path(request.video_url)
    if local_file.is_file():
        return _copy_verified_local_video(local_file, destination, request.expected_size)
    if parsed.scheme.lower() == "file":
        local_file = Path(unquote(parsed.path))
        if not local_file.is_file():
            raise ValueError("Local video path does not exist.")
        return _copy_verified_local_video(local_file, destination, request.expected_size)
    if parsed.scheme.lower() not in {"http", "https"}:
        raise ValueError("Video source must be an HTTP(S) URL or an existing local file.")
    try:
        with requests.get(request.video_url, stream=True, timeout=(15, 120)) as response:
            _stream_response(response, destination, request.expected_size)
        return destination
    except (requests.RequestException, ValueError):
        destination.unlink(missing_ok=True)
    if graph is None:
        raise RuntimeError(
            "Video download requires SharePoint sign-in. Set MS_GRAPH_CLIENT_ID and pass --site-url."
        )
    download_url, graph_size, graph_name = graph.download_url(request.video_url)
    if graph_name:
        destination = destination.with_name(f"row-{request.row_number}-{_safe_filename(graph_name)}")
    expected = request.expected_size or (graph_size or None)
    with requests.get(download_url, stream=True, timeout=(15, 120)) as response:
        _stream_response(response, destination, expected)
    return destination


def authenticate_youtube(client_secrets: str | Path, token_file: str | Path):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    secrets = Path(client_secrets)
    if not secrets.is_file():
        raise FileNotFoundError(f"Google OAuth client JSON not found: {secrets}")
    credentials_file = Path(token_file)
    credentials = (
        Credentials.from_authorized_user_file(str(credentials_file), YOUTUBE_SCOPE)
        if credentials_file.is_file()
        else None
    )
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
    if not credentials or not credentials.valid:
        flow = InstalledAppFlow.from_client_secrets_file(str(secrets), YOUTUBE_SCOPE)
        credentials = flow.run_local_server(port=0, access_type="offline")
    credentials_file.parent.mkdir(parents=True, exist_ok=True)
    credentials_file.write_text(credentials.to_json(), encoding="utf-8")
    return build("youtube", "v3", credentials=credentials, cache_discovery=False)


def upload_video(youtube, request: VideoRequest, path: str | Path) -> str:
    from googleapiclient.http import MediaFileUpload

    metadata = {
        "snippet": {
            "title": request.title[:100],
            "description": request.description,
            "tags": request.tags,
            "categoryId": "22",
        },
        "status": {"privacyStatus": "private"},
    }
    if request.language and re.fullmatch(r"[a-zA-Z]{2,3}(?:-[a-zA-Z0-9]{2,8})*", request.language):
        metadata["snippet"]["defaultLanguage"] = request.language
    media = MediaFileUpload(
        str(path),
        mimetype="video/mp4" if Path(path).suffix.lower() == ".mp4" else None,
        chunksize=8 * 1024 * 1024,
        resumable=True,
    )
    call = youtube.videos().insert(part="snippet,status", body=metadata, media_body=media)
    response = None
    while response is None:
        status, response = call.next_chunk()
        if status:
            print(f"YouTube upload progress: {int(status.progress() * 100)}%")
    if not response.get("id"):
        raise RuntimeError("YouTube finished the upload without returning a video ID.")
    return response["id"]


def download_then_upload(
    request: VideoRequest,
    downloader: Callable[[VideoRequest], Path],
    uploader: Callable[[VideoRequest, Path], str],
    confirm: Callable[[str], str],
    dry_run: bool = False,
) -> str | None:
    path = downloader(request)
    size, digest = verify_download(path, request.expected_size)
    print(f"Verified download: {path} | {size:,} bytes | SHA-256 {digest}")
    if dry_run:
        print("Dry run: no YouTube sign-in or upload was attempted.")
        return None
    answer = confirm("Upload this verified video as PRIVATE? [y/N] ")
    if answer.strip().lower() not in {"y", "yes"}:
        print("Upload cancelled; the verified video remains on disk.")
        return None
    return uploader(request, path)


def _state_path() -> Path:
    return USER_DATA_DIR / "upload_state.json"


def _token_path() -> Path:
    return USER_DATA_DIR / "token.json"


def _config_path() -> Path:
    return APP_DIR / "config.json"


def _read_state(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Upload-state file is invalid JSON: {path}") from exc
    if not isinstance(state, dict):
        raise ValueError("Upload-state file must contain a JSON object.")
    return state


def _write_state(path: Path, state: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(state, indent=2), encoding="utf-8")
    temporary.replace(path)


def _read_config(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    try:
        config = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Uploader configuration is invalid JSON: {path}") from exc
    if not isinstance(config, dict) or not all(
        isinstance(key, str) and isinstance(value, str)
        for key, value in config.items()
    ):
        raise ValueError(f"Uploader configuration must contain only string values: {path}")
    return config


def _prompt_existing_path(
    prompt: str,
    expected_type: str,
    default: Path | None,
    ask: Callable[[str], str],
) -> Path:
    suffix = f" [{default}]" if default else ""
    while True:
        answer = ask(f"{prompt}{suffix}: ").strip().strip('"')
        selected = Path(answer).expanduser() if answer else default
        if selected is None:
            print("A path is required.")
        elif expected_type == "file" and selected.is_file():
            return selected.resolve()
        elif expected_type == "directory" and selected.is_dir():
            return selected.resolve()
        else:
            print(f"That {expected_type} was not found: {selected}")


def _run_setup(config_path: Path, ask: Callable[[str], str] = input) -> dict[str, str]:
    print("\nAutoYouTubeUploader first-time setup")
    print("Press Enter to accept a shown default, or paste the full path.\n")
    workbook = _prompt_existing_path(
        "Local request workbook (.xlsx)",
        "file",
        DEFAULT_WORKBOOK if DEFAULT_WORKBOOK.is_file() else None,
        ask,
    )
    video_dir = _prompt_existing_path(
        "Synced SharePoint video folder",
        "directory",
        DEFAULT_LOCAL_VIDEO_DIR if DEFAULT_LOCAL_VIDEO_DIR.is_dir() else None,
        ask,
    )
    client_secrets = _prompt_existing_path(
        "Google Desktop OAuth client JSON",
        "file",
        APP_DIR / "client.json" if (APP_DIR / "client.json").is_file() else None,
        ask,
    )
    try:
        client_data = json.loads(client_secrets.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Google OAuth client JSON is invalid: {client_secrets}") from exc
    if not isinstance(client_data, dict) or "installed" not in client_data:
        raise ValueError("Google OAuth file must be a Desktop application 'installed' client JSON.")

    config = {
        "workbook": str(workbook),
        "local_video_dir": str(video_dir),
        "client_secrets": str(client_secrets),
    }
    _save_config(config_path, config)
    print(f"\nSetup saved for this Windows profile: {config_path}")
    return config


def _save_config(config_path: Path, config: dict[str, str]) -> None:
    config_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = config_path.with_suffix(config_path.suffix + ".tmp")
    temporary.write_text(json.dumps(config, indent=2), encoding="utf-8")
    temporary.replace(config_path)


def _write_youtube_link(
    workbook_path: str | Path,
    sheet_name: str,
    row_number: int,
    youtube_link: str,
) -> None:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError("Install dependencies with: python -m pip install -r requirements.txt") from exc

    path = Path(workbook_path).expanduser()
    workbook = load_workbook(path)
    temporary = path.with_name(f".{path.stem}.auto-yt.tmp.xlsx")
    try:
        sheet = workbook[sheet_name]
        header_row = None
        youtube_column = None
        for candidate in range(1, min(sheet.max_row, 50) + 1):
            headers = [cell.value for cell in sheet[candidate]]
            if _column(headers, "title") and _column(headers, "video"):
                header_row = candidate
                header_name = _column(headers, "youtube")
                if header_name is None:
                    raise ValueError("The workbook has no recognizable YouTube Link column.")
                youtube_column = headers.index(header_name) + 1
                break
        if header_row is None or youtube_column is None:
            raise ValueError("Could not locate the request table in the workbook.")
        if row_number <= header_row or row_number > sheet.max_row:
            raise ValueError(f"Worksheet row {row_number} is outside the request table.")
        cell = sheet.cell(row=row_number, column=youtube_column)
        if str(cell.value or "").strip():
            raise RuntimeError(
                f"Worksheet row {row_number} already has a YouTube link; it was not overwritten."
            )
        cell.value = youtube_link
        workbook.save(temporary)
        workbook.close()
        temporary.replace(path)
    except Exception:
        workbook.close()
        temporary.unlink(missing_ok=True)
        raise


def _workbook_file(source: str, graph: GraphClient | None) -> tuple[Path, tempfile.TemporaryDirectory | None]:
    local = Path(source).expanduser()
    if local.is_file():
        return local, None
    if urlsplit(source).scheme.lower() not in {"http", "https"}:
        raise FileNotFoundError(f"Workbook not found: {local}")
    if graph is None:
        raise RuntimeError("A SharePoint workbook URL requires MS_GRAPH_CLIENT_ID.")
    download_url, expected_size, name = graph.download_url(source)
    temporary = tempfile.TemporaryDirectory(prefix="auto-yt-workbook-")
    workbook = Path(temporary.name) / _safe_filename(name or "requests.xlsx", "requests.xlsx")
    with requests.get(download_url, stream=True, timeout=(15, 120)) as response:
        response.raise_for_status()
        with workbook.open("wb") as stream:
            for block in response.iter_content(chunk_size=1024 * 1024):
                if block:
                    stream.write(block)
    if expected_size and workbook.stat().st_size != expected_size:
        temporary.cleanup()
        raise ValueError("Downloaded workbook size does not match SharePoint metadata.")
    return workbook, temporary


def _choose(requests_: list[VideoRequest], row: int | None) -> VideoRequest:
    if row is not None:
        for request in requests_:
            if request.row_number == row:
                return request
        raise ValueError(f"Worksheet row {row} is not pending.")
    print("\nPending video requests:")
    for request in requests_:
        print(f"  row {request.row_number}: {request.title}")
    while True:
        choice = input("Choose a row number (or q to quit): ").strip()
        if choice.lower() == "q":
            raise SystemExit(0)
        if choice.isdigit():
            for request in requests_:
                if request.row_number == int(choice):
                    return request
        print("Enter a row number from the pending list.")


def main(argv: list[str] | None = None) -> int:
    config_path = _config_path()
    config = _read_config(config_path)
    workbook_default = config.get("workbook", str(DEFAULT_WORKBOOK))
    video_dir_default = config.get("local_video_dir")
    client_default = config.get("client_secrets", str(APP_DIR / "client.json"))
    download_default = (
        str(USER_DATA_DIR / "downloads") if IS_FROZEN else str(APP_DIR / "downloads")
    )
    parser = argparse.ArgumentParser(
        description="Select an Excel video request, verify its downloaded file, and optionally upload it."
    )
    parser.add_argument(
        "--workbook",
        default=workbook_default,
        help="Local .xlsx path or SharePoint workbook URL",
    )
    parser.add_argument("--sheet", help="Worksheet name; detected automatically if omitted")
    parser.add_argument("--row", type=int, help="Worksheet row; otherwise select interactively")
    parser.add_argument("--site-url", help="SharePoint site root")
    parser.add_argument("--graph-client-id", default=os.environ.get("MS_GRAPH_CLIENT_ID"))
    parser.add_argument("--client-secrets", default=client_default)
    parser.add_argument("--youtube-token", default=str(_token_path()))
    parser.add_argument("--download-dir", default=download_default)
    parser.add_argument(
        "--local-video-dir",
        default=video_dir_default,
        help="Search this OneDrive-synced folder for an exact filename match before using the network",
    )
    parser.add_argument("--state-file", default=str(_state_path()))
    parser.add_argument("--dry-run", action="store_true", help="Download and verify without uploading")
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Confirm the private upload without an interactive prompt",
    )
    parser.add_argument(
        "--configure",
        action="store_true",
        help="Run the first-time path and OAuth client setup again",
    )
    args = parser.parse_args(argv)
    if args.configure:
        _run_setup(config_path)
        return 0
    if IS_FROZEN and not config:
        config = _run_setup(config_path)
        args.workbook = config["workbook"]
        args.local_video_dir = config["local_video_dir"]
        args.client_secrets = config["client_secrets"]
    if args.workbook == str(DEFAULT_WORKBOOK) and args.local_video_dir is None:
        args.local_video_dir = str(DEFAULT_LOCAL_VIDEO_DIR)

    graph = GraphClient(args.graph_client_id, args.site_url) if args.graph_client_id else None
    workbook, temporary = _workbook_file(args.workbook, graph)
    try:
        sheet, rows = load_requests(workbook, args.sheet)
    finally:
        if temporary:
            temporary.cleanup()
    source_key = (
        str(Path(args.workbook).expanduser().resolve())
        if Path(args.workbook).is_file()
        else args.workbook
    )
    source_hash = hashlib.sha256(source_key.encode("utf-8")).hexdigest()
    state_path = Path(args.state_file).expanduser()
    state = _read_state(state_path)
    if args.row is None:
        request = rows[-1]
        if request.youtube_link:
            print(
                f"No upload requested: latest request row {request.row_number} already has a "
                "YouTube link."
            )
            return 0
        prior = state.get(f"{source_hash}|{request.row_id}")
        if isinstance(prior, dict) and prior.get("youtube_link") and temporary is None:
            link = str(prior["youtube_link"])
            _write_youtube_link(Path(args.workbook), sheet, request.row_number, link)
            print(f"Restored upload link in latest request row: {link}")
            return 0
    else:
        pending = [
            item for item in rows
            if not item.youtube_link and f"{source_hash}|{item.row_id}" not in state
        ]
        if not pending:
            print("No pending requests; rows with an existing YouTube link or local success record are skipped.")
            return 0
        request = _choose(pending, args.row)
    print(f"\nSelected '{sheet}' row {request.row_number}: {request.title}")

    def downloader(item: VideoRequest) -> Path:
        return download_video(item, args.download_dir, graph, args.local_video_dir)

    def uploader(item: VideoRequest, path: Path) -> str:
        youtube = authenticate_youtube(args.client_secrets, args.youtube_token)
        return upload_video(youtube, item, path)

    confirm_upload = (lambda _prompt: "y") if args.yes else input
    video_id = download_then_upload(request, downloader, uploader, confirm_upload, args.dry_run)
    if not video_id:
        return 0
    link = f"https://www.youtube.com/watch?v={video_id}"
    state[f"{source_hash}|{request.row_id}"] = {
        "video_id": video_id,
        "youtube_link": link,
        "title": request.title,
    }
    _write_state(state_path, state)
    if temporary is None:
        try:
            _write_youtube_link(Path(args.workbook), sheet, request.row_number, link)
        except Exception as exc:
            raise RuntimeError(
                f"Upload succeeded as {link}, but the YouTube Link could not be written to "
                "the workbook. The saved link remains in the local upload state."
            ) from exc
    else:
        print("The workbook is remote, so it was not modified.")
    if request.playlist:
        print("Playlist insertion is not enabled; upload-only OAuth is used.")
    print(f"Upload complete: {link} (private)")
    if temporary is None:
        print("YouTube Link written to the workbook.")
    return 0


if __name__ == "__main__":
    if len(sys.argv) == 1:
        from client_app import run_client

        raise SystemExit(run_client())
    raise SystemExit(main())
