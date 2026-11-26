"""Modern customtkinter desktop client for scheduled local YouTube upload checks."""

from __future__ import annotations

import contextlib
import io
import json
import os
import queue
import sys
import threading
import time
import tkinter as tk
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from urllib.parse import parse_qs, urlsplit

import customtkinter as ctk
import pystray
from PIL import Image, ImageDraw

import auto_yt
from updater import check_update, download_update, install_update, _update_marker_path

DEFAULT_INTERVAL_MINUTES = 60
MIN_INTERVAL_MINUTES = 5
CARD = ("#ffffff", "#1d212b")
SUBTLE = ("#eef1f6", "#2a303d")
SUBTLE_HOVER = ("#e0e5ee", "#353c4b")
MUTED = ("#5f6b7a", "#9aa4b5")
TEXT = ("#15181f", "#eef1f6")
ACCENT = "#007b3d"
ACCENT_HOVER = "#00632f"
YOUTUBE_RED = "#ff0000"
ERROR_RED = "#e62117"
STATUS_COLORS = {
    "ready": ACCENT,
    "busy": "#d97706",
    "error": ERROR_RED,
    "update_available": ACCENT,
}
DONE_COLOR = ACCENT
# (stage name from auto_yt.main, step label shown in the status card)
STAGE_STEPS = [
    ("checking", "Checking workbook"),
    ("downloading", "Downloading video"),
    ("uploading", "Uploading video"),
    ("thumbnail", "Uploading thumbnail"),
    ("saving", "Saving link"),
]
# Stages without a step of their own map onto the step they belong to.
STAGE_INDEX = {"checking": 0, "downloading": 1, "confirm": 1, "uploading": 2, "thumbnail": 3, "saving": 4, "done": 5}
# (config key, label, browse kind, dialog title); kind is None for plain text fields.
FIELD_SPECS = [
    ("workbook", "Request workbook (.xlsx)", "xlsx", "Select the request workbook"),
    ("local_video_dir", "Synced Video File folder", "dir", "Select the synced Video File folder"),
    ("client_secrets", "Google Desktop OAuth client JSON", "json", "Select the Google OAuth client JSON"),
    ("check_interval_minutes", "Automatic check interval (minutes)", None, ""),
]
FILE_TYPES = {
    "xlsx": [("Excel workbook", "*.xlsx")],
    "json": [("JSON file", "*.json")],
}


def interval_minutes(config: dict[str, str]) -> int:
    try:
        minutes = int(config.get("check_interval_minutes", DEFAULT_INTERVAL_MINUTES))
    except ValueError:
        return DEFAULT_INTERVAL_MINUTES
    return minutes if minutes >= MIN_INTERVAL_MINUTES else DEFAULT_INTERVAL_MINUTES


def _font(size: int, weight: str = "normal", family: str = "Segoe UI") -> ctk.CTkFont:
    return ctk.CTkFont(family=family, size=size, weight=weight)


def _read_changelog() -> str:
    # PyInstaller one-file builds unpack bundled data files into _MEIPASS.
    base = Path(getattr(sys, "_MEIPASS", auto_yt.APP_DIR))
    path = base / "CHANGELOG.md"
    if not path.is_file():
        return "Changelog not found."
    return path.read_text(encoding="utf-8")


def _log_path() -> Path:
    return auto_yt.USER_DATA_DIR / "client.log"


def _is_desktop_oauth_file(path: Path) -> bool:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False
    return isinstance(data, dict) and "installed" in data


def _open_path(path: Path) -> None:
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        webbrowser.open(path.as_uri())


def _youtube_studio_url(url: str) -> str | None:
    parsed = urlsplit(url)
    host = parsed.netloc.casefold().split(":", 1)[0]
    segments = [segment for segment in parsed.path.split("/") if segment]
    video_id = ""
    if host == "youtu.be" and segments:
        video_id = segments[0]
    elif host == "youtube.com" or host.endswith(".youtube.com"):
        if parsed.path.rstrip("/") == "/watch":
            video_id = parse_qs(parsed.query).get("v", [""])[0]
        elif len(segments) > 1 and segments[0] in {"embed", "shorts", "live"}:
            video_id = segments[1]
    if not video_id:
        return None
    return f"https://studio.youtube.com/video/{video_id}/edit"


def _format_features(version: str, notes: str, release_date: str = "") -> str:
    lines = [line.strip() for line in notes.splitlines() if line.strip()]
    date = release_date
    for line in lines[:2]:
        if line.casefold().startswith("released:"):
            date = line.partition(":")[2].strip()
            lines.remove(line)
            break
    heading = f"Version {version}"
    if date:
        heading += f"  ·  {date}"
    bullets = [
        f"• {line[2:].strip() if line.startswith('- ') else line[2:].strip()}"
        for line in lines
        if line.startswith(("- ", "* "))
    ]
    return "\n".join([heading, *bullets]) if bullets else heading


class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, master: ctk.CTk, existing: dict[str, str], required: bool):
        super().__init__(master)
        self.title("Welcome" if required else "Settings")
        self.geometry("680x600")
        self.resizable(False, False)
        self.result: dict[str, str] | None = None
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.fields: dict[str, ctk.StringVar] = {}

        defaults = {
            "workbook": existing.get("workbook", str(auto_yt.DEFAULT_WORKBOOK)),
            "local_video_dir": existing.get("local_video_dir", str(auto_yt.DEFAULT_LOCAL_VIDEO_DIR)),
            "client_secrets": existing.get("client_secrets", str(auto_yt.APP_DIR / "client.json")),
            "check_interval_minutes": existing.get(
                "check_interval_minutes", str(DEFAULT_INTERVAL_MINUTES)
            ),
        }

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(side="bottom", fill="x", padx=28, pady=20)
        ctk.CTkButton(
            footer,
            text="Save and continue" if required else "Save",
            font=_font(14, "bold"),
            height=40,
            corner_radius=10,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=self._save,
        ).pack(side="right")
        ctk.CTkButton(
            footer,
            text="Cancel",
            font=_font(14),
            height=40,
            corner_radius=10,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            text_color="white",
            command=self._cancel,
        ).pack(side="right", padx=8)

        ctk.CTkLabel(
            self,
            text="Welcome to Auto YouTube Uploader" if required else "Settings",
            font=_font(22, "bold"),
            anchor="w",
        ).pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(
            self,
            text="These settings are saved in config.json next to the app and reused on every launch.",
            font=_font(12),
            text_color=MUTED,
            anchor="w",
            justify="left",
            wraplength=600,
        ).pack(fill="x", padx=28, pady=(4, 8))

        for key, label, kind, title in FIELD_SPECS:
            ctk.CTkLabel(self, text=label, font=_font(13, "bold"), anchor="w").pack(
                fill="x", padx=28, pady=(10, 2)
            )
            row = ctk.CTkFrame(self, fg_color="transparent")
            row.pack(fill="x", padx=28)
            var = ctk.StringVar(value=defaults[key])
            self.fields[key] = var
            ctk.CTkEntry(row, textvariable=var, height=36, corner_radius=8).pack(
                side="left", fill="x", expand=True
            )
            if kind:
                ctk.CTkButton(
                    row,
                    text="Browse…",
                    width=96,
                    height=36,
                    corner_radius=8,
                    fg_color=ACCENT,
                    hover_color=ACCENT_HOVER,
                    text_color="white",
                    command=lambda k=key, kd=kind, t=title: self._browse(k, kd, t),
                ).pack(side="left", padx=(8, 0))

        self.after(100, self._grab)

    def _grab(self) -> None:
        with contextlib.suppress(tk.TclError):
            self.grab_set()
        self.focus_force()

    def show(self) -> dict[str, str] | None:
        self.wait_window()
        return self.result

    def _browse(self, key: str, kind: str, title: str) -> None:
        current = Path(self.fields[key].get().strip().strip('"')).expanduser()
        initial_dir = current.parent if current.parent.is_dir() else auto_yt.APP_DIR
        if kind == "dir":
            selected = filedialog.askdirectory(parent=self, title=title, initialdir=str(initial_dir), mustexist=True)
        else:
            selected = filedialog.askopenfilename(
                parent=self,
                title=title,
                initialdir=str(initial_dir),
                filetypes=FILE_TYPES[kind],
            )
        if selected:
            self.fields[key].set(selected)

    def _save(self) -> None:
        values = {key: var.get().strip().strip('"') for key, var in self.fields.items()}
        workbook = Path(values["workbook"]).expanduser()
        video_dir = Path(values["local_video_dir"]).expanduser()
        client_secrets = Path(values["client_secrets"]).expanduser()
        interval_text = values["check_interval_minutes"]

        errors: list[str] = []
        if not workbook.is_file():
            errors.append("The request workbook was not found.")
        if not video_dir.is_dir():
            errors.append("The Video File folder was not found.")
        if not _is_desktop_oauth_file(client_secrets):
            errors.append("The Google OAuth file must be a Desktop application client JSON.")
        if not interval_text.isdigit() or int(interval_text) < MIN_INTERVAL_MINUTES:
            errors.append(f"The check interval must be a whole number of minutes, at least {MIN_INTERVAL_MINUTES}.")
        if errors:
            messagebox.showerror("Please check the settings", "\n".join(errors), parent=self)
            return

        config = {
            "workbook": str(workbook.resolve()),
            "local_video_dir": str(video_dir.resolve()),
            "client_secrets": str(client_secrets.resolve()),
            "check_interval_minutes": str(int(interval_text)),
        }
        try:
            auto_yt._save_config(auto_yt._config_path(), config)
        except OSError as exc:
            messagebox.showerror(
                "Could not save settings",
                f"{exc}\n\nThe app folder must be writable so config.json can be saved:\n{auto_yt.APP_DIR}",
                parent=self,
            )
            return
        self.result = config
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


def show_settings(master: ctk.CTk, required: bool) -> dict[str, str] | None:
    existing = auto_yt._read_config(auto_yt._config_path())
    return SettingsDialog(master, existing, required).show()


class AutoYouTubeApp:
    def __init__(self, root: ctk.CTk, config: dict[str, str]):
        self.root = root
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.busy = False
        self.paused = False
        self.available_update: dict[str, str] | None = None
        self.downloaded_update: Path | None = None
        self.install_after_download = False
        self.announced_update = ""
        self.last_link: str | None = None
        self.last_playlist = ""
        self.log_lines: list[str] = []
        self.features_expanded = False
        self.tray_icon: pystray.Icon | None = None
        self.interval_minutes = interval_minutes(config)
        self.next_check_at = 0.0
        self.update_check_error = False
        self.smoke = bool(os.environ.get("AUTOYT_GUI_SMOKE_MARKER"))

        root.title(f"Auto YouTube Uploader {auto_yt.APP_VERSION}")
        root.geometry("1040x820")
        root.minsize(840, 600)
        root.resizable(True, True)
        root.protocol("WM_DELETE_WINDOW", self._close)
        root.bind("<Unmap>", self._on_minimize)
        root.bind("<Configure>", self._resize_content)
        self._build()
        self._refresh_uploads()
        self._set_status("busy", "Checking…")
        if not self.smoke:
            root.after(300, self._check_interrupted_update)
        if not self.smoke:
            self.next_check_at = time.time() + self.interval_minutes * 60
            root.after(400, self.check_now)
        root.after(150, self._process_events)
        root.after(1000, self._tick)

    # Layout

    def _build(self) -> None:
        root = ctk.CTkScrollableFrame(
            self.root,
            corner_radius=0,
            fg_color="transparent",
            scrollbar_button_color=SUBTLE_HOVER,
            scrollbar_button_hover_color=MUTED,
        )
        self.content = root
        root.grid(row=0, column=0, sticky="nsew")
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(root, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(16, 4))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Auto YouTube Uploader", font=_font(26, "bold"), anchor="w").grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkLabel(
            header,
            text=f"Version {auto_yt.APP_VERSION}",
            font=_font(12),
            text_color=MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="w")
        self.pill = ctk.CTkButton(
            header,
            text="Checking…",
            height=32,
            width=150,
            corner_radius=16,
            fg_color=STATUS_COLORS["ready"],
            hover_color=ACCENT_HOVER,
            text_color="white",
            font=_font(13, "bold"),
            command=self._on_status_click,
        )
        self.pill.grid(row=0, column=1, rowspan=2, sticky="e")
        self.status_menu = ctk.CTkOptionMenu(
            header,
            values=["Choose an option", "Check for client updates", "Latest features"],
            command=self._handle_status_menu,
            width=190,
            height=32,
            font=_font(12),
            fg_color=ACCENT,
            button_color=ACCENT_HOVER,
            button_hover_color=ACCENT,
            dropdown_fg_color=CARD,
            dropdown_text_color=TEXT,
        )
        self.status_menu.set("Choose an option")
        self.status_menu.grid(row=2, column=1, sticky="e", pady=(6, 0))
        self.status_menu.grid_remove()

        self.update_bar = ctk.CTkFrame(root, fg_color=("#edf8f1", "#173b2a"), corner_radius=14)
        update_row = ctk.CTkFrame(self.update_bar, fg_color="transparent")
        update_row.pack(fill="x", padx=14, pady=(10, 0))
        self.update_label = ctk.CTkLabel(update_row, text="", font=_font(13, "bold"), anchor="w", justify="left")
        self.update_label.pack(side="left", padx=(4, 12), fill="x", expand=True)
        self.download_button = ctk.CTkButton(
            update_row,
            text="Download update",
            width=140,
            height=36,
            corner_radius=10,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            font=_font(12, "bold"),
            command=self.download_available_update,
            state="disabled",
        )
        self.download_button.pack(side="left", padx=(0, 8))
        self.install_button = ctk.CTkButton(
            update_row,
            text="Install update",
            width=130,
            height=36,
            corner_radius=10,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            font=_font(12, "bold"),
            command=self.install_available_update,
            state="disabled",
        )
        self.install_button.pack(side="left", padx=(0, 8))
        self.download_install_button = ctk.CTkButton(
            update_row,
            text="Download and install",
            width=170,
            height=36,
            corner_radius=10,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            font=_font(12, "bold"),
            command=self.download_and_install_update,
            state="disabled",
        )
        self.download_install_button.pack(side="left", padx=(0, 4))
        self.update_warning = ctk.CTkLabel(
            self.update_bar,
            text="Don't close the CMD once it opens. It will look like it is doing nothing, but an active process is happening.",
            font=_font(11),
            text_color=MUTED,
            anchor="w",
        )
        self.update_warning.pack(fill="x", padx=18, pady=(2, 8))
        self.update_notes = ctk.CTkLabel(
            self.update_bar,
            text="",
            font=_font(12),
            anchor="w",
            justify="left",
        )
        self.update_notes.pack(fill="x", padx=18, pady=(0, 10))
        self.update_bar.grid(row=1, column=0, sticky="ew", padx=24, pady=(4, 0))
        self.update_bar.grid_remove()

        hero = ctk.CTkFrame(root, fg_color=CARD, corner_radius=18)
        hero.grid(row=2, column=0, sticky="ew", padx=24, pady=6)
        actions = ctk.CTkFrame(hero, fg_color="transparent")
        actions.pack(fill="x", padx=18, pady=12)
        secondary = [
            ("Open workbook", self.open_workbook),
            ("Open video folder", self.open_video_folder),
            ("Settings", self.open_settings),
        ]
        for text, command in secondary:
            self._secondary_button(actions, text, command).pack(side="left", padx=(0, 8))

        self.features_panel = ctk.CTkFrame(root, fg_color=CARD, corner_radius=14)
        ctk.CTkLabel(
            self.features_panel,
            text="LATEST UPDATE FEATURES",
            font=_font(11, "bold"),
            text_color=MUTED,
            anchor="w",
        ).pack(fill="x", padx=18, pady=(10, 2))
        self.features_text = ctk.CTkLabel(
            self.features_panel,
            text=self._latest_features(),
            font=_font(12),
            anchor="w",
            justify="left",
            wraplength=950,
        )
        self.features_text.pack(fill="x", padx=18, pady=(0, 12))
        self.features_panel.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 4))
        self.features_panel.grid_remove()

        cards = ctk.CTkFrame(root, fg_color="transparent")
        cards.grid(row=4, column=0, sticky="ew", padx=24, pady=6)
        cards.grid_columnconfigure((0, 1), weight=1, uniform="cards")

        status = self._card(cards, 0, "STATUS", padx=(0, 8))
        self.status_text = ctk.CTkLabel(status, text="Checking…", font=_font(20, "bold"), anchor="w")
        self.status_text.pack(fill="x", padx=18, pady=(0, 2))
        self.last_check_label = ctk.CTkLabel(
            status, text="Last check: never", font=_font(12), text_color=MUTED, anchor="w"
        )
        self.last_check_label.pack(fill="x", padx=18)
        self.next_check_label = ctk.CTkLabel(status, text="", font=_font(12), text_color=MUTED, anchor="w")
        self.next_check_label.pack(fill="x", padx=18, pady=(0, 5))
        self.countdown_bar = ctk.CTkProgressBar(
            status,
            height=8,
            fg_color=SUBTLE_HOVER,
            progress_color=ACCENT,
        )
        self.countdown_bar.set(0)
        self.countdown_bar.pack(fill="x", padx=18, pady=(0, 8))
        self.progress_area = ctk.CTkFrame(status, fg_color="transparent")
        self.progress = ctk.CTkProgressBar(
            self.progress_area,
            height=18,
            fg_color=ACCENT_HOVER,
            progress_color=ACCENT,
        )
        self.progress.set(0)
        self.progress.pack(fill="x", padx=18)
        self.progress_text = ctk.CTkLabel(
            self.progress_area,
            text="",
            font=_font(10, "bold"),
            text_color="white",
            fg_color="transparent",
        )
        self.progress_text.place(relx=0.5, rely=0.5, anchor="center")
        self.progress_area.pack_forget()
        controls = ctk.CTkFrame(status, fg_color="transparent")
        controls.pack(fill="x", padx=18, pady=(8, 10))
        controls.grid_columnconfigure(0, weight=1)
        self.check_button = ctk.CTkButton(
            controls,
            text="Check for requested videos now",
            font=_font(13, "bold"),
            height=38,
            corner_radius=10,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=self.check_now,
        )
        self.check_button.grid(row=0, column=0, sticky="ew")
        self.pause_switch = ctk.CTkSwitch(
            controls,
            text="Automatic checks",
            font=_font(12),
            progress_color=ACCENT,
            command=self._toggle_pause,
        )
        self.pause_switch.select()
        self.pause_switch.grid(row=1, column=0, sticky="e", pady=(8, 0))

        uploads = self._card(cards, 1, "UPLOADS", padx=(8, 0))
        ctk.CTkLabel(
            uploads, text="Videos uploaded from this computer", font=_font(12), text_color=MUTED, anchor="w"
        ).pack(fill="x", padx=18, pady=(0, 0))
        self.upload_count = ctk.CTkLabel(uploads, text="0", font=_font(34, "bold"), anchor="w")
        self.upload_count.pack(fill="x", padx=18, pady=(0, 0))
        self.last_title = ctk.CTkLabel(
            uploads, text="No uploads yet", font=_font(13), anchor="w", justify="left", wraplength=380
        )
        self.last_title.pack(fill="x", padx=18, pady=(6, 8))
        link_row = ctk.CTkFrame(uploads, fg_color="transparent")
        link_row.pack(fill="x", padx=18, pady=(0, 10))
        self.open_studio_button = self._secondary_button(
            link_row, "Open in YouTube Studio", self.open_last_in_studio, width=170
        )
        self.open_studio_button.configure(state="disabled")
        icon = Image.new("RGBA", (30, 22), (0, 0, 0, 0))
        drawing = ImageDraw.Draw(icon)
        drawing.rounded_rectangle((1, 2, 28, 20), radius=5, fill="white")
        drawing.polygon(((12, 6), (12, 16), (20, 11)), fill=YOUTUBE_RED)
        self.youtube_icon = ctk.CTkImage(light_image=icon, dark_image=icon, size=(24, 18))
        self.open_link_button = ctk.CTkButton(
            link_row,
            text="Open on YouTube",
            image=self.youtube_icon,
            compound="left",
            command=self.open_last_link,
            height=36,
            corner_radius=10,
            font=_font(13, "bold"),
            fg_color=YOUTUBE_RED,
            hover_color="#cc0000",
            text_color="white",
        )
        self.open_link_button.configure(state="disabled")
        self.copy_link_button = self._secondary_button(
            link_row, "Copy YouTube URL", self.copy_last_link, width=145
        )
        self.copy_link_button.configure(state="disabled")
        self.link_buttons = [self.open_studio_button, self.open_link_button, self.copy_link_button]
        link_row.bind("<Configure>", self._layout_link_buttons)
        self._layout_link_buttons()
        self.upload_instruction_frame = ctk.CTkFrame(uploads, fg_color=("#edf8f1", "#173b2a"), corner_radius=10)
        self.upload_instruction = ctk.CTkLabel(
            self.upload_instruction_frame,
            text="",
            font=_font(12),
            anchor="w",
            justify="left",
            wraplength=400,
        )
        self.upload_instruction.pack(fill="x", padx=12, pady=10)
        self.upload_instruction_frame.pack_forget()

        log_card = ctk.CTkFrame(root, fg_color=CARD, corner_radius=16)
        log_card.grid(row=5, column=0, sticky="nsew", padx=24, pady=(6, 16))
        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.pack(fill="x", padx=18, pady=(10, 4))
        ctk.CTkLabel(log_header, text="ACTIVITY", font=_font(11, "bold"), text_color=MUTED).pack(side="left")
        self._secondary_button(log_header, "Clear", self._clear_log, width=74, height=28).pack(side="right")
        self.log_view = ctk.CTkTextbox(
            log_card,
            height=160,
            font=_font(11, family="Consolas"),
            corner_radius=10,
            fg_color=SUBTLE,
            wrap="word",
        )
        self.log_view.pack(fill="both", expand=True, padx=18, pady=(0, 12))
        self.log_view.configure(state="disabled")
        self.debug_label = ctk.CTkLabel(
            root,
            text=f"Debug: app {auto_yt.APP_VERSION} · {sys.platform} · Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
            font=_font(10),
            text_color=MUTED,
            anchor="w",
        )
        self.debug_label.grid(row=6, column=0, sticky="ew", padx=28, pady=(0, 12))
        self._start_tray_icon()

    def _card(self, parent: ctk.CTkFrame, column: int, title: str, padx: tuple[int, int]) -> ctk.CTkFrame:
        card = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=16)
        card.grid(row=0, column=column, sticky="nsew", padx=padx)
        ctk.CTkLabel(card, text=title, font=_font(11, "bold"), text_color=MUTED, anchor="w").pack(
            fill="x", padx=18, pady=(8, 0)
        )
        return card

    def _secondary_button(
        self,
        parent: ctk.CTkFrame,
        text: str,
        command,
        width: int = 0,
        height: int = 36,
    ) -> ctk.CTkButton:
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=height,
            corner_radius=10,
            font=_font(13),
            fg_color="transparent",
            hover_color=("#edf8f1", "#173b2a"),
            border_width=1,
            border_color=ACCENT,
            text_color=(ACCENT, "#74d99b"),
        )

    # Status and timing

    def _set_status(self, kind: str, text: str) -> None:
        self.pill.configure(text=text, fg_color=STATUS_COLORS.get(kind, ERROR_RED))
        self.status_text.configure(text=text)

    def _latest_features(self) -> str:
        changelog = _read_changelog()
        sections = changelog.split("\n## ")
        if len(sections) < 2:
            return "No release notes are available."
        heading, _, notes = sections[1].partition("\n")
        version, separator, release_date = heading.partition(" - ")
        return _format_features(
            version.strip(),
            notes.strip(),
            release_date.strip() if separator else "",
        )

    def _on_status_click(self) -> None:
        if self.available_update:
            self.status_menu.grid_remove()
            visible = bool(self.update_bar.winfo_ismapped())
            if visible:
                self.update_bar.grid_remove()
                self.features_panel.grid_remove()
            else:
                self.update_bar.grid()
                self.features_panel.grid()
                self.features_expanded = True
        elif self.pill.cget("text") in {"Up to date", "Update check failed"}:
            self.status_menu.grid() if not self.status_menu.winfo_ismapped() else self.status_menu.grid_remove()

    def _handle_status_menu(self, choice: str) -> None:
        self.status_menu.grid_remove()
        if choice == "Check for client updates":
            self.check_updates_only()
        elif choice == "Latest features":
            self.features_expanded = not self.features_expanded
            if self.features_expanded:
                self.features_panel.grid()
            else:
                self.features_panel.grid_remove()

    def _resize_content(self, event: tk.Event) -> None:
        if event.widget is not self.root:
            return
        width = max(360, (event.width - 96) // 2)
        self.last_title.configure(wraplength=width - 40)
        self.upload_instruction.configure(wraplength=width - 44)
        self.features_text.configure(wraplength=event.width - 96)
        self.update_warning.configure(wraplength=event.width - 72)
        self.update_notes.configure(wraplength=event.width - 72)

    def _layout_link_buttons(self, _event: tk.Event | None = None) -> None:
        width = self.link_buttons[0].master.winfo_width()
        if width <= 1:
            width = 460
        available = max(1, width - 8)
        widths = [180, 160, 155]
        rows: list[list[int]]
        if sum(widths) + 16 <= available:
            rows = [[0, 1, 2]]
        elif widths[0] + widths[1] + 8 <= available:
            rows = [[0, 1], [2]]
        else:
            rows = [[0], [1], [2]]
        for button in self.link_buttons:
            button.grid_forget()
        for row, indices in enumerate(rows):
            for column, index in enumerate(indices):
                self.link_buttons[index].grid(
                    row=row,
                    column=column,
                    sticky="w",
                    padx=(0, 8) if column < len(indices) - 1 else 0,
                    pady=(0, 6) if row < len(rows) - 1 else 0,
                )

    def _start_tray_icon(self) -> None:
        if self.smoke or os.name != "nt":
            return
        image = Image.new("RGBA", (64, 64), (0, 123, 61, 255))
        drawing = ImageDraw.Draw(image)
        drawing.rounded_rectangle((8, 15, 56, 49), radius=10, fill=(255, 0, 0, 255))
        drawing.polygon(((28, 22), (28, 42), (43, 32)), fill=(255, 255, 255, 255))
        menu = pystray.Menu(
            pystray.MenuItem("Open Auto YouTube Uploader", self._restore_from_tray, default=True),
            pystray.MenuItem("Exit", self._exit_from_tray),
        )
        self.tray_icon = pystray.Icon(
            "AutoYouTubeUploader",
            image,
            "Auto YouTube Uploader",
            menu,
        )
        self.tray_icon.run_detached()

    def _on_minimize(self, _event: tk.Event) -> None:
        if self.smoke or not self.tray_icon:
            return

        def hide_minimized() -> None:
            if self.root.state() == "iconic":
                self.root.withdraw()

        self.root.after(100, hide_minimized)

    def _restore_from_tray(self, _icon=None, _item=None) -> None:
        self.root.after(0, self._show_window)

    def _show_window(self) -> None:
        self.root.deiconify()
        self.root.state("normal")
        self.root.lift()
        self.root.focus_force()

    def _exit_from_tray(self, _icon=None, _item=None) -> None:
        self.root.after(0, self._close)

    def _stop_tray_icon(self) -> None:
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None

    def _check_interrupted_update(self) -> None:
        marker = _update_marker_path()
        if not marker.is_file():
            return
        try:
            reason = marker.read_text(encoding="utf-8").strip()
            marker.unlink()
        except OSError as exc:
            self._log(f"Could not read the previous update status: {exc}")
            return
        if reason.startswith("failed:"):
            message = reason.partition(":")[2].strip()
        else:
            message = "The update did not finish. The CMD window may have been closed before installation completed."
        messagebox.showerror(
            "Update was not completed",
            f"{message}\n\nCheck for updates and try again.",
            parent=self.root,
        )

    def _tick(self) -> None:
        due = self.next_check_at > 0 and time.time() >= self.next_check_at
        if due and not self.busy and not self.paused and not self.smoke:
            self.check_now()
        self._update_countdown()
        self.root.after(1000, self._tick)

    def _update_countdown(self) -> None:
        if self.busy:
            self.countdown_bar.set(0)
            return
        if self.paused:
            self.next_check_label.configure(text="Automatic checks are paused")
            self.countdown_bar.set(0)
            return
        remaining = max(0, int(self.next_check_at - time.time()))
        minutes, seconds = divmod(remaining, 60)
        self.next_check_label.configure(text=f"Next automatic check in {minutes}m {seconds:02d}s")
        elapsed = self.interval_minutes * 60 - remaining
        self.countdown_bar.set(min(1.0, max(0.0, elapsed / (self.interval_minutes * 60))))

    def _toggle_pause(self) -> None:
        self.paused = not bool(self.pause_switch.get())
        if self.paused:
            self._log("Automatic checks paused.")
        else:
            self._log("Automatic checks resumed.")
            if not self.busy:
                self.next_check_at = time.time() + self.interval_minutes * 60
        self._update_countdown()

    # Checks and updates

    def _reset_steps(self) -> None:
        self.current_stage = "checking"

    def _show_stage(self, name: str, detail: str, fraction: float | None) -> None:
        self.current_stage = name
        label = dict(STAGE_STEPS).get(name, "Working")
        text = f"{label} · {detail}" if detail and detail.casefold() != label.casefold() else label
        if fraction is not None:
            text += f" ({int(fraction * 100)}%)"
        self.progress_text.configure(text=text)
        if not self.progress_area.winfo_manager():
            self.progress_area.pack(fill="x", pady=(8, 0))
        if fraction is None:
            self.progress.configure(mode="indeterminate")
            self.progress.start()
        else:
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress.set(fraction)
            self.progress_text.configure(text=f"{text} ({int(fraction * 100)}%)")

    def check_now(self) -> None:
        if self.busy:
            return
        self.busy = True
        self._reset_steps()
        self._set_status("busy", "Checking")
        self.check_button.configure(state="disabled", text="Checking…")
        self.progress_text.configure(text="Checking workbook")
        self.progress_area.pack(fill="x", pady=(8, 0))
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self._log("Checking the workbook for new requests…")
        threading.Thread(target=self._run_check, daemon=True).start()

    def _queue_stage(self, name: str, detail: str, fraction: float | None) -> None:
        self.events.put(("stage", (name, detail, fraction)))

    def _run_check(self) -> None:
        output = io.StringIO()
        result: dict[str, object] = {"ok": True, "uploaded": False, "error": ""}
        try:
            with contextlib.redirect_stdout(output), contextlib.redirect_stderr(output):
                auto_yt.main(["--yes"], on_stage=self._queue_stage)
            result["uploaded"] = "Upload complete" in output.getvalue()
        except Exception as exc:
            result = {"ok": False, "uploaded": False, "error": str(exc)}
        text = output.getvalue().strip()
        if text:
            self.events.put(("log", text))
        if not result["ok"]:
            self.events.put(("log", f"Error: {result['error']}"))
        self._run_update_check()
        self.events.put(("done", result))

    def check_updates_only(self) -> None:
        if self.busy:
            return
        self._log("Checking for client updates…")
        self.busy = True
        self.check_button.configure(state="disabled")
        self._set_status("busy", "Checking updates")
        self._show_progress("Checking for updates", None)
        threading.Thread(target=self._run_update_check, args=(True,), daemon=True).start()

    def _run_update_check(self, standalone: bool = False) -> None:
        try:
            update = check_update(auto_yt.DEFAULT_UPDATE_MANIFEST_URL, auto_yt.APP_VERSION)
            if update:
                note = f"Installed {auto_yt.APP_VERSION} · version {update['version']} available"
            else:
                note = f"Installed {auto_yt.APP_VERSION} · no newer version"
            self.events.put(("update", (update, note)))
        except Exception as exc:
            self.events.put(("log", f"Update check failed: {exc}"))
            self.events.put(("update", (None, "Update check failed; see activity log")))
        if standalone:
            self.events.put(("update_check_done", None))

    def _set_version_indicator(self) -> None:
        if self.update_check_error:
            self._set_status("error", "Update check failed")
        elif self.available_update:
            self._set_status("update_available", "Update available")
        else:
            self._set_status("ready", "Up to date")

    def _apply_update(self, payload: tuple[dict[str, str] | None, str]) -> None:
        update, note = payload
        self.version_note = note
        self.update_label_note = note
        self.update_check_error = update is None and "update check failed" in note.casefold()
        if update:
            self.available_update = update
            staged = self._staged_update_dir(update["version"])
            self.downloaded_update = staged if (staged / "AutoYouTubeUploader.exe").is_file() else None
            self.update_label.configure(
                text=f"Version {update['version']} is available"
            )
            self.download_button.configure(state="normal" if not self.downloaded_update else "disabled")
            self.install_button.configure(state="normal" if self.downloaded_update else "disabled")
            self.download_install_button.configure(state="normal" if not self.downloaded_update else "disabled")
            release_notes = update.get("notes", "")
            self.features_text.configure(
                text=_format_features(update["version"], release_notes)
                if release_notes
                else self._latest_features()
            )
            self.update_notes.configure(
                text=_format_features(update["version"], release_notes)
                if release_notes
                else self._latest_features()
            )
            if self.announced_update != update["version"]:
                self.announced_update = update["version"]
                self._log(f"Client update {update['version']} is available.")
        else:
            self.available_update = None
            self.downloaded_update = None
            self.features_text.configure(text=self._latest_features())
            self.download_button.configure(state="disabled")
            self.install_button.configure(state="disabled")
            self.download_install_button.configure(state="disabled")
            self.update_bar.grid_remove()
            self.features_panel.grid_remove()
            self.features_expanded = False
        self.last_check_label.configure(text=f"{self.last_check_label.cget('text').split('  ·  ')[0]}  ·  {note}")

    def _finish_check(self, result: dict) -> None:
        self.busy = False
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(0)
        self.progress_area.pack_forget()
        self.check_button.configure(state="normal", text="Check for requested videos now")
        self.last_check_label.configure(text=f"Last check: {datetime.now():%d %b %Y, %H:%M}")
        self.next_check_at = time.time() + self.interval_minutes * 60
        self._refresh_uploads()
        if not result["ok"]:
            failed_step = dict(STAGE_STEPS).get(self.current_stage, "start")
            self._set_status("error", f"Failed: {failed_step.lower()}")
        else:
            self._set_version_indicator()
        if result["ok"] and result["uploaded"]:
            self._show_upload_instruction()
            self._notify_upload()

    def _staged_update_dir(self, version: str) -> Path:
        base = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "AutoYouTubeUploader" / "updates"
        return base / f"AutoYouTubeUploader-{version}"

    def download_available_update(self) -> None:
        self._download_available_update(install_after=False)

    def download_and_install_update(self) -> None:
        self._download_available_update(install_after=True)

    def _download_available_update(self, install_after: bool) -> None:
        if not self.available_update or self.busy:
            return
        self.busy = True
        self.check_button.configure(state="disabled")
        self.install_after_download = install_after
        self._set_status("busy", "Downloading update")
        self._set_update_buttons("disabled")
        self._show_progress("Downloading update", None)
        self._log(f"Downloading client update {self.available_update['version']}…")
        threading.Thread(
            target=self._download_update_worker,
            args=(self.available_update, install_after),
            daemon=True,
        ).start()

    def _download_update_worker(self, update: dict[str, str], install_after: bool) -> None:
        try:
            staged = download_update(
                update,
                on_progress=lambda written, total: self.events.put(("update_progress", (written, total))),
            )
            self.events.put(("update_downloaded", (staged, install_after)))
        except Exception as exc:
            self.events.put(("log", f"Update download failed: {exc}"))
            self.events.put(("update_download_failed", str(exc)))

    def _update_download_progress(self, written: int, total: int | None) -> None:
        detail = f"Downloading update · {written / 1048576:,.0f} MB"
        if total:
            detail += f" / {total / 1048576:,.0f} MB"
            fraction = min(1.0, written / total)
        else:
            fraction = None
        self._show_progress(detail, fraction)

    def _finish_update_download(self, staged: Path, install_after: bool) -> None:
        self.busy = False
        self.downloaded_update = staged
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(1)
        self.progress_area.pack_forget()
        self.check_button.configure(state="normal")
        self._set_status("ready", "Update ready to install")
        if self.available_update:
            self.update_label.configure(text=f"Version {self.available_update['version']} downloaded and verified")
        self._restore_update_buttons()
        self.install_button.configure(state="normal")
        self._log("Update downloaded and verified. Install it when ready.")
        if install_after:
            self._start_update_install(ask_confirmation=False)

    def _finish_update_download_error(self, error: str) -> None:
        self.busy = False
        self.progress.stop()
        self.progress_area.pack_forget()
        self.check_button.configure(state="normal")
        self._set_status("error", "Update download failed")
        self._restore_update_buttons()
        self._log(f"Update download error: {error}")
        messagebox.showerror("Could not download update", error, parent=self.root)

    def _set_update_buttons(self, state: str) -> None:
        self.download_button.configure(state=state)
        self.install_button.configure(state=state)
        self.download_install_button.configure(state=state)

    def _restore_update_buttons(self) -> None:
        if not self.available_update:
            self._set_update_buttons("disabled")
            return
        is_downloaded = self.downloaded_update is not None
        self.download_button.configure(state="disabled" if is_downloaded else "normal")
        self.install_button.configure(state="normal" if is_downloaded else "disabled")
        self.download_install_button.configure(state="disabled" if is_downloaded else "normal")

    def _show_progress(self, text: str, fraction: float | None) -> None:
        self.progress_text.configure(text=text)
        if not self.progress_area.winfo_manager():
            self.progress_area.pack(fill="x", pady=(8, 0))
        if fraction is None:
            self.progress.stop()
            self.progress.configure(mode="indeterminate")
            self.progress.start()
        else:
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress.set(fraction)

    def install_available_update(self) -> None:
        self._start_update_install(ask_confirmation=True)

    def _start_update_install(self, ask_confirmation: bool) -> None:
        if not self.downloaded_update or self.busy:
            return
        if ask_confirmation and not messagebox.askyesno(
            "Install update",
            "The verified update is ready.\n\nThe client will close and a CMD window will open. "
            "Do not close that window while it is running.",
            parent=self.root,
        ):
            return
        self.busy = True
        self._set_status("busy", "Installing update")
        self._set_update_buttons("disabled")
        self._show_progress("Starting installer", None)
        threading.Thread(target=self._install_update_worker, args=(self.downloaded_update,), daemon=True).start()

    def _install_update_worker(self, staged: Path) -> None:
        try:
            install_update(staged)
            self.events.put(("log", "Installer CMD opened; leaving it open while it updates the client."))
            self.events.put(("exit", None))
        except Exception as exc:
            self.events.put(("log", f"Update failed: {exc}"))
            self.events.put(("install_failed", str(exc)))

    # Uploads, settings and files

    def _refresh_uploads(self) -> None:
        try:
            state = auto_yt._read_state(auto_yt._state_path())
        except (OSError, ValueError) as exc:
            self._log(f"Could not read upload history: {exc}")
            return
        entries = [value for value in state.values() if isinstance(value, dict)]
        self.upload_count.configure(text=str(len(entries)))
        if entries:
            latest = entries[-1]
            self.last_link = str(latest.get("youtube_link") or "") or None
            self.last_title.configure(text=str(latest.get("title") or "Untitled video"))
            self.last_playlist = str(latest.get("playlist") or "").strip()
        else:
            self.last_link = None
            self.last_title.configure(text="No uploads yet")
            self.last_playlist = ""
        link_state = "normal" if self.last_link else "disabled"
        self.open_studio_button.configure(state=link_state)
        self.open_link_button.configure(state=link_state)
        self.copy_link_button.configure(state=link_state)

    def _show_upload_instruction(self) -> None:
        playlist = self.last_playlist or "the playlist listed in the Excel document (if one is selected)"
        self.upload_instruction.configure(
            text=(
                "The upload is done. Please go to the uploaded video in YouTube Studio, check that it looks good, "
                f"add the playlist: {playlist}, and make the video public."
            )
        )
        if not self.upload_instruction_frame.winfo_manager():
            self.upload_instruction_frame.pack(fill="x", padx=18, pady=(0, 12))

    def _notify_upload(self) -> None:
        if not self.tray_icon:
            return
        try:
            title = self.last_title.cget("text")
            self.tray_icon.notify(f"Upload complete: {title}", "Auto YouTube Uploader")
        except Exception as exc:
            self._log(f"Could not show the Windows upload notification: {exc}")

    def open_last_link(self) -> None:
        if self.last_link:
            webbrowser.open(self.last_link)

    def open_last_in_studio(self) -> None:
        if not self.last_link:
            return
        studio_url = _youtube_studio_url(self.last_link)
        if studio_url:
            webbrowser.open(studio_url)
        else:
            messagebox.showerror(
                "Could not open YouTube Studio",
                "The saved YouTube URL does not contain a recognizable video ID.",
                parent=self.root,
            )

    def copy_last_link(self) -> None:
        if self.last_link:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.last_link)
            self._log("Copied the YouTube link to the clipboard.")

    def open_workbook(self) -> None:
        self._open_target(self._config().get("workbook", str(auto_yt.DEFAULT_WORKBOOK)))

    def open_video_folder(self) -> None:
        self._open_target(self._config().get("local_video_dir", str(auto_yt.DEFAULT_LOCAL_VIDEO_DIR)))

    def _open_target(self, target: str) -> None:
        try:
            _open_path(Path(target))
        except OSError as exc:
            messagebox.showerror("Could not open", f"{target}\n\n{exc}", parent=self.root)

    def _config(self) -> dict[str, str]:
        return auto_yt._read_config(auto_yt._config_path())

    def open_settings(self) -> None:
        if self.busy:
            messagebox.showwarning("Please wait", "A check is running. Try again when it completes.", parent=self.root)
            return
        config = show_settings(self.root, required=False)
        if config:
            self.interval_minutes = interval_minutes(config)
            self.next_check_at = time.time() + self.interval_minutes * 60
            self._log("Settings saved.")
            self.check_now()

    # Activity log and event loop

    def _log(self, text: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        lines = [line for line in text.splitlines() if line.strip()]
        if not lines:
            return
        block = "\n".join(f"[{stamp}] {line}" for line in lines) + "\n"
        self.log_lines.extend(block.rstrip().splitlines())
        self.log_view.configure(state="normal")
        if len(self.log_lines) > 6000:
            self.log_lines = self.log_lines[-5000:]
            self.log_view.delete("1.0", "end")
            self.log_view.insert("end", "\n".join(self.log_lines) + "\n")
        else:
            self.log_view.insert("end", block)
        self.log_view.see("end")
        self.log_view.configure(state="disabled")
        with contextlib.suppress(OSError):
            log_path = _log_path()
            log_path.parent.mkdir(parents=True, exist_ok=True)
            with log_path.open("a", encoding="utf-8") as stream:
                stream.write(block)

    def _clear_log(self) -> None:
        self.log_lines.clear()
        self.log_view.configure(state="normal")
        self.log_view.delete("1.0", "end")
        self.log_view.configure(state="disabled")

    def _process_events(self) -> None:
        while True:
            try:
                kind, value = self.events.get_nowait()
            except queue.Empty:
                break
            if kind == "log":
                self._log(str(value))
            elif kind == "update":
                self._apply_update(value)  # type: ignore[arg-type]
            elif kind == "stage":
                self._show_stage(*value)  # type: ignore[misc]
            elif kind == "done":
                self._finish_check(value)  # type: ignore[arg-type]
            elif kind == "update_progress":
                self._update_download_progress(*value)  # type: ignore[arg-type]
            elif kind == "update_downloaded":
                self._finish_update_download(*value)  # type: ignore[arg-type]
            elif kind == "update_download_failed":
                self._finish_update_download_error(str(value))
            elif kind == "update_check_done":
                self.busy = False
                self.progress.stop()
                self.progress_area.pack_forget()
                self.check_button.configure(state="normal")
                self._set_version_indicator()
            elif kind == "install_failed":
                self.busy = False
                self.progress.stop()
                self.progress_area.pack_forget()
                self._set_status("error", "Update failed")
                self._restore_update_buttons()
                messagebox.showerror("Could not start update", str(value), parent=self.root)
            elif kind == "exit":
                self._stop_tray_icon()
                self.root.destroy()
                return
        self.root.after(150, self._process_events)

    def _close(self) -> None:
        if self.busy:
            messagebox.showwarning(
                "Please wait",
                "A workbook check or update is still running. You can minimize the client to the system tray.",
                parent=self.root,
            )
            return
        self._stop_tray_icon()
        self.root.destroy()


def run_client() -> int:
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    try:
        config = auto_yt._read_config(auto_yt._config_path())
        if not config:
            # Keep the root visible so the first-run dialog cannot open hidden behind it.
            root.deiconify()
            config = show_settings(root, required=True)
            if config is None:
                root.destroy()
                return 0
        root.deiconify()
        app = AutoYouTubeApp(root, config)
        marker = os.environ.get("AUTOYT_GUI_SMOKE_MARKER")
        if marker:
            def finish_smoke_test() -> None:
                root.update()
                Path(marker).write_text(
                    f"title={root.title()}\n"
                    f"mapped={root.winfo_ismapped()}\n"
                    f"scrollable_page={isinstance(app.content, ctk.CTkScrollableFrame)}\n"
                    f"scrollable_activity={isinstance(app.log_view, ctk.CTkTextbox)}\n"
                    f"check_in_status={app.check_button.master.master is app.status_text.master}\n"
                    f"studio_button_first={app.link_buttons[0].cget('text') == 'Open in YouTube Studio'}\n",
                    encoding="utf-8",
                )
                root.destroy()

            root.after(1500, finish_smoke_test)
        root.mainloop()
        return 0
    except Exception as exc:
        messagebox.showerror("Auto YouTube Uploader", str(exc), parent=root)
        root.destroy()
        return 1
