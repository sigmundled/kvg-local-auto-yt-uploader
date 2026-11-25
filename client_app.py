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

import customtkinter as ctk

import auto_yt
from updater import check_update, download_update, install_update

DEFAULT_INTERVAL_MINUTES = 60
MIN_INTERVAL_MINUTES = 5
CARD = ("#ffffff", "#1d212b")
SUBTLE = ("#eef1f6", "#2a303d")
SUBTLE_HOVER = ("#e0e5ee", "#353c4b")
MUTED = ("#5f6b7a", "#9aa4b5")
TEXT = ("#15181f", "#eef1f6")
ACCENT = "#e62117"
ACCENT_HOVER = "#b91c14"
STATUS_COLORS = {
    "ready": "#16a34a",
    "busy": "#d97706",
    "error": "#dc2626",
}
# (config key, label, browse kind, dialog title); kind is None for plain text fields.
FIELD_SPECS = [
    ("workbook", "Request workbook (.xlsx)", "xlsx", "Select the request workbook"),
    ("local_video_dir", "Synced Video File folder", "dir", "Select the synced Video File folder"),
    ("client_secrets", "Google Desktop OAuth client JSON", "json", "Select the Google OAuth client JSON"),
    ("update_manifest_url", "Update manifest URL (https, optional)", None, ""),
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
            "update_manifest_url": existing.get("update_manifest_url") or auto_yt.DEFAULT_UPDATE_MANIFEST_URL,
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
            text="Quit" if required else "Cancel",
            font=_font(14),
            height=40,
            corner_radius=10,
            fg_color=SUBTLE,
            hover_color=SUBTLE_HOVER,
            text_color=TEXT,
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
                    fg_color=SUBTLE,
                    hover_color=SUBTLE_HOVER,
                    text_color=TEXT,
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
        manifest = values["update_manifest_url"]
        interval_text = values["check_interval_minutes"]

        errors: list[str] = []
        if not workbook.is_file():
            errors.append("The request workbook was not found.")
        if not video_dir.is_dir():
            errors.append("The Video File folder was not found.")
        if not _is_desktop_oauth_file(client_secrets):
            errors.append("The Google OAuth file must be a Desktop application client JSON.")
        if manifest and not manifest.lower().startswith("https://"):
            errors.append("The update manifest URL must start with https://.")
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
        if manifest:
            config["update_manifest_url"] = manifest
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
        self.announced_update = ""
        self.last_link: str | None = None
        self.interval_minutes = interval_minutes(config)
        self.next_check_at = 0.0
        self.smoke = bool(os.environ.get("AUTOYT_GUI_SMOKE_MARKER"))

        root.title(f"Auto YouTube Uploader {auto_yt.APP_VERSION}")
        root.geometry("1040x780")
        root.minsize(900, 700)
        root.protocol("WM_DELETE_WINDOW", self._close)
        self._build()
        self._refresh_uploads()
        self._set_status("ready", "Ready")
        if not self.smoke:
            self.next_check_at = time.time() + self.interval_minutes * 60
            root.after(400, self.check_now)
        root.after(150, self._process_events)
        root.after(1000, self._tick)

    # Layout

    def _build(self) -> None:
        root = self.root
        root.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(4, weight=1)

        header = ctk.CTkFrame(root, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=28, pady=(22, 6))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Auto YouTube Uploader", font=_font(26, "bold"), anchor="w").grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkLabel(
            header,
            text=f"Version {auto_yt.APP_VERSION}  ·  Settings: {auto_yt._config_path()}",
            font=_font(12),
            text_color=MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="w")
        self.pill = ctk.CTkLabel(
            header,
            text="Ready",
            height=32,
            corner_radius=16,
            padx=18,
            fg_color=STATUS_COLORS["ready"],
            text_color="white",
            font=_font(13, "bold"),
        )
        self.pill.grid(row=0, column=1, rowspan=2, sticky="e")

        self.update_bar = ctk.CTkFrame(root, fg_color=("#dbeafe", "#1e3a5f"), corner_radius=14)
        self.update_label = ctk.CTkLabel(self.update_bar, text="", font=_font(13), anchor="w", justify="left")
        self.update_label.pack(side="left", padx=18, pady=12, fill="x", expand=True)
        self.install_button = ctk.CTkButton(
            self.update_bar,
            text="Install and restart",
            width=180,
            height=36,
            corner_radius=10,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            font=_font(13, "bold"),
            command=self.install_available_update,
            state="disabled",
        )
        self.install_button.pack(side="right", padx=14, pady=10)
        self.update_bar.grid(row=1, column=0, sticky="ew", padx=28, pady=(6, 0))
        self.update_bar.grid_remove()

        hero = ctk.CTkFrame(root, fg_color=CARD, corner_radius=18)
        hero.grid(row=2, column=0, sticky="ew", padx=28, pady=8)
        top = ctk.CTkFrame(hero, fg_color="transparent")
        top.pack(fill="x", padx=22, pady=(20, 10))
        self.check_button = ctk.CTkButton(
            top,
            text="Check for requested videos now",
            font=_font(15, "bold"),
            height=50,
            corner_radius=12,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=self.check_now,
        )
        self.check_button.pack(side="left")
        self.pause_switch = ctk.CTkSwitch(
            top,
            text="Automatic checks",
            font=_font(13),
            progress_color=ACCENT,
            command=self._toggle_pause,
        )
        self.pause_switch.select()
        self.pause_switch.pack(side="right")

        actions = ctk.CTkFrame(hero, fg_color="transparent")
        actions.pack(fill="x", padx=22, pady=(0, 20))
        secondary = [
            ("Open workbook", self.open_workbook),
            ("Open video folder", self.open_video_folder),
            ("Settings", self.open_settings),
            ("Check for updates", self.check_updates_only),
            ("Changelog", self.open_changelog),
        ]
        for text, command in secondary:
            self._secondary_button(actions, text, command).pack(side="left", padx=(0, 10))
        ctk.CTkButton(
            actions,
            text="Quit",
            font=_font(13),
            height=36,
            corner_radius=10,
            fg_color="transparent",
            border_width=1,
            border_color=MUTED,
            text_color=TEXT,
            hover_color=SUBTLE,
            command=self._close,
        ).pack(side="right")

        cards = ctk.CTkFrame(root, fg_color="transparent")
        cards.grid(row=3, column=0, sticky="ew", padx=28, pady=8)
        cards.grid_columnconfigure((0, 1), weight=1, uniform="cards")

        status = self._card(cards, 0, "STATUS", padx=(0, 8))
        self.status_text = ctk.CTkLabel(status, text="Starting…", font=_font(20, "bold"), anchor="w")
        self.status_text.pack(fill="x", padx=20, pady=(4, 2))
        self.last_check_label = ctk.CTkLabel(
            status, text="Last check: never", font=_font(13), text_color=MUTED, anchor="w"
        )
        self.last_check_label.pack(fill="x", padx=20)
        self.next_check_label = ctk.CTkLabel(status, text="", font=_font(13), text_color=MUTED, anchor="w")
        self.next_check_label.pack(fill="x", padx=20, pady=(0, 10))
        self.progress = ctk.CTkProgressBar(status, height=8, progress_color=ACCENT)
        self.progress.set(0)
        self.progress.pack(fill="x", padx=20, pady=(0, 18))

        uploads = self._card(cards, 1, "UPLOADS", padx=(8, 0))
        self.upload_count = ctk.CTkLabel(uploads, text="0", font=_font(34, "bold"), anchor="w")
        self.upload_count.pack(fill="x", padx=20, pady=(0, 0))
        ctk.CTkLabel(
            uploads, text="videos uploaded from this computer", font=_font(12), text_color=MUTED, anchor="w"
        ).pack(fill="x", padx=20)
        self.last_title = ctk.CTkLabel(
            uploads, text="No uploads yet", font=_font(13), anchor="w", justify="left", wraplength=380
        )
        self.last_title.pack(fill="x", padx=20, pady=(10, 8))
        link_row = ctk.CTkFrame(uploads, fg_color="transparent")
        link_row.pack(fill="x", padx=20, pady=(0, 18))
        self.open_link_button = self._secondary_button(link_row, "Open on YouTube", self.open_last_link)
        self.open_link_button.configure(state="disabled")
        self.open_link_button.pack(side="left")
        self.copy_link_button = self._secondary_button(link_row, "Copy link", self.copy_last_link)
        self.copy_link_button.configure(state="disabled")
        self.copy_link_button.pack(side="left", padx=(10, 0))

        log_card = ctk.CTkFrame(root, fg_color=CARD, corner_radius=16)
        log_card.grid(row=4, column=0, sticky="nsew", padx=28, pady=(8, 24))
        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.pack(fill="x", padx=20, pady=(14, 6))
        ctk.CTkLabel(log_header, text="ACTIVITY", font=_font(11, "bold"), text_color=MUTED).pack(side="left")
        self._secondary_button(log_header, "Open log file", self.open_log_file, width=120, height=28).pack(
            side="right"
        )
        self._secondary_button(log_header, "Clear view", self._clear_log, width=100, height=28).pack(
            side="right", padx=8
        )
        self.log_box = ctk.CTkTextbox(
            log_card,
            font=_font(12, family="Consolas"),
            wrap="word",
            corner_radius=10,
            fg_color=SUBTLE,
            state="disabled",
        )
        self.log_box.pack(fill="both", expand=True, padx=20, pady=(0, 18))

    def _card(self, parent: ctk.CTkFrame, column: int, title: str, padx: tuple[int, int]) -> ctk.CTkFrame:
        card = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=16)
        card.grid(row=0, column=column, sticky="nsew", padx=padx)
        ctk.CTkLabel(card, text=title, font=_font(11, "bold"), text_color=MUTED, anchor="w").pack(
            fill="x", padx=20, pady=(16, 6)
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
            fg_color=SUBTLE,
            hover_color=SUBTLE_HOVER,
            text_color=TEXT,
        )

    # Status and timing

    def _set_status(self, kind: str, text: str) -> None:
        self.pill.configure(text=text, fg_color=STATUS_COLORS[kind])
        self.status_text.configure(text=text)

    def _tick(self) -> None:
        due = self.next_check_at > 0 and time.time() >= self.next_check_at
        if due and not self.busy and not self.paused and not self.smoke:
            self.check_now()
        self._update_countdown()
        self.root.after(1000, self._tick)

    def _update_countdown(self) -> None:
        if self.busy:
            return
        if self.paused:
            self.next_check_label.configure(text="Automatic checks are paused")
            self.progress.set(0)
            return
        remaining = max(0, int(self.next_check_at - time.time()))
        minutes, seconds = divmod(remaining, 60)
        self.next_check_label.configure(text=f"Next automatic check in {minutes}m {seconds:02d}s")
        total = self.interval_minutes * 60
        self.progress.set(min(1.0, 1 - remaining / total))

    def _toggle_pause(self) -> None:
        self.paused = not bool(self.pause_switch.get())
        if self.paused:
            self._log("Automatic checks paused.")
        else:
            self._log("Automatic checks resumed.")
            if not self.busy:
                self.next_check_at = time.time() + self.interval_minutes * 60

    # Checks and updates

    def check_now(self) -> None:
        if self.busy:
            return
        self.busy = True
        self._set_status("busy", "Checking…")
        self.check_button.configure(state="disabled", text="Checking…")
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self._log("Checking the workbook for new requests…")
        threading.Thread(target=self._run_check, daemon=True).start()

    def _run_check(self) -> None:
        output = io.StringIO()
        result: dict[str, object] = {"ok": True, "uploaded": False, "error": ""}
        try:
            with contextlib.redirect_stdout(output), contextlib.redirect_stderr(output):
                auto_yt.main(["--yes"])
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
        threading.Thread(target=self._run_update_check, daemon=True).start()

    def _run_update_check(self) -> None:
        try:
            config = auto_yt._read_config(auto_yt._config_path())
            manifest = config.get("update_manifest_url") or auto_yt.DEFAULT_UPDATE_MANIFEST_URL
            if not manifest:
                self.events.put(("update", (None, "Updates are not configured")))
                return
            update = check_update(manifest, auto_yt.APP_VERSION)
            if update:
                note = f"Installed {auto_yt.APP_VERSION} · version {update['version']} available"
            else:
                note = f"Installed {auto_yt.APP_VERSION} · up to date"
            self.events.put(("update", (update, note)))
        except Exception as exc:
            self.events.put(("log", f"Update check failed: {exc}"))
            self.events.put(("update", (None, "Update check failed; see activity log")))

    def _apply_update(self, payload: tuple[dict[str, str] | None, str]) -> None:
        update, note = payload
        self.version_note = note
        self.update_label_note = note
        if update:
            self.available_update = update
            self.update_label.configure(
                text=f"Version {update['version']} is ready to install.  {update.get('notes', '')}".strip()
            )
            self.install_button.configure(state="normal")
            self.update_bar.grid()
            if self.announced_update != update["version"]:
                self.announced_update = update["version"]
                self._log(f"Client update {update['version']} is available.")
        else:
            self.available_update = None
            self.install_button.configure(state="disabled")
            self.update_bar.grid_remove()
        self.last_check_label.configure(text=f"{self.last_check_label.cget('text').split('  ·  ')[0]}  ·  {note}")

    def _finish_check(self, result: dict) -> None:
        self.busy = False
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(0)
        self.check_button.configure(state="normal", text="Check for requested videos now")
        self.last_check_label.configure(text=f"Last check: {datetime.now():%d %b %Y, %H:%M}")
        self.next_check_at = time.time() + self.interval_minutes * 60
        self._refresh_uploads()
        if not result["ok"]:
            self._set_status("error", "Check failed")
        elif result["uploaded"]:
            self._set_status("ready", "Uploaded")
        else:
            self._set_status("ready", "Up to date")

    def install_available_update(self) -> None:
        if not self.available_update or self.busy:
            return
        if not messagebox.askyesno(
            "Install update",
            f"Install AutoYouTubeUploader {self.available_update['version']} and restart?",
            parent=self.root,
        ):
            return
        self.busy = True
        self._set_status("busy", "Downloading update…")
        self.install_button.configure(state="disabled")
        threading.Thread(target=self._install_update_worker, args=(self.available_update,), daemon=True).start()

    def _install_update_worker(self, update: dict[str, str]) -> None:
        try:
            downloaded = download_update(update)
            install_update(downloaded)
            self.events.put(("log", "Verified update installed; restarting…"))
            self.events.put(("exit", None))
        except Exception as exc:
            self.events.put(("log", f"Update failed: {exc}"))
            self.events.put(("install_failed", None))

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
        else:
            self.last_link = None
            self.last_title.configure(text="No uploads yet")
        link_state = "normal" if self.last_link else "disabled"
        self.open_link_button.configure(state=link_state)
        self.copy_link_button.configure(state=link_state)

    def open_last_link(self) -> None:
        if self.last_link:
            webbrowser.open(self.last_link)

    def copy_last_link(self) -> None:
        if self.last_link:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.last_link)
            self._log("Copied the YouTube link to the clipboard.")

    def open_workbook(self) -> None:
        self._open_target(self._config().get("workbook", str(auto_yt.DEFAULT_WORKBOOK)))

    def open_video_folder(self) -> None:
        self._open_target(self._config().get("local_video_dir", str(auto_yt.DEFAULT_LOCAL_VIDEO_DIR)))

    def open_log_file(self) -> None:
        log_path = _log_path()
        log_path.parent.mkdir(parents=True, exist_ok=True)
        log_path.touch(exist_ok=True)
        self._open_target(str(log_path))

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

    def open_changelog(self) -> None:
        dialog = ctk.CTkToplevel(self.root)
        dialog.title(f"Changelog - Auto YouTube Uploader {auto_yt.APP_VERSION}")
        dialog.geometry("640x560")
        dialog.configure(fg_color=CARD)
        dialog.transient(self.root)
        text = ctk.CTkTextbox(dialog, font=_font(13), fg_color=SUBTLE, text_color=TEXT, wrap="word")
        text.pack(fill="both", expand=True, padx=20, pady=20)
        text.insert("1.0", _read_changelog())
        text.configure(state="disabled")
        dialog.after(100, dialog.lift)

    # Activity log and event loop

    def _log(self, text: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        lines = [line for line in text.splitlines() if line.strip()]
        if not lines:
            return
        block = "\n".join(f"[{stamp}] {line}" for line in lines) + "\n"
        self.log_box.configure(state="normal")
        self.log_box.insert("end", block)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        with contextlib.suppress(OSError):
            log_path = _log_path()
            log_path.parent.mkdir(parents=True, exist_ok=True)
            with log_path.open("a", encoding="utf-8") as stream:
                stream.write(block)

    def _clear_log(self) -> None:
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

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
            elif kind == "done":
                self._finish_check(value)  # type: ignore[arg-type]
            elif kind == "install_failed":
                self.busy = False
                self._set_status("error", "Update failed")
                if self.available_update:
                    self.install_button.configure(state="normal")
            elif kind == "exit":
                self.root.destroy()
                return
        self.root.after(150, self._process_events)

    def _close(self) -> None:
        if self.busy:
            messagebox.showwarning("Please wait", "A workbook check or update is still running.", parent=self.root)
            return
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
        AutoYouTubeApp(root, config)
        marker = os.environ.get("AUTOYT_GUI_SMOKE_MARKER")
        if marker:
            def finish_smoke_test() -> None:
                root.update()
                Path(marker).write_text(
                    f"title={root.title()}\nmapped={root.winfo_ismapped()}\n",
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
