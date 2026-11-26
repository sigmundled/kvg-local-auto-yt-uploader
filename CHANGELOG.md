# Changelog

All notable changes to Auto YouTube Uploader are listed here, newest first.

## 1.4.0

- Downloads: column M links to Cloudinary files and to `download.kvernelandgroup.com` pages
  are supported. Kverneland pages are resolved to their embedded Cloudinary MP4, so files over
  1 GB can be used. Other download failures are reported directly and no longer fall back to
  SharePoint.
- Thumbnails: an image named in the thumbnail column (column N, or the "Thumbnail" header) is
  read from the `Thumbnail Image` folder beside `Video File` and set on the uploaded video.
  Images must be JPG or PNG and at most 2 MB. A thumbnail failure is reported as a warning and
  does not fail the upload.
- Workbook sync: for a remote (SharePoint) workbook, the client does not write the link back.
  It shows the row and link to paste manually, and it never uploads the same row twice.
- Client: the window shows the current step (Check, Download, Upload, Thumbnail, Link) and a
  progress bar, and reports which step failed.
- The update manifest URL is no longer shown in settings and cannot be changed.

## 1.3.0

- Client is now packaged as a folder (`AutoYouTubeUploader\` with `AutoYouTubeUploader.exe`
  and `_internal`) instead of a single file, so it starts faster.
- Updates are delivered as a ZIP of that folder. The update replaces program files and keeps
  `config.json` and per-user data.
- Default update source is the GitHub Releases `latest.json`, so no URL needs to be typed.
- Build can optionally code-sign the EXE with `AUTOYT_SIGN_PFX` (see README).

## 1.2.0

- Settings are stored in `config.json` beside the app, so setup is not repeated on each launch.
- New modern client window (customtkinter) replaces the old Tk window. It shows status, a
  "Check now" button, pause, a countdown to the next check, progress, and the upload link.
- Launching with no arguments opens the client. Passing arguments runs the command-line tool.
- First-run Welcome dialog is shown on the visible window, so it no longer opens hidden.
- Added this Changelog view.
- Build bundles `customtkinter` and this changelog.

## 1.1.0

- Hourly automatic check for the latest request row with a blank YouTube Link.
- Update checks against an HTTPS manifest with SHA-256 verification.
- Version is shown in the window title and status.
- Packaged as a Windows EXE with first-run setup, so colleagues do not need Python.
