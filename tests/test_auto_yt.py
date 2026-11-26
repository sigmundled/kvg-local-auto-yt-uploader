import contextlib
import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, Mock, patch

import requests

from auto_yt import (
    APP_DIR,
    GraphClient,
    VideoRequest,
    _config_path,
    _make_request,
    _read_config,
    _run_setup,
    _source,
    _write_youtube_link,
    download_video,
    download_then_upload,
    load_requests,
    main,
    resolve_thumbnail,
    set_thumbnail,
    verify_download,
)


class ConfigLocationTests(unittest.TestCase):
    def test_config_is_stored_beside_the_app(self):
        self.assertEqual(_config_path(), APP_DIR / "config.json")


class WorkbookTests(unittest.TestCase):
    def test_first_run_setup_saves_local_paths_and_oauth_client(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = root / "requests.xlsx"
            workbook.write_bytes(b"workbook")
            video_dir = root / "videos"
            video_dir.mkdir()
            client = root / "client.json"
            client.write_text(
                json.dumps({"installed": {"client_id": "test"}}),
                encoding="utf-8",
            )
            config_path = root / "profile" / "config.json"
            answers = iter([str(workbook), str(video_dir), str(client)])

            config = _run_setup(config_path, lambda _prompt: next(answers))

            self.assertEqual(config["workbook"], str(workbook.resolve()))
            self.assertEqual(config["local_video_dir"], str(video_dir.resolve()))
            self.assertEqual(config["client_secrets"], str(client.resolve()))
            self.assertEqual(_read_config(config_path), config)

    def test_relative_excel_hyperlink_falls_back_to_visible_url(self):
        visible_url = "https://tenant.example/sites/Marketing/video.mp4"
        source, name, size = _source(
            visible_url,
            "Apps/Video%20File/video.mp4",
        )
        self.assertEqual(source, visible_url)
        self.assertEqual(name, "")
        self.assertIsNone(size)

    def test_reads_form_upload_json(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as directory:
            workbook_path = Path(directory) / "requests.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Id", "Email", "Brand", "Title", "Description", "Tags", "Playlist", "Video File", "YouTube Link"])
            sheet.append([
                7, "user@example.test", "Brand", "A video", "Description",
                "tag-one, #tag-two", "",
                json.dumps([{"name": "clip.mp4", "link": "https://tenant.example/clip.mp4", "size": 20}]),
                "",
            ])
            workbook.save(workbook_path)
            sheet_name, requests_ = load_requests(workbook_path)

        self.assertEqual(sheet_name, "Sheet")
        self.assertEqual(len(requests_), 1)
        self.assertEqual(requests_[0].row_id, "7")
        self.assertEqual(requests_[0].video_url, "https://tenant.example/clip.mp4")
        self.assertEqual(requests_[0].expected_size, 20)
        self.assertEqual(requests_[0].tags, ["tag-one", "tag-two"])

    def test_accepts_sharepoint_onedrive_header(self):
        headers = ["Id", "Title", "on SharePoint / OneDrive"]
        values = {
            "Id": 2,
            "Title": "Test",
            "on SharePoint / OneDrive": "https://tenant.example/video.mp4",
        }
        request = _make_request(headers, values, {}, 2)
        self.assertIsNotNone(request)
        self.assertEqual(request.video_url, "https://tenant.example/video.mp4")

    def test_resolves_doc_aspx_by_sourcedoc_id(self):
        office_url = (
            "https://tenant.example/:x:/r/sites/Marketing/_layouts/15/Doc.aspx"
            "?sourcedoc=%7B9B57AE70-1607-4148-A612-4BB5937350B3%7D"
            "&file=Request%20to%20upload%20video.xlsx&action=edit"
        )
        item = {
            "id": "drive-item-id",
            "sharepointIds": {"listItemUniqueId": "9b57ae70-1607-4148-a612-4bb5937350b3"},
        }
        graph = GraphClient.__new__(GraphClient)
        graph.site_url = "https://tenant.example/sites/Marketing"
        graph._json = Mock(side_effect=[
            requests.HTTPError("not a shares link"),
            {"id": "site-id"},
            {"value": [{"id": "drive-id", "webUrl": "https://tenant.example/sites/Marketing/Shared Documents"}]},
            {"value": [item]},
        ])

        self.assertEqual(graph._item(office_url), item)
        self.assertIn(
            "/root/search(q='Request%20to%20upload%20video.xlsx')",
            graph._json.call_args_list[-1].args[0],
        )


class ThumbnailTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp())
        self.video_dir = self.root / "Video File"
        self.thumb_dir = self.root / "Thumbnail Image"
        self.video_dir.mkdir()
        self.thumb_dir.mkdir()
        self.request = VideoRequest(
            6, "6", "Plough", "", [], "", "", "https://example.test/v.mp4", "v.mp4", None, "",
            thumbnail_name="thumb.jpg",
        )

    def tearDown(self):
        import shutil

        shutil.rmtree(self.root, ignore_errors=True)

    def test_thumbnail_is_found_in_sibling_folder(self):
        thumb = self.thumb_dir / "thumb.jpg"
        thumb.write_bytes(b"\xff\xd8\xff" + b"x" * 10)
        self.assertEqual(resolve_thumbnail(self.request, self.video_dir), thumb)

    def test_oversized_thumbnail_is_rejected(self):
        (self.thumb_dir / "thumb.jpg").write_bytes(b"0" * (2 * 1024 * 1024 + 1))
        with self.assertRaisesRegex(ValueError, "larger than 2 MB"):
            resolve_thumbnail(self.request, self.video_dir)

    def test_missing_thumbnail_fails_before_upload(self):
        with self.assertRaises(FileNotFoundError):
            resolve_thumbnail(self.request, self.video_dir)

    def test_set_thumbnail_calls_youtube_api_with_video_id(self):
        path = self.thumb_dir / "thumb.png"
        path.write_bytes(b"\x89PNG" + b"x" * 10)
        youtube = Mock()
        set_thumbnail(youtube, "vid123", path)
        self.assertEqual(youtube.thumbnails.return_value.set.call_args.kwargs["videoId"], "vid123")

    def test_column_n_is_used_when_no_thumbnail_header(self):
        headers = ["Id", "Title", "Video"] + [None] * 10 + ["Extra"]
        values = {"Id": 1, "Title": "T", "Video": "https://example.test/v.mp4", "Extra": "thumb.jpg"}
        request = _make_request(headers, values, {}, 2)
        self.assertEqual(request.thumbnail_name, "thumb.jpg")


class DownloadGateTests(unittest.TestCase):
    def setUp(self):
        self.request = VideoRequest(
            2, "2", "Test video", "", [], "", "", "https://example.test/video.mp4",
            "video.mp4", None, "",
        )

    def test_uploader_is_not_called_when_download_fails(self):
        uploaded = []

        def fail(_request):
            raise OSError("network failure")

        with self.assertRaises(OSError):
            download_then_upload(
                self.request, fail, lambda *_: uploaded.append(True), lambda _: "y"
            )
        self.assertEqual(uploaded, [])

    def test_upload_requires_verified_file_and_explicit_confirmation(self):
        uploaded = []
        with tempfile.TemporaryDirectory() as directory:
            video = Path(directory) / "clip.mp4"
            video.write_bytes(b"\x00\x00\x00\x18ftypisom" + b"x" * 20)
            size, digest = verify_download(video)
            self.assertEqual(size, video.stat().st_size)
            self.assertEqual(len(digest), 64)
            cancelled = download_then_upload(
                self.request,
                lambda _: video,
                lambda _request, path: uploaded.append(path) or "video-id",
                lambda _: "n",
            )
            self.assertIsNone(cancelled)
            self.assertEqual(uploaded, [])
            result = download_then_upload(
                self.request,
                lambda _: video,
                lambda _request, path: uploaded.append(path) or "video-id",
                lambda _: "y",
            )
        self.assertEqual(result, "video-id")
        self.assertEqual(uploaded, [video])

    def test_html_signin_page_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            html = Path(directory) / "not-a-video.mp4"
            html.write_text("<html>sign in</html>", encoding="utf-8")
            with self.assertRaises(ValueError):
                verify_download(html)

    def test_local_video_folder_matches_sharepoint_filename_without_graph(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source_dir = root / "OneDrive"
            nested_dir = source_dir / "subfolder"
            nested_dir.mkdir(parents=True)
            source = nested_dir / "clip.mp4"
            source.write_bytes(b"\x00\x00\x00\x18ftypisom" + b"x" * 24)

            workbook_path = root / "requests.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Id", "Title", "Video File", "YouTube Link"])
            sheet.append([
                1,
                "Local test",
                json.dumps([{"name": "clip.mp4", "link": "https://tenant.example/clip.mp4"}]),
                "",
            ])
            workbook.save(workbook_path)
            output = io.StringIO()
            with patch("auto_yt.requests.get", side_effect=AssertionError("network must not be used")):
                with contextlib.redirect_stdout(output):
                    result = main([
                        "--workbook", str(workbook_path),
                        "--row", "2",
                        "--local-video-dir", str(source_dir),
                        "--download-dir", str(root / "downloads"),
                        "--state-file", str(root / "state.json"),
                        "--dry-run",
                    ])

            self.assertEqual(result, 0)
            self.assertEqual(
                (root / "downloads" / "row-2-clip.mp4").read_bytes(),
                source.read_bytes(),
            )
            self.assertIn("Verified download:", output.getvalue())

    def _http_response(self, content_type, body=b"", text=""):
        response = MagicMock()
        response.headers = {"Content-Type": content_type}
        response.text = text
        response.iter_content.return_value = [body]
        response.__enter__.return_value = response
        response.__exit__.return_value = False
        return response

    def test_kverneland_page_resolves_to_embedded_cloudinary_video(self):
        mp4 = b"\x00\x00\x00\x18ftypisom" + b"x" * 24
        page = self._http_response(
            "text/html; charset=utf-8",
            text='<video src="https://res.cloudinary.com/kvg/video/upload/a.mp4?x=1&amp;y=2"></video>',
        )
        video = self._http_response("video/mp4", body=mp4)
        graph = Mock()
        with tempfile.TemporaryDirectory() as directory:
            with patch("auto_yt.requests.get", side_effect=[page, video]) as get:
                request = VideoRequest(
                    3, "3", "Plough", "", [], "", "", "https://download.kvernelandgroup.com/x", "", None, "",
                )
                result = download_video(request, directory, graph)
                self.assertEqual(result.read_bytes(), mp4)
        self.assertEqual(get.call_args_list[1].args[0], "https://res.cloudinary.com/kvg/video/upload/a.mp4?x=1&y=2")
        graph.download_url.assert_not_called()

    def test_landing_page_without_video_fails_clearly_without_graph(self):
        page = self._http_response("text/html", text="<html>no video here</html>")
        graph = Mock()
        with tempfile.TemporaryDirectory() as directory:
            with patch("auto_yt.requests.get", return_value=page):
                request = VideoRequest(
                    4, "4", "Plough", "", [], "", "", "https://download.kvernelandgroup.com/x", "", None, "",
                )
                with self.assertRaisesRegex(RuntimeError, "no embedded video"):
                    download_video(request, directory, graph)
        graph.download_url.assert_not_called()

    def test_non_sharepoint_http_failure_does_not_fall_back_to_graph(self):
        graph = Mock()
        with tempfile.TemporaryDirectory() as directory:
            with patch("auto_yt.requests.get", side_effect=requests.ConnectionError("down")):
                request = VideoRequest(
                    5, "5", "Plough", "", [], "", "", "https://example.test/video.mp4", "", None, "",
                )
                with self.assertRaisesRegex(RuntimeError, "Could not download"):
                    download_video(request, directory, graph)
        graph.download_url.assert_not_called()

    def test_no_row_argument_uses_latest_request_and_does_not_upload_older_rows(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source_dir = root / "OneDrive"
            source_dir.mkdir()
            old_video = source_dir / "old.mp4"
            latest_video = source_dir / "latest.mp4"
            content = b"\x00\x00\x00\x18ftypisom" + b"x" * 24
            old_video.write_bytes(content)
            latest_video.write_bytes(content)

            workbook_path = root / "requests.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Id", "Title", "Video File", "YouTube Link"])
            sheet.append([
                1, "Old request",
                json.dumps([{"name": old_video.name, "link": "https://tenant.test/old.mp4"}]), "",
            ])
            sheet.append([
                2, "Latest request",
                json.dumps([{"name": latest_video.name, "link": "https://tenant.test/latest.mp4"}]), "",
            ])
            workbook.save(workbook_path)

            with patch("auto_yt.DEFAULT_WORKBOOK", workbook_path), patch(
                "auto_yt.DEFAULT_LOCAL_VIDEO_DIR", source_dir
            ):
                result = main([
                    "--download-dir", str(root / "downloads"),
                    "--state-file", str(root / "state.json"),
                    "--dry-run",
                ])

            self.assertEqual(result, 0)
            self.assertFalse((root / "downloads" / "row-2-old.mp4").exists())
            self.assertTrue((root / "downloads" / "row-3-latest.mp4").exists())

    def test_writes_youtube_link_to_request_row(self):
        from openpyxl import Workbook, load_workbook

        with tempfile.TemporaryDirectory() as directory:
            workbook_path = Path(directory) / "requests.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Id", "Title", "Video File", "YouTube Link"])
            sheet.append([1, "Local test", "video.mp4", ""])
            workbook.save(workbook_path)

            link = "https://www.youtube.com/watch?v=test123"
            _write_youtube_link(workbook_path, "Sheet", 2, link)

            saved = load_workbook(workbook_path, data_only=True)
            self.assertEqual(saved["Sheet"]["D2"].value, link)
            saved.close()
            with self.assertRaisesRegex(RuntimeError, "already has a YouTube link"):
                _write_youtube_link(workbook_path, "Sheet", 2, link)

    def test_local_video_folder_rejects_ambiguous_filename(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source_dir = root / "OneDrive"
            (source_dir / "one").mkdir(parents=True)
            (source_dir / "two").mkdir()
            for folder in ("one", "two"):
                (source_dir / folder / "video.mp4").write_bytes(
                    b"\x00\x00\x00\x18ftypisom" + b"x" * 8
                )

            request = VideoRequest(
                2, "2", "Test video", "", [], "", "",
                "https://example.test/video.mp4", "video.mp4", None, "",
            )
            with self.assertRaisesRegex(ValueError, "Multiple local videos"):
                download_video(request, root / "downloads", None, source_dir)

    def test_cli_dry_run_downloads_and_verifies_without_upload(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.mp4"
            source.write_bytes(b"\x00\x00\x00\x18ftypisom" + b"x" * 24)
            workbook_path = root / "requests.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Id", "Title", "Video File", "YouTube Link"])
            sheet.append([1, "Local test", str(source), ""])
            workbook.save(workbook_path)
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                result = main([
                    "--workbook", str(workbook_path),
                    "--row", "2",
                    "--download-dir", str(root / "downloads"),
                    "--state-file", str(root / "state.json"),
                    "--dry-run",
                ])
            self.assertEqual(result, 0)
            self.assertTrue((root / "downloads" / "row-2-source.mp4").is_file())
            self.assertIn("Verified download:", output.getvalue())
            self.assertIn("no YouTube sign-in or upload was attempted", output.getvalue())


if __name__ == "__main__":
    unittest.main()
