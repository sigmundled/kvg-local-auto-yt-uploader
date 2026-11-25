import json
import unittest
from unittest.mock import Mock, patch

from updater import check_update, newer_version


class UpdateTests(unittest.TestCase):
    def test_compares_numeric_dotted_versions(self):
        self.assertTrue(newer_version("1.9.0", "1.10.0"))
        self.assertFalse(newer_version("2.0.0", "1.99.0"))

    def test_requires_https_manifest_and_download(self):
        with self.assertRaisesRegex(ValueError, "HTTPS"):
            check_update("http://example.test/latest.json", "1.0.0")

        response = Mock()
        response.__enter__ = Mock(return_value=response)
        response.__exit__ = Mock(return_value=False)
        response.raise_for_status.return_value = None
        response.json.return_value = {
            "version": "2.0.0",
            "download_url": "http://example.test/client.exe",
            "sha256": "0" * 64,
        }
        with patch("updater.requests.get", return_value=response):
            with self.assertRaisesRegex(ValueError, "download URL must use HTTPS"):
                check_update("https://example.test/latest.json", "1.0.0")

    def test_returns_only_a_newer_valid_update(self):
        response = Mock()
        response.__enter__ = Mock(return_value=response)
        response.__exit__ = Mock(return_value=False)
        response.raise_for_status.return_value = None
        response.json.return_value = {
            "version": "2.1.0",
            "download_url": "https://example.test/client.exe",
            "sha256": "A" * 64,
            "notes": "Hourly client improvements",
        }
        with patch("updater.requests.get", return_value=response):
            update = check_update("https://example.test/latest.json", "2.0.0")
        self.assertIsNotNone(update)
        self.assertEqual(update["version"], "2.1.0")
        self.assertEqual(update["sha256"], "a" * 64)


class DownloadUpdateTests(unittest.TestCase):
    def _run_download(self, archive_bytes, members):
        import hashlib
        import io
        import shutil
        import tempfile
        import zipfile

        from updater import download_update

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            for name in members:
                archive.writestr(name, b"x")
        payload = buffer.getvalue()
        response = Mock()
        response.__enter__ = Mock(return_value=response)
        response.__exit__ = Mock(return_value=False)
        response.raise_for_status.return_value = None
        response.iter_content.return_value = [payload]
        update = {
            "version": "9.9.9",
            "download_url": "https://example.test/AutoYouTubeUploader-9.9.9.zip",
            "sha256": hashlib.sha256(payload).hexdigest(),
        }
        local_appdata = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, local_appdata, ignore_errors=True)
        with patch.dict("os.environ", {"LOCALAPPDATA": local_appdata}), \
                patch("updater.requests.get", return_value=response):
            return download_update(update)

    def test_extracts_update_archive_containing_the_client(self):
        staged = self._run_download(None, ["AutoYouTubeUploader.exe", "_internal/base.dll"])
        self.assertTrue((staged / "AutoYouTubeUploader.exe").is_file())

    def test_rejects_archive_without_the_client_executable(self):
        with self.assertRaisesRegex(ValueError, "does not contain AutoYouTubeUploader.exe"):
            self._run_download(None, ["_internal/base.dll"])


if __name__ == "__main__":
    unittest.main()
