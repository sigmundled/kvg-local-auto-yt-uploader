import unittest

from client_app import _format_features, _youtube_studio_url


class YouTubeStudioUrlTests(unittest.TestCase):
    def test_builds_studio_edit_url_from_short_youtube_url(self):
        self.assertEqual(
            _youtube_studio_url("https://youtu.be/Z_7_0X5oeLs"),
            "https://studio.youtube.com/video/Z_7_0X5oeLs/edit",
        )

    def test_builds_studio_edit_url_from_watch_url(self):
        self.assertEqual(
            _youtube_studio_url("https://www.youtube.com/watch?v=Z_7_0X5oeLs"),
            "https://studio.youtube.com/video/Z_7_0X5oeLs/edit",
        )

    def test_rejects_urls_without_a_video_id(self):
        self.assertIsNone(_youtube_studio_url("https://www.youtube.com/"))


class LatestFeaturesTests(unittest.TestCase):
    def test_formats_version_date_and_unordered_release_notes(self):
        notes = "Released: 2026-11-26\n- First change\n- Second change"
        self.assertEqual(
            _format_features("1.4.2", notes),
            "Version 1.4.2  ·  2026-11-26\n• First change\n• Second change",
        )


if __name__ == "__main__":
    unittest.main()
