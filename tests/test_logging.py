import unittest

from core.logging import _format_log_html


class LogHtmlFormattingTests(unittest.TestCase):
    def test_format_log_html_escapes_user_content(self):
        html = _format_log_html("<script>alert('x')</script> & text")

        self.assertIn("&lt;script&gt;", html)
        self.assertIn("&amp; text", html)
        self.assertNotIn("<script>", html)

    def test_format_log_html_keeps_filename_highlighting(self):
        html = _format_log_html("Erstellt 'template.docx'")

        self.assertIn("<strong>'template.docx'</strong>", html)

    def test_format_log_html_converts_newlines_to_breaks(self):
        html = _format_log_html("Zeile 1\nZeile 2")

        self.assertEqual(html, "Zeile 1<br>Zeile 2")


if __name__ == "__main__":
    unittest.main()
