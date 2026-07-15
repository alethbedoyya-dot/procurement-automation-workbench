import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import web_query


class FakeDownload:
    def __init__(self, filename):
        self.suggested_filename = filename

    def save_as(self, path):
        Path(path).write_bytes(b"test download")


class FakeDownloadExpectation:
    def __init__(self, page):
        self.page = page

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        return False

    @property
    def value(self):
        return FakeDownload(self.page.current_filename)


class FakeLink:
    def click(self):
        pass


class FakeLocator:
    @property
    def first(self):
        return self

    def scroll_into_view_if_needed(self):
        pass


class FakePage:
    url = "https://example.test/detail"

    def wait_for_selector(self, *args, **kwargs):
        pass

    def locator(self, selector):
        return FakeLocator()

    def evaluate(self, expression):
        pass

    def expect_download(self, timeout):
        return FakeDownloadExpectation(self)


class DownloadFolderClassificationTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def _download(self, attachment_names):
        page = FakePage()
        links = []
        for name in attachment_names:
            link = FakeLink()
            # expect_download() only needs the filename of the current link.
            original_click = link.click

            def click(file_name=name, click=original_click):
                page.current_filename = file_name
                click()

            link.click = click
            links.append((link, name))

        fld_dir = self.root / "FLD文件"
        no_fld_dir = self.root / "无FLD文件"
        with patch.object(web_query, "FLD_DIR", str(fld_dir)), \
             patch.object(web_query, "NO_FLD_DIR", str(no_fld_dir)), \
             patch.object(web_query, "_collect_all_attachment_links", return_value=links), \
             patch("web_query.time.sleep"):
            result = web_query._download_support_files(page, lambda _msg: None, "PO-001")
        return result, fld_dir / "PO-001", no_fld_dir / "PO-001"

    def test_po_with_fld_creates_only_fld_folder_and_keeps_all_attachments_together(self):
        result, fld_target, no_fld_target = self._download(
            ["FLD价格审批.msg", "普通报价单.pdf"]
        )

        self.assertEqual(result["fld_files"], ["FLD价格审批.msg"])
        self.assertEqual(result["non_fld_files"], ["普通报价单.pdf"])
        self.assertTrue(fld_target.is_dir())
        self.assertFalse(no_fld_target.exists())

    def test_po_without_fld_creates_only_no_fld_folder(self):
        result, fld_target, no_fld_target = self._download(
            ["普通报价单.pdf"]
        )

        self.assertEqual(result["fld_files"], [])
        self.assertEqual(result["non_fld_files"], ["普通报价单.pdf"])
        self.assertFalse(fld_target.exists())
        self.assertTrue(no_fld_target.is_dir())

    def test_po_without_attachments_creates_no_fld_folder(self):
        result, fld_target, no_fld_target = self._download([])

        self.assertEqual(result["fld_files"], [])
        self.assertEqual(result["non_fld_files"], [])
        self.assertFalse(fld_target.exists())
        self.assertTrue(no_fld_target.is_dir())


if __name__ == "__main__":
    unittest.main()
