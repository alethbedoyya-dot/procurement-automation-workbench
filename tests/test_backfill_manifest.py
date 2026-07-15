import os
import tempfile
import unittest
from unittest.mock import patch

import web_query


class PendingBackfillManifestTests(unittest.TestCase):
    def test_saves_and_loads_a_manifest_for_its_own_category(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.object(web_query, "DOWNLOAD_DIR", temp_dir), patch.object(
                web_query, "BACKFILL_RUN_LOG_FILE", os.path.join(temp_dir, "backfill.log")
            ):
                manifest = {
                    "category_label": "装潢",
                    "target_sheet": "装潢透视表",
                    "results": [{"po": "4000000001", "fld_files": ["FLD.msg"]}],
                }

                path = web_query._save_pending_backfill_manifest(manifest)
                loaded, error = web_query._load_pending_backfill_manifest(
                    "装潢", "装潢透视表"
                )

                self.assertTrue(os.path.exists(path))
                self.assertIsNone(error)
                self.assertEqual(loaded["results"], manifest["results"])
                self.assertEqual(
                    os.path.basename(path), web_query.PENDING_BACKFILL_FILENAME
                )

    def test_rejects_a_manifest_for_a_different_category_or_sheet(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.object(web_query, "DOWNLOAD_DIR", temp_dir), patch.object(
                web_query, "BACKFILL_RUN_LOG_FILE", os.path.join(temp_dir, "backfill.log")
            ):
                web_query._save_pending_backfill_manifest({
                    "category_label": "空调",
                    "target_sheet": "空调透视表",
                    "results": [],
                })

                manifest, error = web_query._load_pending_backfill_manifest(
                    "装潢", "装潢透视表"
                )

                self.assertIsNone(manifest)
                self.assertIn("空调", error)
                self.assertIn("装潢", error)

    def test_backfill_finalizes_fld_results_and_archives_the_manifest(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.object(web_query, "DOWNLOAD_DIR", temp_dir), patch.object(
                web_query, "BACKFILL_RUN_LOG_FILE", os.path.join(temp_dir, "backfill.log")
            ):
                web_query._save_pending_backfill_manifest({
                    "category_label": "装潢",
                    "target_sheet": "装潢透视表",
                    "results": [
                        {"po": "4000000001", "fld_files": ["FLD.msg"]},
                        {"po": "4000000002", "fld_files": []},
                    ],
                })

                with patch.object(web_query, "_finalize_po_result") as finalize:
                    success, message = web_query.backfill_downloaded_results(
                        target_sheet="装潢透视表", category_label="装潢"
                    )

                self.assertTrue(success)
                self.assertIn("回填完成", message)
                self.assertEqual(finalize.call_count, 1)
                self.assertFalse(
                    os.path.exists(web_query._pending_backfill_path("装潢"))
                )
                archived = [
                    name for name in os.listdir(os.path.join(temp_dir, "装潢"))
                    if name.startswith(web_query.COMPLETED_BACKFILL_PREFIX)
                ]
                self.assertEqual(len(archived), 1)

    def test_backfill_keeps_manifest_when_a_result_cannot_be_finalized(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.object(web_query, "DOWNLOAD_DIR", temp_dir), patch.object(
                web_query, "BACKFILL_RUN_LOG_FILE", os.path.join(temp_dir, "backfill.log")
            ):
                web_query._save_pending_backfill_manifest({
                    "category_label": "空调",
                    "target_sheet": "空调透视表",
                    "results": [{"po": "4000000003", "fld_files": ["FLD.msg"]}],
                })

                with patch.object(
                    web_query, "_finalize_po_result", side_effect=RuntimeError("解析失败")
                ):
                    success, message = web_query.backfill_downloaded_results(
                        target_sheet="空调透视表", category_label="空调"
                    )

                self.assertFalse(success)
                self.assertIn("待回填任务已保留", message)
                self.assertTrue(
                    os.path.exists(web_query._pending_backfill_path("空调"))
                )


if __name__ == "__main__":
    unittest.main()
