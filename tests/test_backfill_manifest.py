import os
import tempfile
import unittest
from unittest.mock import patch

import openpyxl

import web_query


class PendingBackfillManifestTests(unittest.TestCase):
    def test_extracts_wbs_from_the_detail_page_header(self):
        class DetailPage:
            def inner_text(self, selector):
                assert selector == "body"
                return "VIEW编号 343130\n项目WBS No  FSSO-71380\n项目名称 测试项目"

        messages = []
        wbs = web_query._extract_wbs(DetailPage(), messages.append)

        self.assertEqual(wbs, "FSSO-71380")
        self.assertTrue(any("WBS" in message for message in messages))

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

    def test_backfill_writes_wbs_for_results_with_and_without_fld(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = os.path.join(temp_dir, "采购记录.xlsx")
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.title = "装潢透视表"
            worksheet.append(["采购凭证", "WBS"])
            worksheet.append(["4000000001", None])
            worksheet.append(["4000000001", None])
            worksheet.append(["4000000002", None])
            workbook.save(excel_path)
            workbook.close()

            with patch.object(web_query, "DOWNLOAD_DIR", temp_dir), patch.object(
                web_query, "BACKFILL_RUN_LOG_FILE", os.path.join(temp_dir, "backfill.log")
            ), patch.object(web_query, "EXCEL_FILE", excel_path):
                web_query._save_pending_backfill_manifest({
                    "category_label": "装潢",
                    "target_sheet": "装潢透视表",
                    "results": [
                        {
                            "po": "4000000001",
                            "wbs": "FSSO-71380",
                            "fld_files": ["FLD审批.msg"],
                        },
                        {
                            "po": "4000000002",
                            "wbs": "FSSO-71381",
                            "fld_files": [],
                        },
                    ],
                })

                success, _ = web_query.backfill_downloaded_results(
                    target_sheet="装潢透视表", category_label="装潢"
                )

            self.assertTrue(success)
            workbook = openpyxl.load_workbook(excel_path, data_only=True)
            worksheet = workbook["装潢透视表"]
            self.assertEqual(worksheet["B2"].value, "FSSO-71380")
            self.assertEqual(worksheet["B3"].value, "FSSO-71380")
            self.assertEqual(worksheet["B4"].value, "FSSO-71381")
            workbook.close()

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
