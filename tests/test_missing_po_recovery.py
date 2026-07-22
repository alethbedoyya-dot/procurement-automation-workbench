import os
import tempfile
import unittest
from unittest.mock import patch

import web_query


class MissingPoRecoveryTests(unittest.TestCase):
    def _make_po_folder(self, root, category, bucket, po):
        path = os.path.join(root, category, bucket, str(po))
        os.makedirs(path)

    def test_audits_missing_po_using_existing_folders_from_an_older_run(self):
        with tempfile.TemporaryDirectory() as temp_dir, patch.object(
            web_query, "DOWNLOAD_DIR", temp_dir
        ):
            self._make_po_folder(temp_dir, "空调", "FLD文件", "4000000001")
            self._make_po_folder(temp_dir, "空调", "无FLD文件", "4000000002")

            audit = web_query.audit_download_completeness(
                ["4000000001", "4000000002", "4000000003"], "空调"
            )

        self.assertEqual(audit["expected_pos"], ["4000000001", "4000000002", "4000000003"])
        self.assertEqual(audit["folder_pos"], ["4000000001", "4000000002"])
        self.assertEqual(audit["missing_pos"], ["4000000003"])

    def test_retry_only_submits_pos_that_are_missing_from_existing_folders(self):
        with tempfile.TemporaryDirectory() as temp_dir, patch.object(
            web_query, "DOWNLOAD_DIR", temp_dir
        ), patch.object(
            web_query, "extract_all_pos_from_pivot",
            return_value=["4000000001", "4000000002", "4000000003"],
        ), patch.object(
            web_query, "query_and_download_attachments",
            return_value=(True, "补查完成"),
        ) as query:
            self._make_po_folder(temp_dir, "空调", "FLD文件", "4000000001")
            self._make_po_folder(temp_dir, "空调", "无FLD文件", "4000000002")
            web_query._save_pending_backfill_manifest({
                "category_label": "空调",
                "target_sheet": "空调透视表",
                # 旧版本会把详情页未打开的 PO 误记进 results；补查必须以目录为准。
                "results": [
                    {"po": "4000000001", "fld_files": ["FLD.msg"]},
                    {"po": "4000000002", "fld_files": []},
                    {"po": "4000000003", "fld_files": []},
                ],
            })

            success, message = web_query.retry_missing_pos(
                target_sheet="空调透视表", category_label="空调"
            )

        self.assertTrue(success)
        self.assertIn("4000000003", message)
        self.assertEqual(query.call_count, 1)
        self.assertEqual(query.call_args.kwargs["po_numbers"], ["4000000003"])
        self.assertTrue(query.call_args.kwargs["is_recovery_run"])

    def test_retry_can_recover_an_old_run_after_its_manifest_was_archived(self):
        """用户完成过一次回填后，仍可依据保留的 PO 目录补查少数遗漏项。"""
        with tempfile.TemporaryDirectory() as temp_dir, patch.object(
            web_query, "DOWNLOAD_DIR", temp_dir
        ), patch.object(
            web_query, "extract_all_pos_from_pivot",
            return_value=["4000000001", "4000000002", "4000000003"],
        ), patch.object(
            web_query, "query_and_download_attachments",
            return_value=(True, "补查完成"),
        ) as query:
            self._make_po_folder(temp_dir, "空调", "FLD文件", "4000000001")
            self._make_po_folder(temp_dir, "空调", "无FLD文件", "4000000002")

            success, message = web_query.retry_missing_pos(
                target_sheet="空调透视表", category_label="空调"
            )

        self.assertTrue(success)
        self.assertIn("4000000003", message)
        self.assertEqual(query.call_args.kwargs["po_numbers"], ["4000000003"])
        self.assertIsNone(query.call_args.kwargs["resume_manifest"])

    def test_retry_does_not_start_a_full_run_when_no_previous_query_evidence_exists(self):
        with tempfile.TemporaryDirectory() as temp_dir, patch.object(
            web_query, "DOWNLOAD_DIR", temp_dir
        ), patch.object(
            web_query, "extract_all_pos_from_pivot",
            return_value=["4000000001", "4000000002"],
        ), patch.object(web_query, "query_and_download_attachments") as query:
            success, message = web_query.retry_missing_pos(
                target_sheet="空调透视表", category_label="空调"
            )

        self.assertFalse(success)
        self.assertIn("先点击「③ 打开网站查询」", message)
        query.assert_not_called()

    def test_retry_requeries_a_completed_po_when_its_wbs_is_blank(self):
        """目录齐全但 WBS 漏写时，补查应只重新查询该 PO。"""
        with tempfile.TemporaryDirectory() as temp_dir, patch.object(
            web_query, "DOWNLOAD_DIR", temp_dir
        ), patch.object(
            web_query, "extract_all_pos_from_pivot",
            return_value=["4000000001", "4000000002"],
        ), patch.object(
            web_query, "_find_pos_with_missing_wbs", create=True,
            return_value=["4000000002"],
        ), patch.object(
            web_query, "query_and_download_attachments",
            return_value=(True, "补全完成"),
        ) as query:
            self._make_po_folder(temp_dir, "空调", "FLD文件", "4000000001")
            self._make_po_folder(temp_dir, "空调", "无FLD文件", "4000000002")

            success, message = web_query.retry_missing_pos(
                target_sheet="空调透视表", category_label="空调"
            )

        self.assertTrue(success)
        self.assertIn("WBS", message)
        self.assertEqual(query.call_args.kwargs["po_numbers"], ["4000000002"])

    def test_merges_recovery_result_by_replacing_the_old_result_for_the_same_po(self):
        merged = web_query.merge_query_results(
            [
                {"po": "4000000001", "fld_files": ["old.msg"]},
                {"po": "4000000003", "fld_files": []},
            ],
            [{"po": "4000000003", "fld_files": ["FLD.msg"]}],
            ["4000000001", "4000000002", "4000000003"],
        )

        self.assertEqual(
            merged,
            [
                {"po": "4000000001", "fld_files": ["old.msg"]},
                {"po": "4000000003", "fld_files": ["FLD.msg"]},
            ],
        )

    def test_drops_stale_result_when_a_retried_po_fails_again(self):
        merged = web_query.merge_query_results(
            [
                {"po": "4000000001", "fld_files": ["old.msg"]},
                {"po": "4000000003", "fld_files": []},
            ],
            [],
            ["4000000001", "4000000002", "4000000003"],
            replaced_pos=["4000000003"],
        )

        self.assertEqual(
            merged,
            [{"po": "4000000001", "fld_files": ["old.msg"]}],
        )


if __name__ == "__main__":
    unittest.main()
