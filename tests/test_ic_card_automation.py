import gc
import runpy
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl


PROJECT_DIR = Path(__file__).resolve().parents[1]
GUI_SCRIPT = next(
    path
    for path in PROJECT_DIR.glob("*.py")
    if path.name != "web_query.py" and path.stat().st_size > 50_000
)


class ICCardAutomationTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        ttk_stub = types.ModuleType("ttkbootstrap")
        with patch.dict(sys.modules, {"ttkbootstrap": ttk_stub}):
            self.module = runpy.run_path(str(GUI_SCRIPT))
        self.module_globals = self.module["fill_air_conditioning_old_new_prices"].__globals__
        self.excel_path = Path(self.temp_dir.name) / "采购记录.xlsx"
        self.price_path = Path(self.temp_dir.name) / "价格表.xlsx"
        self.tracking_path = Path(self.temp_dir.name) / "NI PM Saving Tracking.xlsx"
        self.download_dir = Path(self.temp_dir.name) / "downloads"
        self.module_globals["EXCEL_FILE"] = str(self.excel_path)
        self.module_globals["DOWNLOAD_DIR"] = str(self.download_dir)
        self.module_globals["TRACKING_FILE"] = str(self.tracking_path)

    def tearDown(self):
        gc.collect()
        self.temp_dir.cleanup()

    def test_ic_card_reuses_air_conditioning_configuration_and_price_flow(self):
        cfg = self.module["CATEGORIES"]["IC卡"]
        self.assertEqual(cfg["filter_materials"], [1000027312])
        self.assertEqual(cfg["data_sheet"], "IC卡数据")
        self.assertEqual(cfg["target_sheet"], "IC卡透视表")
        self.assertEqual(cfg["content_filter_values"], ("IC卡",))
        self.assertEqual(cfg["price_list_sheet"], "IC卡25.10降价")

        workbook = openpyxl.Workbook()
        data_sheet = workbook.active
        data_sheet.title = "IC卡数据"
        data_sheet.append([
            "采购凭证", "短文本", "订单净值", "老价格", "新价格", "Saving", "CHECK", "净价", "采购订单数量",
        ])
        data_sheet.append(["PO-IC-001", "IC CARD 001", 100, None, None, None, None, 90, 2])
        pivot_sheet = workbook.create_sheet("IC卡透视表")
        pivot_sheet.append(["采购凭证", "求和项:订单净值", "支持文件名(FLD 审批)", "总saving"])
        workbook.save(self.excel_path)
        workbook.close()

        (self.download_dir / "IC卡" / "无FLD文件" / "PO-IC-001").mkdir(parents=True)

        price_book = openpyxl.Workbook()
        price_sheet = price_book.active
        price_sheet.title = "IC卡25.10降价"
        price_sheet.append([None] * 10)
        price_sheet.append([None, None, None, None, "IC CARD 001", "普通物料", None, 120, None, 110])
        price_book.save(self.price_path)
        price_book.close()

        success, _ = self.module["fill_category_old_new_prices"]("IC卡", str(self.price_path))

        self.assertTrue(success)
        workbook = openpyxl.load_workbook(self.excel_path, data_only=True)
        data_sheet = workbook["IC卡数据"]
        self.assertEqual(
            tuple(data_sheet.cell(2, column).value for column in range(4, 8)),
            (120, 110, 60, 20),
        )
        workbook.close()

    def test_monitor_configuration_is_ready_for_the_shared_six_step_workflow(self):
        cfg = self.module["CATEGORIES"]["监控"]

        self.assertEqual(cfg["filter_materials"], [1000027313])
        self.assertEqual(cfg["data_sheet"], "监控数据")
        self.assertEqual(cfg["target_sheet"], "监控透视表")
        self.assertEqual(cfg["content_filter_values"], ("监控",))
        self.assertTrue(cfg["content_filter_exact_on_multiple_matches"])
        self.assertEqual(cfg["price_list_sheet"], "Monitor 25.10降价")

    def test_monitor_pm_tracking_uses_exact_monitor_content(self):
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = "监控透视表"
        worksheet.append([
            "采购凭证", "求和项:订单净值", "E2E项目名", "项目总金额(审批)",
            "支持文件名(FLD 审批)", "Price", "PlanCost", "总saving",
            "订单是否下完=price-订单净值",
        ])
        worksheet.append(["PO-MON-PM", 200, "监控项目测试", None, None, None, None, None, None])
        workbook.save(self.excel_path)
        workbook.close()

        tracking_book = openpyxl.Workbook()
        tracking_sheet = tracking_book.active
        tracking_sheet.title = "Base Data"
        tracking_sheet.append(["无关表头"])
        tracking_sheet.append(["Project Name", "Price", "PlanCost", "Content"])
        tracking_sheet.append(["监控项目测试采购", 500, 100, "监控"])
        tracking_sheet.append(["监控项目测试历史", 900, 200, "监控-历史"])
        tracking_book.save(self.tracking_path)
        tracking_book.close()

        success, _ = self.module["match_pm_tracking_data"](category="监控")

        self.assertTrue(success)
        workbook = openpyxl.load_workbook(self.excel_path, data_only=True)
        worksheet = workbook["监控透视表"]
        self.assertEqual(
            tuple(worksheet.cell(2, column).value for column in range(6, 10)),
            (500, 100, 40, 300),
        )
        workbook.close()

    def test_ic_card_pm_tracking_uses_exact_ic_card_content(self):
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = "IC卡透视表"
        worksheet.append([
            "采购凭证", "求和项:订单净值", "E2E项目名", "项目总金额(审批)",
            "支持文件名(FLD 审批)", "Price", "PlanCost", "总saving",
            "订单是否下完=price-订单净值",
        ])
        worksheet.append(["PO-IC-PM", 200, "IC卡项目测试", None, None, None, None, None, None])
        workbook.save(self.excel_path)
        workbook.close()

        tracking_book = openpyxl.Workbook()
        tracking_sheet = tracking_book.active
        tracking_sheet.title = "Base Data"
        tracking_sheet.append(["无关表头"])
        tracking_sheet.append(["Project Name", "Price", "PlanCost", "Content"])
        tracking_sheet.append(["IC卡项目测试-采购", 500, 100, "IC卡"])
        tracking_sheet.append(["IC卡项目测试-历史", 900, 200, "IC卡-历史"])
        tracking_book.save(self.tracking_path)
        tracking_book.close()

        success, _ = self.module["match_pm_tracking_data"](category="IC卡")

        self.assertTrue(success)
        workbook = openpyxl.load_workbook(self.excel_path, data_only=True)
        worksheet = workbook["IC卡透视表"]
        self.assertEqual(
            tuple(worksheet.cell(2, column).value for column in range(6, 10)),
            (500, 100, 40, 300),
        )
        workbook.close()


if __name__ == "__main__":
    unittest.main()
