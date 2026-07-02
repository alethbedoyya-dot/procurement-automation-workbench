# -*- coding: utf-8 -*-
"""
装潢 Excel 原生透视表自动化工具
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
工作流程：
  1. 读取 工厂清单 7.XLSX（Plant→区域+分公司 映射）
  2. openpyxl 增强 EXPORT1 2 1.xlsx：在"订单净值"右侧插入"区域"和"分公司"列
  3. 按 工厂↔Plant 匹配填值，保留原 Sheet 全部格式
  4. pandas 筛选 物料∈{1000027307, 1000027308}
  5. COM 生成原生透视表 → Sheet "装潢透视表"（位于 Sheet1 右侧）

格式保留：
  → Sheet1 原有列宽、行高、填充色、字体、边框、合并单元格等均完整保留
  → 仅新增两列（区域、分公司，默认格式）和一个透视表 Sheet

生成效果：
  → 带有折叠/展开层级按钮的原生 Excel 数据透视表
  → 行层级：区域 → 分公司 → 供应商名称 → 采购凭证
  → 值：订单净值（求和）
  → 位置：紧邻 "装潢" Sheet 右侧
"""

import os
import sys
import time
import traceback
import threading
import tkinter as tk
from tkinter import messagebox


# ═══════════════════ 配置区 ═══════════════════

# 脚本所在目录（数据文件也放这里，替换同名文件即可更新）
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 主数据文件
EXCEL_FILE = os.path.join(SCRIPT_DIR, "EXPORT1 2 1.xlsx")

# 工厂清单（Plant → 区域 + 分公司 映射表）
FACTORY_FILE = os.path.join(SCRIPT_DIR, "工厂清单 7.XLSX")

# Sheet 名称
SOURCE_SHEET = "Sheet1"                  # 数据源（主数据所在 Sheet）
DATA_SHEET = "装潢数据"                   # 筛选后的中间数据 Sheet
TARGET_SHEET = "装潢透视表"               # 目标透视表 Sheet

# 透视表结构
ROW_FIELDS = ["区域", "分公司", "供应商名称", "采购凭证"]
VALUE_FIELD = "订单净值"
VALUE_NAME = "求和项:订单净值"

# 物料筛选条件
FILTER_MATERIALS = [1000027307, 1000027308]

# 校验所需列（含将要新增的列）
CHECK_COLS = ["工厂", "物料", "供应商名称", "采购凭证", "订单净值"]

# 透视表右侧扩展列（预留，后续手动填入）
EXTRA_COLUMNS = [
    "E2E项目名",
    "E2E订单数量",
    "项目总金额(审批)",
    "订单差异(单独计算)",
    "支持文件名(FLD 审批)",
    "Price",
    "PlanCost",
    "总saving",
    "SAVING差异",
]

# Excel COM 常量
XlDatabase = 1
XlRowField = 1
XlSum = -4157
XlPivotTableVersion15 = 6


# ═══════════════════ 工具函数 ═══════════════════

def _col_letter(idx):
    """1→A, 26→Z, 27→AA, ..."""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _com_start():
    """启动 Excel COM 实例"""
    import pythoncom
    pythoncom.CoInitialize()
    from win32com.client import Dispatch
    excel = Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    return excel


def _com_stop(excel, wb=None):
    """安全关闭 COM 资源"""
    if wb is not None:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
    if excel is not None:
        try:
            excel.ScreenUpdating = True
            excel.DisplayAlerts = True
            excel.Quit()
        except Exception:
            pass
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except Exception:
        pass


def _com_try_open(excel, path):
    """尝试用 COM 打开文件，失败返回 None"""
    try:
        return excel.Workbooks.Open(path)
    except Exception:
        return None


def _com_create_pivot(wb_com, total_cols, total_rows, source_name=DATA_SHEET):
    """在 COM 工作簿中创建/更新透视表"""
    # 删除旧透视表 Sheet
    for s in list(wb_com.Sheets):
        if s.Name == TARGET_SHEET:
            s.Delete()
            break

    # 定位数据源区域
    src_ws = wb_com.Sheets(source_name)
    data_range = src_ws.Range(f"A1:{_col_letter(total_cols)}{total_rows + 1}")

    # 在 Sheet1 右侧新建透视表 Sheet
    sheet1 = wb_com.Sheets(SOURCE_SHEET)
    new_ws = wb_com.Sheets.Add(After=sheet1)
    new_ws.Name = TARGET_SHEET

    # 创建数据透视表缓存
    pivot_cache = wb_com.PivotCaches().Create(
        SourceType=XlDatabase, SourceData=data_range, Version=XlPivotTableVersion15
    )

    # 在工作表上创建透视表
    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=new_ws.Cells(1, 1), TableName="装潢透视表"
    )

    # 添加行字段（层级顺序）
    for i, field_name in enumerate(ROW_FIELDS, start=1):
        pf = pivot_table.PivotFields(field_name)
        pf.Orientation = XlRowField
        pf.Position = i

    # 添加值字段（求和）
    pivot_table.AddDataField(
        pivot_table.PivotFields(VALUE_FIELD), VALUE_NAME, XlSum
    )


# ═══════════════════ 数据增强 ═══════════════════

def _build_plant_mapping(_log):
    """
    读取 工厂清单 7.XLSX，建立 Plant → (区域, 分公司) 映射字典。
    按列位置读取：A=Plant, J=区域(英文), K=分公司(拼音)
    """
    import pandas as pd
    _log("正在读取工厂清单...")
    df_factory = pd.read_excel(FACTORY_FILE, sheet_name=0, header=None)

    # ── 诊断：打印前 3 行原始数据 ──
    debug_lines = ["[诊断] 工厂清单前 3 行 A/J/K 列原始值:"]
    for i in range(min(3, len(df_factory))):
        a_raw = repr(df_factory.iloc[i, 0])
        j_raw = repr(df_factory.iloc[i, 9])
        k_raw = repr(df_factory.iloc[i, 10])
        debug_lines.append(f"  行{i+1}: A={a_raw}  J={j_raw}  K={k_raw}")

    plant_map = {}
    skipped = 0
    for _, row in df_factory.iterrows():
        plant = str(row[0]).strip().upper().replace('\u3000', '').replace('\xa0', '')   # A 列: Plant
        region = str(row[9]).strip()  # J 列: 区域 (South/East 等)
        branch = str(row[10]).strip() # K 列: 分公司 (Zhongshan 拼音)
        if plant and plant != "NAN" and plant != "PLANT":
            plant_map[plant] = (region if region != "nan" else "", branch if branch != "nan" else "")
        else:
            skipped += 1

    # ── 诊断：打印前 5 个映射 key ──
    sample_keys = list(plant_map.keys())[:5]
    debug_lines.append(f"[诊断] 跳过行数: {skipped}, 有效映射条目: {len(plant_map)}")
    debug_lines.append(f"[诊断] 映射样例 (前5个): {[repr(k) for k in sample_keys]}")

    _log(f"工厂映射条目: {len(plant_map)}")
    return plant_map, "\n".join(debug_lines)


def _enhance_and_filter(plant_map, _log):
    """
    openpyxl 增强 Sheet1：
      1. 找到"订单净值"列 → 在右侧插入两列："区域""分公司"
      2. 按"工厂"列值查 plant_map → 填入区域和分公司
      3. pandas 读增强后的数据 → 筛选 物料∈{1000027307, 1000027308}
      4. 将筛选结果写入新 Sheet "装潢数据"
    返回: (total_cols, total_rows) 筛选后数据的行列数（供 COM 透视表使用）
    """
    import openpyxl, pandas as pd

    # ── 1. 打开原文件，找列位置 ──
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb[SOURCE_SHEET]

    # 扫描表头，定位各列
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is not None:
            headers[str(v).strip()] = c

    if "工厂" not in headers or "订单净值" not in headers:
        wb.close()
        raise RuntimeError("Sheet1 缺少必需列：工厂 或 订单净值")
    factory_col = headers["工厂"]
    order_col = headers["订单净值"]

    # ── 2. 在"订单净值"右侧插入"区域"和"分公司" ──
    has_region = "区域" in headers
    has_branch = "分公司" in headers
    if not has_region and not has_branch:
        ws.insert_cols(order_col + 1, 2)
        region_col = order_col + 1
        branch_col = order_col + 2
        ws.cell(1, region_col, "区域")
        ws.cell(1, branch_col, "分公司")
        _log("已在「订单净值」右侧插入「区域」和「分公司」列")
    elif has_region:
        region_col = headers["区域"]
        branch_col = headers.get("分公司", region_col + 1)
        _log("区域列已存在，将更新值")
    else:
        branch_col = headers["分公司"]
        ws.insert_cols(branch_col, 1)
        region_col = branch_col
        branch_col = branch_col + 1
        ws.cell(1, region_col, "区域")
        _log("补插「区域」列")

    # ── 诊断：打印 EXPORT1 工厂列前 5 个值 ──
    debug_lines = ["[诊断] EXPORT1 工厂列前 5 个值:"]
    for r in range(2, min(7, ws.max_row + 1)):
        raw_val = ws.cell(r, factory_col).value
        debug_lines.append(f"  行{r}: {repr(raw_val)}  → str.strip: {repr(str(raw_val or '').strip())}")

    # ── 3. 填值 ──
    matched, unmatched = 0, 0
    unmatched_samples = []
    for r in range(2, ws.max_row + 1):
        fv = str(ws.cell(r, factory_col).value or "").strip().upper().replace('\u3000', '').replace('\xa0', '')
        if fv in plant_map:
            reg, brh = plant_map[fv]
            ws.cell(r, region_col, reg)
            ws.cell(r, branch_col, brh)
            matched += 1
        else:
            unmatched += 1
            if len(unmatched_samples) < 5:
                unmatched_samples.append(repr(fv))
    debug_lines.append(f"[诊断] 匹配成功: {matched}, 未匹配: {unmatched}")
    if unmatched_samples:
        debug_lines.append(f"[诊断] 未匹配样例: {unmatched_samples}")
    _log(f"区域/分公司匹配: 成功 {matched}, 未匹配 {unmatched}")

    debug_str = "\n".join(debug_lines)

    wb.save(EXCEL_FILE)
    wb.close()

    # ── 4. pandas 读增强后数据 → 筛选 → 写入「装潢数据」Sheet ──
    _log("正在筛选物料...")
    df_all = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET)
    # 确保物料列为数值类型用于比较
    df_all["物料"] = pd.to_numeric(df_all["物料"], errors="coerce")
    df_filtered = df_all[df_all["物料"].isin(FILTER_MATERIALS)].copy()
    df_filtered = df_filtered.reset_index(drop=True)

    if len(df_filtered) == 0:
        raise RuntimeError(f"筛选后无数据！物料 {FILTER_MATERIALS} 在数据中不存在。")
    _log(f"筛选后数据行数: {len(df_filtered)}")

    # 写入「装潢数据」Sheet
    wb2 = openpyxl.load_workbook(EXCEL_FILE)
    if DATA_SHEET in wb2.sheetnames:
        del wb2[DATA_SHEET]
    ds = wb2.create_sheet(title=DATA_SHEET)
    # 写表头
    for ci, col_name in enumerate(df_filtered.columns, 1):
        ds.cell(1, ci, col_name)
    # 写数据
    for ri, row in df_filtered.iterrows():
        for ci, val in enumerate(row, 1):
            ds.cell(ri + 2, ci, val)
    wb2.save(EXCEL_FILE)
    wb2.close()

    return len(df_filtered.columns), len(df_filtered), debug_str


# ═══════════════════ 扩展列 ═══════════════════

def add_extra_columns():
    """
    在「装潢透视表」Sheet 中，透视表右侧追加 EXTRA_COLUMNS 表头。
    重复调用会自动更新列名（旧列名自动替换为新列名，已有数据不丢失）。
    """
    import openpyxl

    if not os.path.exists(EXCEL_FILE):
        return False, f"找不到文件：\n{EXCEL_FILE}"

    wb = openpyxl.load_workbook(EXCEL_FILE)
    if TARGET_SHEET not in wb.sheetnames:
        wb.close()
        return False, f"Sheet \"{TARGET_SHEET}\" 不存在，请先生成透视表。"

    ws = wb[TARGET_SHEET]

    # 找到第 1 行最右侧已用列
    max_col = ws.max_column
    last_used = 0
    for c in range(max_col, 0, -1):
        if ws.cell(1, c).value is not None:
            last_used = c
            break

    if last_used == 0:
        wb.close()
        return False, f"Sheet \"{TARGET_SHEET}\" 第 1 行为空，无法定位。"

    # 定位扩展列起始列：优先找已有「E2E项目名」，否则从 last_used + 1 开始
    start_col = None
    for c in range(1, max_col + 1):
        if str(ws.cell(1, c).value or "").strip() == EXTRA_COLUMNS[0]:
            start_col = c
            break
    if start_col is None:
        start_col = last_used + 1

    # 写入/更新表头（覆盖旧列名，保留已有数据）
    updated = 0
    for i, col_name in enumerate(EXTRA_COLUMNS):
        cell = ws.cell(1, start_col + i)
        existing = str(cell.value or "").strip()
        if existing == col_name:
            continue
        cell.value = col_name
        updated += 1

    wb.save(EXCEL_FILE)
    wb.close()

    col_range = f"{_col_letter(start_col)}~{_col_letter(start_col + len(EXTRA_COLUMNS) - 1)}"
    action = f"更新了 {updated} 个列名" if updated > 0 else "列名已是最新，无需更新"
    return True, f"已在「{TARGET_SHEET}」透视表右侧配置 {len(EXTRA_COLUMNS)} 个扩展列表头。\n列范围：{col_range}\n{action}\n列名：{', '.join(EXTRA_COLUMNS)}"


# ═══════════════════ NI PM Saving Tracking 匹配 ═══════════════════

TRACKING_FILE = os.path.join(SCRIPT_DIR, "NI PM Saving Tracking.xlsx")
TRACKING_SHEET = "Base Data"
TRACKING_PROJECT_COL = "Project Name"


def match_pm_tracking_data():
    """
    读取「装潢透视表」中的 E2E项目名，去 NI PM Saving Tracking.xlsx
    的 Base Data Sheet 的 Project Name 列中做模糊匹配（包含即匹配）。

    匹配成功：从 Tracking 提取 Price/PlanCost，读取透视表订单净值，
    计算总saving（= 订单净值 × PlanCost ÷ Price）和 SAVING差异（= 总saving - 订单净值），
    写入透视表对应列。Price 为 0 或空时填入 "N/A"。
    未匹配：跳过不写。

    返回: (success: bool, message: str)
    """
    import openpyxl

    if not os.path.exists(EXCEL_FILE):
        return False, f"找不到文件：\n{EXCEL_FILE}"
    if not os.path.exists(TRACKING_FILE):
        return False, f"找不到 NI PM Saving Tracking 文件：\n{TRACKING_FILE}"

    # ── 1. 打开透视表，扫描第1行定位各列 ──
    wb = openpyxl.load_workbook(EXCEL_FILE)
    if TARGET_SHEET not in wb.sheetnames:
        wb.close()
        return False, f"Sheet「{TARGET_SHEET}」不存在，请先生成透视表。"

    ws = wb[TARGET_SHEET]

    e2e_col = None
    order_value_col = None   # 订单净值（列B "求和项:订单净值"）
    price_col = None
    plancost_col = None
    total_saving_col = None
    saving_diff_col = None

    for col in range(1, ws.max_column + 1):
        val = str(ws.cell(1, col).value or "").strip()
        if val == "E2E项目名":
            e2e_col = col
        elif val == "求和项:订单净值":
            order_value_col = col
        elif val == "Price":
            price_col = col
        elif val == "PlanCost":
            plancost_col = col
        elif val == "总saving":
            total_saving_col = col
        elif val == "SAVING差异":
            saving_diff_col = col

    if e2e_col is None:
        wb.close()
        return False, f"「{TARGET_SHEET}」中未找到「E2E项目名」列，请先点击「添加扩展列」按钮。"
    if order_value_col is None:
        wb.close()
        return False, f"「{TARGET_SHEET}」中未找到「求和项:订单净值」列，透视表可能未正确生成。"
    if price_col is None or plancost_col is None or total_saving_col is None or saving_diff_col is None:
        wb.close()
        return False, (
            f"「{TARGET_SHEET}」中缺少扩展列（Price/PlanCost/总saving/SAVING差异），"
            f"请先点击「添加扩展列」按钮。"
        )

    # ── 2. 收集所有 E2E项目名（跳过空值）──
    project_names = []  # (row_number, project_name)
    for row in range(2, ws.max_row + 1):
        val = str(ws.cell(row, e2e_col).value or "").strip()
        if val:
            project_names.append((row, val))

    if not project_names:
        wb.close()
        return False, "透视表中未找到任何 E2E项目名，请先运行「打开网站查询」。"

    # ── 3. 打开 NI PM Saving Tracking（只读）──
    wb_track = openpyxl.load_workbook(TRACKING_FILE, read_only=True, data_only=True)
    if TRACKING_SHEET not in wb_track.sheetnames:
        wb_track.close()
        wb.close()
        return False, (
            f"NI PM Saving Tracking 中不存在 Sheet「{TRACKING_SHEET}」。\n"
            f"当前 Sheet 列表：{wb_track.sheetnames}"
        )

    ws_track = wb_track[TRACKING_SHEET]

    # 定位 Tracking 中的 Project Name / Price / PlanCost 列（表头在第2行）
    HEADER_ROW = 2
    project_col = None
    track_price_col = None
    track_plancost_col = None

    for col in range(1, ws_track.max_column + 1):
        val = str(ws_track.cell(HEADER_ROW, col).value or "").strip()
        if val == TRACKING_PROJECT_COL:
            project_col = col
        elif val == "Price":
            track_price_col = col
        elif val == "PlanCost":
            track_plancost_col = col

    # 诊断：找不到列时打印表头所有非空列
    def _tracking_headers():
        h = []
        for c in range(1, ws_track.max_column + 1):
            v = str(ws_track.cell(HEADER_ROW, c).value or "")
            if v:
                h.append(f"第{c}列={v}")
        return f"共{ws_track.max_column}列，非空表头：" + (', '.join(h) if h else '(空)')

    if project_col is None:
        headers_str = _tracking_headers()
        wb_track.close()
        wb.close()
        return False, (
            f"NI PM Saving Tracking 的「{TRACKING_SHEET}」Sheet 中未找到列「{TRACKING_PROJECT_COL}」。\n"
            f"{headers_str}"
        )

    missing_track = []
    if track_price_col is None:
        missing_track.append("Price")
    if track_plancost_col is None:
        missing_track.append("PlanCost")
    if missing_track:
        headers_str = _tracking_headers()
        wb_track.close()
        wb.close()
        return False, (
            f"NI PM Saving Tracking 中缺少列：{', '.join(missing_track)}。\n"
            f"{headers_str}"
        )

    # 收集 Tracking 中所有行数据（iter_rows 批量读取，比逐格快100倍+）
    tracking_rows = []
    max_col = max(project_col, track_price_col, track_plancost_col)
    idx_pn = project_col - 1
    idx_price = track_price_col - 1
    idx_pc = track_plancost_col - 1
    for row_vals in ws_track.iter_rows(
        min_row=HEADER_ROW + 1,
        min_col=1, max_col=max_col,
        values_only=True,
    ):
        pn_val = str(row_vals[idx_pn] or "").strip()
        if pn_val:
            tracking_rows.append((
                pn_val.lower(),  # 预计算小写，加速后续匹配
                row_vals[idx_price],
                row_vals[idx_pc],
            ))

    wb_track.close()

    # ── 4. 逐行匹配 + 写入 ──
    matched = 0
    written = 0
    unmatched = 0
    unmatched_samples = []
    na_count = 0

    for row_num, pn in project_names:
        pn_lower = pn.lower()

        # 找第一个模糊匹配的 tracking 行
        found_track = None
        for t_pn, t_price, t_plancost in tracking_rows:
            if pn_lower in t_pn:  # t_pn 已在收集时预计算为小写
                found_track = (t_pn, t_price, t_plancost)
                break

        if found_track is None:
            unmatched += 1
            if len(unmatched_samples) < 5:
                unmatched_samples.append(pn)
            continue

        matched += 1
        _, t_price, t_plancost = found_track

        # 读取订单净值
        order_value = ws.cell(row_num, order_value_col).value

        # 转换数值
        try:
            price_num = float(t_price) if t_price is not None else 0.0
        except (ValueError, TypeError):
            price_num = 0.0
        try:
            plancost_num = float(t_plancost) if t_plancost is not None else 0.0
        except (ValueError, TypeError):
            plancost_num = 0.0
        try:
            order_num = float(order_value) if order_value is not None else 0.0
        except (ValueError, TypeError):
            order_num = 0.0

        # 写入 Price 和 PlanCost（保持原始值类型）
        ws.cell(row_num, price_col, t_price)
        ws.cell(row_num, plancost_col, t_plancost)

        # 计算总saving 和 SAVING差异
        if price_num == 0:
            ws.cell(row_num, total_saving_col, "N/A")
            ws.cell(row_num, saving_diff_col, "N/A")
            na_count += 1
        else:
            total_saving = order_num * plancost_num / price_num
            saving_diff = total_saving - order_num
            ws.cell(row_num, total_saving_col, round(total_saving, 2))
            ws.cell(row_num, saving_diff_col, round(saving_diff, 2))
            written += 1

    # ── 5. 保存 ──
    wb.save(EXCEL_FILE)
    wb.close()

    # ── 6. 返回摘要 ──
    summary_lines = [
        f"NI PM Saving Tracking 匹配完成！",
        f"",
        f"待匹配项目名总数：{len(project_names)}",
        f"匹配成功：{matched}（已写入 Price/PlanCost/总saving/SAVING差异: {written}）",
        f"未匹配（跳过）：{unmatched}",
    ]
    if na_count > 0:
        summary_lines.append(f"Price 为 0/空（填入 N/A）：{na_count}")
    summary_lines.extend([
        f"",
        f"Tracking 文件总行数：{len(tracking_rows)}",
        f"Sheet：{TRACKING_SHEET}，匹配列：{TRACKING_PROJECT_COL}",
    ])
    if unmatched_samples:
        summary_lines.append(f"")
        summary_lines.append(f"未匹配样例（前 5 个）：")
        for s in unmatched_samples:
            summary_lines.append(f"  - {s}")

    return True, "\n".join(summary_lines)


# ═══════════════════ 主流程 ═══════════════════

def generate_pivot_table(status_callback=None):
    """
    完整工作流：
      1. 读工厂清单 → 建 Plant→(区域,分公司) 映射
      2. openpyxl 增强 Sheet1（插入区域/分公司列，按工厂填值，保留格式）
      3. pandas 筛选 物料=1000027307/1000027308 → 写入「装潢数据」Sheet
      4. COM 生成原生透视表 → 「装潢透视表」（Sheet1 右侧）
    """
    _log = lambda msg: status_callback and status_callback(msg)

    # ── 1. 检查文件 ──
    if not os.path.exists(EXCEL_FILE):
        return False, f"找不到主数据文件：\n{EXCEL_FILE}"
    if not os.path.exists(FACTORY_FILE):
        return False, f"找不到工厂清单文件：\n{FACTORY_FILE}"

    # ── 2. 校验 Sheet1 列名 ──
    _log("正在校验列名...")
    try:
        import pandas as pd
        df_head = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, nrows=0)
        missing = [c for c in CHECK_COLS if c not in df_head.columns]
        if missing:
            return False, (
                f"\"{SOURCE_SHEET}\" 中缺少以下列：\n"
                f"{', '.join(missing)}\n\n"
                f"当前列名：{list(df_head.columns)}"
            )
    except ValueError:
        return False, f"工作簿中不存在名为 \"{SOURCE_SHEET}\" 的 Sheet。"
    except Exception as e:
        return False, f"读取 Excel 失败：\n{e}"

    # ── 3. 读取工厂清单，建映射 ──
    try:
        plant_map, factory_debug = _build_plant_mapping(_log)
    except Exception as e:
        return False, f"读取工厂清单失败：\n{e}"

    # ── 4. 增强 Sheet1 + 筛选 → 写入「装潢数据」──
    try:
        total_cols, total_rows, match_debug = _enhance_and_filter(plant_map, _log)
    except Exception as e:
        return False, f"数据增强/筛选失败：\n{e}"

    # 汇总诊断信息
    all_debug = f"{factory_debug}\n\n{match_debug}"

    # ── 5. COM 生成原生透视表 ──
    _log("正在调用 Excel 引擎生成原生透视表...")
    excel = None
    wb = None
    try:
        excel = _com_start()
        _log("COM 打开增强后的文件...")
        wb = excel.Workbooks.Open(EXCEL_FILE)
        if wb is None:
            return False, "COM 无法打开文件，请关闭其他正在使用该文件的程序后重试。"

        _com_create_pivot(wb, total_cols, total_rows, source_name=DATA_SHEET)

        _log("正在保存...")
        wb.Save()
        wb.Close(SaveChanges=True)
        wb = None
        time.sleep(0.3)
    except Exception as e:
        return False, f"COM 透视表生成失败：\n{e}"
    finally:
        _com_stop(excel, wb)

    return True, (
        f"原生透视表生成成功！\n\n"
        f"筛选物料：{FILTER_MATERIALS}\n"
        f"筛选后数据行数：{total_rows}\n"
        f"行层级：{' → '.join(ROW_FIELDS)}\n"
        f"值字段：{VALUE_NAME}（求和）\n"
        f"目标 Sheet：\"{TARGET_SHEET}\"（紧邻 Sheet1 右侧）\n\n"
        f"Sheet1 原有格式完整保留，仅在「订单净值」右侧新增了「区域」和「分公司」两列。\n\n"
        f"════════════ 诊断信息 ════════════\n{all_debug}"
    )


# ═══════════════════ GUI 界面 ═══════════════════

class PivotTableApp:
    """Tkinter 主窗口"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("装潢透视表生成工具")
        self.root.geometry("520x440")
        self.root.resizable(False, False)
        self.root.configure(bg="#f5f6fa")

        # 居中
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"+{(sw - 520) // 2}+{(sh - 440) // 2}")

        self._build_ui()

    def _build_ui(self):
        bg = "#f5f6fa"

        # ── 标题 ──
        tk.Label(
            self.root, text="装潢透视表生成工具",
            font=("Microsoft YaHei", 14, "bold"),
            fg="#2c3e50", bg=bg,
        ).pack(pady=(25, 8))

        # ── 说明 ──
        desc = (
            f"数据目录：{SCRIPT_DIR}\n"
            f"数据源：EXPORT1 2 1.xlsx → {SOURCE_SHEET}\n"
            f"工厂清单：工厂清单 7.XLSX（Plant→区域+分公司）\n"
            f"筛选：物料 ∈ {FILTER_MATERIALS}\n"
            f"透视：{' → '.join(ROW_FIELDS)} | {VALUE_FIELD}（求和）\n\n"
            f"更新方法：替换本目录下同名文件后重新运行即可"
        )
        tk.Label(
            self.root, text=desc,
            font=("Microsoft YaHei", 9), fg="#7f8c8d", bg=bg,
            justify=tk.LEFT,
        ).pack(pady=(0, 18))

        # ── 核心按钮 ──
        self.btn = tk.Button(
            self.root, text="生成装潢透视表",
            font=("Microsoft YaHei", 13, "bold"),
            bg="#27ae60", fg="white",
            activebackground="#2ecc71", activeforeground="white",
            relief=tk.FLAT, padx=40, pady=12,
            cursor="hand2", borderwidth=0,
            command=self._on_click,
        )
        self.btn.pack(pady=(0, 8))

        # ── 扩展列按钮 ──
        self.btn_extra = tk.Button(
            self.root, text="添加扩展列（透视表右侧）",
            font=("Microsoft YaHei", 10),
            bg="#3498db", fg="white",
            activebackground="#5dade2", activeforeground="white",
            relief=tk.FLAT, padx=25, pady=8,
            cursor="hand2", borderwidth=0,
            command=self._on_extra_click,
        )
        self.btn_extra.pack(pady=(0, 5))

        # ── 网站查询按钮 ──
        self.btn_web = tk.Button(
            self.root, text="打开网站查询",
            font=("Microsoft YaHei", 10),
            bg="#8e44ad", fg="white",
            activebackground="#a569bd", activeforeground="white",
            relief=tk.FLAT, padx=25, pady=8,
            cursor="hand2", borderwidth=0,
            command=self._on_web_click,
        )
        self.btn_web.pack(pady=(0, 5))

        # ── PM Tracking 匹配按钮 ──
        self.btn_tracking = tk.Button(
            self.root, text="匹配PM Tracking数据",
            font=("Microsoft YaHei", 10),
            bg="#16a085", fg="white",
            activebackground="#1abc9c", activeforeground="white",
            relief=tk.FLAT, padx=25, pady=8,
            cursor="hand2", borderwidth=0,
            command=self._on_tracking_click,
        )
        self.btn_tracking.pack(pady=(0, 15))

        # ── 状态 ──
        self.lbl_status = tk.Label(
            self.root, text="就绪 — 点击按钮开始",
            font=("Microsoft YaHei", 9), fg="#95a5a6", bg=bg,
        )
        self.lbl_status.pack()

    def _on_click(self):
        """点击按钮"""
        self.btn.config(state=tk.DISABLED, text="处理中，请稍候...")
        self._update_status("正在处理...", "#e67e22")

        def _run():
            try:
                success, msg = generate_pivot_table(
                    status_callback=lambda m: self.root.after(0, self._update_status, m, "#e67e22")
                )
                self.root.after(0, lambda: self._done(success, msg))
            except Exception as e:
                err_msg = f"未预料的错误：\n{e}\n\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._done(False, err_msg))

        self.root.after(100, _run)

    def _update_status(self, msg, color):
        self.lbl_status.config(text=msg, fg=color)

    def _on_extra_click(self):
        """点击添加扩展列"""
        self.btn_extra.config(state=tk.DISABLED, text="正在添加...")
        self._update_status("正在添加扩展列...", "#3498db")

        def _run():
            try:
                success, msg = add_extra_columns()
                self.root.after(0, lambda: self._extra_done(success, msg))
            except Exception as e:
                err_msg = f"未预料的错误：\n{e}\n\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._extra_done(False, err_msg))

        self.root.after(100, _run)

    def _extra_done(self, success, msg):
        self.btn_extra.config(state=tk.NORMAL, text="添加扩展列（透视表右侧）")
        if success:
            self._update_status("扩展列添加成功！", "#27ae60")
            messagebox.showinfo("操作成功", msg)
        else:
            self._update_status("添加失败，见弹窗", "#e74c3c")
            messagebox.showerror("操作失败", msg)

    def _on_web_click(self):
        """点击网站查询按钮"""
        self.btn_web.config(state=tk.DISABLED, text="正在启动浏览器...")
        self._update_status("正在启动 Edge 浏览器...", "#8e44ad")

        def _run():
            try:
                import web_query
                success, msg = web_query.open_website_and_search(
                    status_callback=lambda m: self.root.after(0, self._update_status, m, "#8e44ad")
                )
                self.root.after(0, lambda: self._web_done(success, msg))
            except ImportError:
                self.root.after(0, lambda: self._web_done(
                    False, "缺少 playwright 模块，请先运行：\npip install playwright\nplaywright install msedge"
                ))
            except Exception as e:
                err_msg = f"未预料的错误：\n{e}\n\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._web_done(False, err_msg))

        threading.Thread(target=_run, daemon=True).start()

    def _web_done(self, success, msg):
        self.btn_web.config(state=tk.NORMAL, text="打开网站查询")
        if success:
            self._update_status("网站查询完成！", "#27ae60")
            messagebox.showinfo("操作完成", msg)
        else:
            self._update_status("网站查询失败，见弹窗", "#e74c3c")
            messagebox.showerror("操作失败", msg)

    def _on_tracking_click(self):
        """点击 PM Tracking 匹配按钮"""
        self.btn_tracking.config(state=tk.DISABLED, text="正在匹配...")
        self._update_status("正在匹配 NI PM Saving Tracking...", "#16a085")

        def _run():
            try:
                success, msg = match_pm_tracking_data()
                self.root.after(0, lambda: self._tracking_done(success, msg))
            except Exception as e:
                err_msg = f"错误：\n{e}\n\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._tracking_done(False, err_msg))

        threading.Thread(target=_run, daemon=True).start()

    def _tracking_done(self, success, msg):
        self.btn_tracking.config(state=tk.NORMAL, text="匹配PM Tracking数据")
        if success:
            self._update_status("PM Tracking 匹配完成！", "#27ae60")
            messagebox.showinfo("匹配完成", msg)
        else:
            self._update_status("匹配失败，见弹窗", "#e74c3c")
            messagebox.showerror("匹配失败", msg)

    def _done(self, success, msg):
        self.btn.config(state=tk.NORMAL, text="生成装潢透视表")
        if success:
            self._update_status("生成成功！", "#27ae60")
            messagebox.showinfo("操作成功", msg)
        else:
            self._update_status("生成失败，见弹窗", "#e74c3c")
            messagebox.showerror("操作失败", msg)


# ═══════════════════ 入口 ═══════════════════

def main():
    root = tk.Tk()
    app = PivotTableApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
