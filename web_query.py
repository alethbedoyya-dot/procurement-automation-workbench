# -*- coding: utf-8 -*-
"""
TKE VIEW 网站自动化查询模块
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
功能：
  1. 从「装潢透视表」Sheet 提取全部采购凭证号
  2. Playwright 启动 Edge → 登录 TKE VIEW（Microsoft SSO，会话持久化）
  3. 逐个 PO：跳转搜索页 → 填PO号 → 查找 → 点击请求ID → 详情页
  4. 详情页提取：项目名称 / 梯台明细合计数量 / 下载全部附件
  5. 附件分类：FLD 开头 → downloads/FLD文件/<PO号>/
               非 FLD   → downloads/无FLD文件/<PO号>/
  6. FLD .msg 解析审批价格 → 计算差异 → 写回 Excel

会话管理：
  → 首次运行需手动完成 Microsoft 账户登录
  → 登录后自动保存 cookies + localStorage 到 auth.json
  → 后续运行加载 auth.json，跳过登录
"""

import os
import re
import json
import sys
import time
import subprocess
import traceback

# ═══════════════════ 配置 ═══════════════════

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(SCRIPT_DIR, "EXPORT1 2 1.xlsx")
TARGET_SHEET = "装潢透视表"
HOME_URL = "https://view.tkelevator.com.cn/"
SEARCH_URL = "https://view.tkelevator.com.cn/vivid/niops/purchasing/materials/search"
AUTH_FILE = os.path.join(SCRIPT_DIR, "auth.json")
RUN_LOG_FILE = os.path.join(SCRIPT_DIR, "automation_run_log.txt")
DOWNLOAD_DIR = os.path.join(SCRIPT_DIR, "downloads")
FLD_DIR = os.path.join(DOWNLOAD_DIR, "FLD文件")
NO_FLD_DIR = os.path.join(DOWNLOAD_DIR, "无FLD文件")
PAGE_SCAN_FILE = os.path.join(SCRIPT_DIR, "page_scan.txt")
PDF_PARSE_TIMEOUT = 45
OCR_TIMEOUT = 60
TARGET_MSG_PREFIXES = ("fld",)
APPROVAL_PRICE_LABEL = "\u5ba1\u6279\u4ef7\u683c"

# ── 测试模式 ──
TEST_MODE = False       # True=仅处理前N个PO, False=全量
TEST_LIMIT = 7           # 测试模式下处理的PO数量


# ═══════════════════ PO 号提取 ═══════════════════

def extract_all_pos_from_pivot(excel_path=None, sheet_name=None):
    """
    从透视表 Sheet 中提取所有有效的采购凭证号（去重）。
    采购凭证为 7~12 位纯数字。
    """
    import openpyxl

    path = excel_path or EXCEL_FILE
    name = sheet_name or TARGET_SHEET

    if not os.path.exists(path):
        return []

    wb = openpyxl.load_workbook(path, data_only=True)
    if name not in wb.sheetnames:
        wb.close()
        return []

    ws = wb[name]
    seen = set()
    pos = []

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            val = str(cell.value or "").strip()
            if val and re.match(r'^\d{7,12}$', val) and val not in seen:
                seen.add(val)
                pos.append(val)

    wb.close()
    return pos


# ═══════════════════ 辅助函数 ═══════════════════

def _col_letter(idx):
    """1→A, 26→Z, 27→AA, ..."""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _is_on_login_page(page):
    """判断当前页面是否为登录页面。"""
    url = page.url.lower()
    return (
        "login.php" in url
        or "login.microsoftonline" in url
        or "microsoft" in url
    )


# ═══════════════════ 元素定位 ═══════════════════

def _locate_po_input(page, log):
    """多策略定位「PO号」输入框，返回第一个可见 input 或 None。"""
    selectors = [
        "//label[contains(text(),'PO号')]/following-sibling::*/descendant::input[1]",
        "//label[contains(text(),'PO号')]/following::input[1]",
        "//td[contains(text(),'PO')]/following-sibling::td//input[1]",
        "//th[contains(text(),'PO')]/following-sibling::td//input[1]",
        "//span[contains(text(),'PO号')]/ancestor::td/following-sibling::td//input[1]",
        "input[placeholder*='PO']",
        "input[placeholder*='po']",
        "input[name*='po']",
        "input[name*='PO']",
        "input[name*='purchase']",
        "input[id*='po']",
        "input[id*='PO']",
        "//label[contains(text(),'PO')]/following::input[1]",
        "//*[contains(text(),'PO号')]/following::input[1]",
    ]
    for sel in selectors:
        try:
            el = page.locator(sel).first
            if el.is_visible(timeout=800):
                log(f"  ✓ PO输入框定位成功（{sel[:50]}...）")
                return el
        except Exception:
            continue
    return None


def _locate_search_button(page, log):
    """多策略定位「查找」按钮。找不到时诊断页面所有可见按钮。"""
    selectors = [
        "button:has-text('查找')",
        "//button[contains(text(),'查找')]",
        "//button[normalize-space(text())='查找']",
        "//a[contains(text(),'查找')]",
        "//input[@type='submit'][@value='查找']",
        "//input[@type='button'][@value='查找']",
        "//*[@role='button'][contains(text(),'查找')]",
        "button:has-text('查')",
        "//button[contains(@class,'search')]",
        "//button[contains(@class,'btn')][contains(text(),'查')]",
        "//*[contains(@class,'el-button')][contains(text(),'查找')]",
    ]
    for sel in selectors:
        try:
            el = page.locator(sel).first
            if el.is_visible(timeout=800):
                log(f"  ✓ 「查找」按钮定位成功（{sel[:50]}）")
                return el
        except Exception:
            continue

    # ── 诊断：列出所有可见按钮 ──
    try:
        all_btns = page.locator("button:visible").all()
        log(f"[诊断] 共 {len(all_btns)} 个可见按钮：")
        for i, btn in enumerate(all_btns[:15]):
            try:
                txt = (btn.inner_text() or "").strip()[:40]
                cls = (btn.get_attribute("class") or "")[:40]
                log(f"  [{i}] text='{txt}' | class='{cls}'")
            except Exception:
                pass
        input_btns = page.locator("input[type='submit']:visible, input[type='button']:visible").all()
        if input_btns:
            log(f"[诊断] 共 {len(input_btns)} 个可见 input 按钮：")
            for i, ib in enumerate(input_btns[:10]):
                try:
                    v = (ib.get_attribute("value") or "")[:40]
                    log(f"  [{i}] value='{v}'")
                except Exception:
                    pass
    except Exception:
        pass

    return None


def _wait_for_search_results(page, log, max_retries=2):
    """等待搜索结果出现，支持重试。"""
    wait_selectors = [
        "//th[contains(text(),'请求ID')]",
        "//a[contains(text(),'LM')]",
        "table:visible",
    ]
    for attempt in range(max_retries):
        if attempt > 0:
            log(f"  搜索结果未出现，第 {attempt + 1} 次重试...")
            try:
                search_btn = _locate_search_button(page, log)
                if search_btn:
                    search_btn.click()
            except Exception:
                pass
            time.sleep(3)
        time.sleep(2)
        for sel in wait_selectors:
            try:
                page.wait_for_selector(sel, state="visible", timeout=8000)
                log("  ✓ 检测到搜索结果")
                return True
            except Exception:
                continue
        time.sleep(3)
    return False


# ═══════════════════ 请求ID点击 & 导航 ═══════════════════

def _click_request_id(page, context, log):
    """
    在搜索结果中找到请求ID并点击，等待新标签页打开。
    返回 (请求ID文本, 详情页page对象 或 None)。
    """
    request_el = None
    selectors = [
        "//a[contains(text(),'LM')]",
        "//td//a",
        "//a[contains(@href,'request')]",
        "//td[contains(text(),'LM')]",
        "//span[contains(text(),'LM')]",
        "//*[contains(text(),'LM80')]",
    ]
    for sel in selectors:
        try:
            el = page.locator(sel).first
            if el.is_visible(timeout=1000):
                text = el.inner_text().strip()
                if len(text) >= 6:
                    request_el = el
                    log(f"  找到请求ID: {text} (selector: {sel[:40]})")
                    break
        except Exception:
            continue

    # ── 诊断：列出搜索页所有链接 ──
    if request_el is None:
        try:
            all_links = page.locator("a:visible").all()
            log(f"[诊断] 搜索结果页共 {len(all_links)} 个可见链接：")
            for i, link in enumerate(all_links[:15]):
                try:
                    txt = (link.inner_text() or "").strip()[:50]
                    href = (link.get_attribute("href") or "")[:60]
                    log(f"  [{i}] text='{txt}' | href='{href}'")
                except Exception:
                    pass
            tables = page.locator("table:visible").all()
            log(f"[诊断] 共 {len(tables)} 个可见表格")
        except Exception:
            pass

    if request_el is None:
        log("  ⚠ 未找到请求ID链接")
        return ("", None)

    rid_text = request_el.inner_text().strip()
    current_pages = list(context.pages)
    current_url = page.url

    # 多策略点击
    click_methods = [
        ("标准点击", lambda: request_el.click()),
        ("JS点击", lambda: request_el.evaluate("el => el.click()")),
        ("dispatchEvent", lambda: request_el.dispatch_event("click")),
        ("父元素点击", lambda: request_el.locator("..").click()),
    ]

    for method_name, method_fn in click_methods:
        try:
            log(f"  尝试点击: {method_name}")
            method_fn()
            for _ in range(7):
                time.sleep(1.5)
                # 检测新标签页
                new_pages = [p for p in context.pages if p not in current_pages]
                if new_pages:
                    np = new_pages[0]
                    np.bring_to_front()
                    try:
                        np.wait_for_load_state("domcontentloaded", timeout=15000)
                    except Exception:
                        pass
                    log(f"  ✓ {method_name} 成功（新标签页）")
                    return (rid_text, np)
                # 检测 URL 变化
                if page.url != current_url and "search" not in page.url.lower():
                    try:
                        page.wait_for_load_state("domcontentloaded", timeout=15000)
                    except Exception:
                        pass
                    log(f"  ✓ {method_name} 成功（URL变化）")
                    return (rid_text, page)
                # 检测详情页特征文本
                for marker in ["VIEW编号", "项目经理", "梯台明细", "物料供应商"]:
                    try:
                        if page.locator(f"//*[contains(text(),'{marker}')]").is_visible(timeout=500):
                            log(f"  ✓ {method_name} 成功（检测到「{marker}」）")
                            return (rid_text, page)
                    except Exception:
                        continue
            log(f"  {method_name} 未检测到导航，尝试下一个...")
        except Exception as e:
            log(f"  ✗ {method_name} 异常: {e}")

    log("  ⚠ 所有点击方法均未触发导航")
    return (rid_text, None)


# ═══════════════════ 详情页数据提取 ═══════════════════

def _extract_project_name(page, log):
    """从详情页文本中提取「项目名称」字段值。"""
    try:
        body_text = page.inner_text("body")
    except Exception:
        return None
    if not body_text:
        return None

    lines = [l.strip() for l in body_text.split("\n") if l.strip()]

    # 搜索 "项目名称" 标签，取紧随其后的中文值
    for i, line in enumerate(lines):
        if "项目名称" not in line:
            continue
        # 尝试在当前行取「项目名称」之后的文本
        m = re.search(r'项目名称\s*[：:]*\s*([\u4e00-\u9fff][\u4e00-\u9fff\w（）()\-·\s]+)', line)
        if m:
            val = m.group(1).strip()
            if len(val) > 2:
                log(f"  项目名称（同行）: {val[:60]}")
                return val
        # 尝试下一行
        if i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if next_line and not next_line.startswith("项目") and any('\u4e00' <= c <= '\u9fff' for c in next_line):
                log(f"  项目名称（下行）: {next_line[:60]}")
                return next_line

    # 兜底：正则搜索
    m = re.search(r'项目名称\s*[：:]*\s*([\u4e00-\u9fff][\u4e00-\u9fff\w（）()\-·]+)', body_text)
    if m:
        val = m.group(1).strip()
        if len(val) > 2:
            log(f"  项目名称（正则）: {val[:60]}")
            return val

    log("  ⚠ 未提取到项目名称")
    return None


def _extract_order_qty(page, log):
    """
    从详情页「梯台明细」表中提取合计行的数量值。
    策略1：找表格→定位「数量」列→定位「合计」行→取交叉单元格。
    策略2（兜底）：找不到「合计」行时，取表格末行的数量列值。
    """
    try:
        # 滚动到梯台明细区域
        try:
            target = page.locator("//*[contains(text(),'梯台明细')]").first
            target.scroll_into_view_if_needed()
            time.sleep(0.5)
        except Exception:
            pass

        for table in page.locator("table").all():
            try:
                if not table.is_visible():
                    continue
                tbl_text = table.inner_text()
                if "数量" not in tbl_text:
                    continue

                # 找「数量」列索引
                qty_col = None
                ths = table.locator("th").all()
                for i, th in enumerate(ths):
                    if (th.inner_text() or "").strip() == "数量":
                        qty_col = i
                        break
                if qty_col is None:
                    first_row = table.locator("tr").first
                    tds = first_row.locator("td").all()
                    for i, td in enumerate(tds):
                        if (td.inner_text() or "").strip() == "数量":
                            qty_col = i
                            break
                if qty_col is None:
                    continue

                # 策略1：找「合计」行 → 取数量列单元格
                for row in table.locator("tr").all():
                    row_text = row.inner_text() or ""
                    if "合计" not in row_text:
                        continue
                    cells = row.locator("td").all()
                    if len(cells) > qty_col:
                        val = cells[qty_col].inner_text().strip()
                        if val:
                            log(f"  梯台明细合计数量: {val}")
                            return val

                # 策略2：兜底 — 末行取值（合计行可能不以"合计"文字出现）
                all_rows = table.locator("tr").all()
                for row in reversed(all_rows):
                    row_text = (row.inner_text() or "").strip()
                    if not row_text:
                        continue
                    cells = row.locator("td, th").all()
                    if len(cells) > qty_col:
                        val = cells[qty_col].inner_text().strip()
                        if val and re.match(r'^[\d,.]+$', val):
                            log(f"  梯台明细合计数量（末行兜底）: {val}")
                            return val
                    break  # 只尝试最后一个非空行
            except Exception:
                continue
    except Exception:
        pass

    log("  ⚠ 未提取到梯台明细合计数量")
    return None


def _extract_target_msg_name(text):
    """Return a FLD .msg filename from link text or URL-like text."""
    if not text:
        return ""
    value = str(text).strip()
    try:
        from urllib.parse import unquote
        value = unquote(value)
    except Exception:
        pass

    # Only filenames starting with FLD are valid approval mail files.
    pattern = r"(?i)(?:^|[\\/\s=&?])(fld[^\\/\r\n<>]*?\.msg)"
    match = re.search(pattern, value)
    if not match:
        return ""
    name = match.group(1).strip().strip("'\"")
    return os.path.basename(name)


def _is_target_msg_filename(name):
    """True only for .msg files whose basename starts with FLD."""
    extracted = _extract_target_msg_name(name)
    if not extracted:
        return False
    lower = extracted.lower()
    return lower.endswith(".msg") and lower.startswith(TARGET_MSG_PREFIXES)


def _extract_any_file_name(text):
    """Extract a filename (with extension) from link text or URL-like text (any file type)."""
    if not text:
        return ""
    value = str(text).strip()
    try:
        from urllib.parse import unquote
        value = unquote(value)
    except Exception:
        pass
    # Match filenames with common extensions
    pattern = r"(?:^|[\\/\s=&?])([^\\/\r\n<>]*?\.[a-zA-Z0-9]{2,5})(?:\s|$|[\\/?&])"
    match = re.search(pattern, value)
    if not match:
        return ""
    name = match.group(1).strip().strip("'\"")
    return os.path.basename(name)


def _link_file_name(link):
    """Extract filename from a link element (any file type, not just FLD .msg)."""
    candidates = []
    try:
        candidates.append(link.inner_text().strip())
    except Exception:
        pass
    for attr in ("download", "title", "aria-label", "href"):
        try:
            value = link.get_attribute(attr)
            if value:
                candidates.append(value)
        except Exception:
            pass
    for candidate in candidates:
        name = _extract_any_file_name(candidate)
        if name:
            return name
    return ""


def _classify_file(file_name):
    """Classify a file as 'fld' (starts with fld) or 'non_fld'."""
    lower = (file_name or "").lower()
    if lower.startswith("fld"):
        return "fld"
    return "non_fld"


def _extract_project_from_fld_name(filename, log):
    """
    兜底：从 FLD 文件名中提取项目名称。
    文件名格式: FLD新物料价格审批--<项目名>--<分类>--价格审批回复.msg
    策略：-- 分割后，取中间片段中「包含中文且最长」的作为项目名。
    """
    name = os.path.basename(str(filename or ""))
    parts = name.split("--")
    if len(parts) < 3:
        return None
    best = None
    for part in parts[1:-1]:  # 跳过首段(FLD前缀)和末段(价格审批回复.msg)
        part = part.strip()
        if any('\u4e00' <= c <= '\u9fff' for c in part):
            if best is None or len(part) > len(best):
                best = part
    if best:
        log(f"  项目名称（FLD文件名兜底）: {best[:60]}")
    return best


def _target_msg_name_from_link(link):
    candidates = []
    try:
        candidates.append(link.inner_text().strip())
    except Exception:
        pass
    for attr in ("download", "title", "aria-label", "href"):
        try:
            value = link.get_attribute(attr)
            if value:
                candidates.append(value)
        except Exception:
            pass

    for candidate in candidates:
        name = _extract_target_msg_name(candidate)
        if name:
            return name
    return ""


def _collect_all_attachment_links(page):
    """Collect ALL attachment links from the detail page (not just FLD .msg)."""
    selectors = [
        "//th[contains(text(),'\u652f\u6301\u6587\u4ef6')]/ancestor::table//a",
        "//td[contains(text(),'\u652f\u6301\u6587\u4ef6')]/ancestor::table//a",
        "a",
    ]
    found = []
    seen = set()
    for selector in selectors:
        try:
            links = page.locator(selector).all()
        except Exception:
            continue
        for link in links:
            try:
                if not link.is_visible():
                    continue
            except Exception:
                pass
            name = _link_file_name(link)
            if not name:
                continue
            try:
                href = link.get_attribute("href") or ""
            except Exception:
                href = ""
            key = (name.lower(), href)
            if key in seen:
                continue
            seen.add(key)
            found.append((link, name))
    return found


def _download_support_files(page, log, po):
    """
    下载详情页中的全部附件，存入同一目录。
    有 FLD 附件 → downloads/FLD文件/<PO号>/（全部文件，含非FLD）
    无 FLD 附件 → downloads/无FLD文件/<PO号>/
    返回: {fld_files: [...], non_fld_files: [...]}
    """
    fld_target = os.path.join(FLD_DIR, str(po))
    non_fld_target = os.path.join(NO_FLD_DIR, str(po))

    # 滚动到审批记录区域
    try:
        target = page.locator("//*[contains(text(),'审批记录')]").first
        target.scroll_into_view_if_needed()
        time.sleep(0.5)
    except Exception:
        pass

    fld_files = []
    non_fld_files = []
    all_links = _collect_all_attachment_links(page)
    if not all_links:
        log("  未发现任何附件链接。")
        return {"fld_files": fld_files, "non_fld_files": non_fld_files}

    # 有任一 FLD 文件 → 全部存入 FLD文件夹；否则存入 无FLD文件夹
    has_any_fld = any(_classify_file(name) == "fld" for _, name in all_links)
    if has_any_fld:
        target_dir = fld_target
        dir_label = "FLD文件"
    else:
        target_dir = non_fld_target
        dir_label = "无FLD文件"
    os.makedirs(target_dir, exist_ok=True)

    for link, file_name in all_links:
        try:
            category = _classify_file(file_name)
            log(f"  正在下载: {file_name} [{category}]")
            with page.expect_download(timeout=30000) as download_info:
                link.click()
            download = download_info.value

            safe_name = file_name.replace("/", "_").replace("\\", "_").replace(":", "_")
            suggested = download.suggested_filename or ""
            if suggested:
                suggested = os.path.basename(suggested)
                safe_name = suggested.replace("/", "_").replace("\\", "_").replace(":", "_")

            save_path = os.path.join(target_dir, safe_name)
            download.save_as(save_path)
            if category == "fld":
                fld_files.append(safe_name)
            else:
                non_fld_files.append(safe_name)
            log(f"    ✓ 已保存: {dir_label}/{po}/{safe_name}")
            time.sleep(0.5)
        except Exception as e:
            log(f"    ✗ 下载失败: {str(e)[:100]}")

    log(f"  附件下载完成: FLD {len(fld_files)} 个, 非FLD {len(non_fld_files)} 个 → {dir_label}")
    return {"fld_files": fld_files, "non_fld_files": non_fld_files}


def _has_fld_msg_attachment(page, log):
    """
    检查详情页附件中是否存在文件名以 FLD 开头的 .msg 文件。
    多策略扫描，返回 True/False。
    """
    # 先滚动到审批记录区域（懒加载）
    try:
        target = page.locator("//*[contains(text(),'审批记录')]").first
        target.scroll_into_view_if_needed()
        time.sleep(1)
    except Exception:
        pass

    all_links = _collect_all_attachment_links(page)
    for _, name in all_links:
        if _classify_file(name) == "fld":
            log(f"  ✓ 发现 FLD 附件: {name[:80]}")
            return True

    log("  未发现 FLD 开头的附件")
    return False


def _has_any_msg_attachment(page):
    """快速扫描是否存在任何 .msg 附件（不管是不是 FLD）"""
    try:
        links = page.locator("a").all()
        for link in links[:50]:
            try:
                if ".msg" in link.inner_text().strip().lower():
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


# ═══════════════════ FLD .msg 审批价格提取 ═══════════════════

def _compact_text(value):
    return re.sub(r"\s+", "", str(value or ""))


def _extract_amount_from_text(value):
    text = str(value or "").replace("\xa0", " ")
    amount_pattern = r"(?<![\d])(?:\d{1,3}(?:,\d{3})+|\d{4,}|\d{1,3})(?:\.\d+)?(?![\d])"
    for match in re.finditer(amount_pattern, text):
        raw = match.group(0)
        cleaned = raw.replace(",", "")
        try:
            if float(cleaned) > 100:
                return cleaned
        except ValueError:
            continue
    return None


def _extract_approval_price_from_tables(soup, log):
    for table_no, table in enumerate(soup.find_all("table"), 1):
        rows = []
        for tr in table.find_all("tr"):
            cells = [
                cell.get_text(" ", strip=True)
                for cell in tr.find_all(["td", "th"])
            ]
            if cells:
                rows.append(cells)

        for row_idx, row in enumerate(rows):
            price_cols = [
                col_idx
                for col_idx, cell_text in enumerate(row)
                if APPROVAL_PRICE_LABEL in _compact_text(cell_text)
            ]
            if not price_cols:
                continue

            for col_idx in price_cols:
                # Vertical/key-value table: label and value can be in the same row.
                for candidate in row[col_idx + 1: col_idx + 3]:
                    amount = _extract_amount_from_text(candidate)
                    if amount:
                        log(f"  [MSG] 表格{table_no} 同行提取审批价格: {amount}")
                        return amount

                # Horizontal table: label is a header, value is in the next data row.
                for next_row in rows[row_idx + 1:]:
                    if col_idx >= len(next_row):
                        continue
                    amount = _extract_amount_from_text(next_row[col_idx])
                    if amount:
                        log(f"  [MSG] 表格{table_no} 第{col_idx + 1}列审批价格: {amount}")
                        return amount
    return None


def _extract_approval_price_from_plain_text(text, log):
    lines = [_compact_text(line) for line in str(text or "").splitlines()]
    lines = [line for line in lines if line]
    for idx, line in enumerate(lines):
        if APPROVAL_PRICE_LABEL not in line:
            continue

        # Only use this fallback when the label is isolated. If a whole header row
        # is flattened into one line, the next line may start with audit IDs.
        if len(line) > len(APPROVAL_PRICE_LABEL) + 4:
            continue

        amount = _extract_amount_from_text(line.replace(APPROVAL_PRICE_LABEL, "", 1))
        if amount:
            log(f"  [MSG] 纯文本同行提取审批价格: {amount}")
            return amount

        for next_line in lines[idx + 1: idx + 5]:
            amount = _extract_amount_from_text(next_line)
            if amount:
                log(f"  [MSG] 纯文本下一行提取审批价格: {amount}")
                return amount
    return None


def _extract_approval_price_from_msg(msg_path, log):
    """
    解析 FLD .msg 邮件，从正文表格中提取「审批价格」。
    返回金额字符串（如 "246900"），失败返回 None。
    """
    import extract_msg
    from bs4 import BeautifulSoup

    log(f"  [MSG] 解析: {os.path.basename(msg_path)}")
    try:
        msg = extract_msg.Message(msg_path)
        html = msg.htmlBody or ""
        body_text = msg.body or ""
        msg.close()

        if not html.strip() and not body_text.strip():
            log("  [MSG] 邮件正文为空（HTML和纯文本均无内容）")
            return None

        # 优先用 HTML 解析，回退到纯文本
        if html.strip():
            source = html
            log(f"  [MSG] 正文来源: HTML, 长度: {len(source)}")
        else:
            source = body_text
            log(f"  [MSG] 正文来源: 纯文本(无HTML), 长度: {len(source)}")

        soup = BeautifulSoup(source, "html.parser")
        
        # 诊断：输出解析后的纯文本前 200 字符
        plain_preview = soup.get_text()[:200]
        log(f"  [MSG] 文本预览: {plain_preview}")
        
        amount = _extract_approval_price_from_tables(soup, log)
        if amount:
            return amount

        amount = _extract_approval_price_from_plain_text(soup.get_text("\n"), log)
        if amount:
            return amount

        log("  [MSG] 未找到「审批价格」")
        return None
    except ImportError:
        log("  [MSG] extract-msg 模块未安装，请运行: pip install extract-msg beautifulsoup4")
        return None
    except Exception as e:
        log(f"  [MSG] 解析失败: {e}")
        return None


# ═══════════════════ PDF 金额提取 ═══════════════════

def _extract_total_from_pdf(pdf_path, log):
    """
    读取 PDF 文件，提取「总计（不含税）」后的金额。
    使用子进程执行 pdfplumber 解析，防止卡死主流程。
    """
    log(f"  [PDF] 解析: {os.path.basename(pdf_path)}")
    worker_code = r'''
import json, sys, pdfplumber, re

pdf_path = sys.argv[1]
logs = []

def emit(status, total=None, error=""):
    print(json.dumps({
        "status": status, "total": total, "error": error, "logs": logs,
    }, ensure_ascii=False))

try:
    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            logs.append(f"  [PDF] 第 {page_no}/{len(pdf.pages)} 页")
            # 归一化旋转（摆正颠倒/横排页面）
            rot = getattr(page, 'rotation', 0) or 0
            if rot != 0:
                logs.append(f"  [PDF] 检测到旋转 {rot}°，正在摆正...")
                try:
                    page = page.rotate(360 - rot)
                except Exception:
                    pass
            text = page.extract_text() or ""
            # 补充表格文字（总计可能在表格中）
            try:
                tables = page.extract_tables()
                for tbl in tables:
                    for row in tbl:
                        for cell in row:
                            if cell and ("总计" in str(cell) or "小计" in str(cell) or "合计" in str(cell)):
                                text += "\n" + str(cell)
            except Exception:
                pass
            if not text.strip():
                continue
            lines = text.split("\n")
            for i, line in enumerate(lines):
                # 匹配关键词：精确匹配优先，"总计"兜底
                if "总计（不含税）" not in line and "总计(不含税)" not in line and "总计" not in line:
                    continue
                logs.append(f"  [PDF] 匹配行: {line[:150]}")
                # ── 正则提取数字（解决千分位空格/逗号切分BUG）──
                # 匹配数字模式：支持 "9 ,594.00"、"9,594.00"、"6520.00" 等
                candidates = []
                for raw in re.findall(r'[\d][\d, ]*\.?\d*', line):
                    cleaned = raw.replace(",", "").replace(" ", "").replace("\uffe5", "").replace("\xa5", "")
                    try:
                        v = float(cleaned)
                        if v > 100:
                            candidates.append((v, cleaned))
                    except ValueError:
                        continue
                # 也检查下一行
                if i + 1 < len(lines):
                    for raw in re.findall(r'[\d][\d, ]*\.?\d*', lines[i + 1]):
                        cleaned = raw.replace(",", "").replace(" ", "").replace("\uffe5", "").replace("\xa5", "")
                        try:
                            v = float(cleaned)
                            if v > 100:
                                candidates.append((v, cleaned))
                        except ValueError:
                            continue
                if candidates:
                    # 取最大值（排除可能的小额杂数）
                    best_val, best_str = max(candidates, key=lambda x: x[0])
                    logs.append(f"  [PDF] 提取总计: {best_str}")
                    emit("ok", best_str)
                    sys.exit(0)
    emit("ok")
except Exception as e:
    emit("error", error=str(e))
'''

    try:
        completed = subprocess.run(
            [sys.executable, "-c", worker_code, pdf_path],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=PDF_PARSE_TIMEOUT,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.TimeoutExpired:
        log(f"  [PDF] 解析超时（{PDF_PARSE_TIMEOUT}秒），跳过")
        return None
    except Exception as e:
        log(f"  [PDF] 子进程启动失败: {e}")
        return None

    output = (completed.stdout or "").strip()
    if not output:
        stderr = (completed.stderr or "").strip()
        if stderr:
            log(f"  [PDF] 无结果: {stderr[:200]}")
        return None

    try:
        data = json.loads(output.splitlines()[-1])
    except Exception:
        log(f"  [PDF] 解析JSON失败: {output[:200]}")
        return None

    for line in data.get("logs", []):
        log(line)

    if data.get("status") == "error":
        log(f"  [PDF] 读取失败: {data.get('error')}")
        return None

    return data.get("total")


# ═══════════════════ OCR 回退 ═══════════════════

def _extract_total_from_pdf_ocr(pdf_path, log):
    """
    OCR 回退：当 pdfplumber 无法提取文字时，将 PDF 首页转为图片，
    用 Tesseract 中文识别，搜索「总计（不含税）」后的金额。
    返回金额字符串，失败返回 None。
    """
    log(f"  [OCR] 尝试 OCR: {os.path.basename(pdf_path)}")
    worker_code = r'''
import json, sys, re, os

try:
    import pytesseract
    from pdf2image import convert_from_path
except ImportError as e:
    print(json.dumps({"status": "error", "error": f"OCR模块未安装: {e}", "logs": []}, ensure_ascii=False))
    sys.exit(0)

pdf_path = sys.argv[1]
logs = []

def emit(status, total=None, error=""):
    print(json.dumps({
        "status": status, "total": total, "error": error, "logs": logs,
    }, ensure_ascii=False))

try:
    # 检测 Tesseract 路径
    tesseract_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR", "tesseract.exe"),
    ]
    tesseract_found = None
    for tp in tesseract_paths:
        if os.path.exists(tp):
            tesseract_found = tp
            pytesseract.pytesseract.tesseract_cmd = tp
            break
    if not tesseract_found:
        # 尝试 PATH 中的 tesseract
        import shutil
        found = shutil.which("tesseract")
        if found:
            pytesseract.pytesseract.tesseract_cmd = found
            tesseract_found = found
    if not tesseract_found:
        emit("error", error="Tesseract OCR 未安装。请从 https://github.com/UB-Mannheim/tesseract/wiki 下载安装，并勾选中文语言包。")
        sys.exit(0)

    logs.append(f"  [OCR] Tesseract: {tesseract_found}")

    # 只转换首页（供应商报价单通常是单页）
    images = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=300)
    if not images:
        emit("error", error="PDF转图片失败")
        sys.exit(0)

    img = images[0]
    logs.append(f"  [OCR] 图片尺寸: {img.size}")

    # 中文 OCR
    text = pytesseract.image_to_string(img, lang="chi_sim+eng", config="--psm 6")
    if not text or not text.strip():
        emit("ok")
        sys.exit(0)

    logs.append(f"  [OCR] 识别文字 {len(text)} 字符")

    # 搜索「总计（不含税）」及其变体
    lines = text.split("\n")
    candidates = []
    for i, line in enumerate(lines):
        # 多种关键词匹配
        if not ("总计" in line or "合计" in line or "总价" in line or "总金额" in line):
            continue
        logs.append(f"  [OCR] 候选行: {line[:150]}")
        # 正则提取本行及下一行的数字
        for check_line in (line, lines[i + 1] if i + 1 < len(lines) else ""):
            for raw in re.findall(r'[\d][\d, ]*\.?\d*', check_line):
                cleaned = raw.replace(",", "").replace(" ", "").replace("\uffe5", "").replace("\xa5", "")
                try:
                    v = float(cleaned)
                    if v > 100:
                        candidates.append((v, cleaned))
                except ValueError:
                    continue

    if candidates:
        best_val, best_str = max(candidates, key=lambda x: x[0])
        logs.append(f"  [OCR] 提取总计: {best_str}")
        emit("ok", best_str)
    else:
        logs.append("  [OCR] 未找到有效金额")
        emit("ok")
except Exception as e:
    emit("error", error=str(e))
'''

    try:
        completed = subprocess.run(
            [sys.executable, "-c", worker_code, pdf_path],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=OCR_TIMEOUT,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.TimeoutExpired:
        log(f"  [OCR] 识别超时（{OCR_TIMEOUT}秒），跳过")
        return None
    except Exception as e:
        log(f"  [OCR] 子进程启动失败: {e}")
        return None

    output = (completed.stdout or "").strip()
    if not output:
        stderr = (completed.stderr or "").strip()
        if stderr:
            log(f"  [OCR] 无结果: {stderr[:200]}")
        return None

    try:
        data = json.loads(output.splitlines()[-1])
    except Exception:
        log(f"  [OCR] 解析JSON失败: {output[:200]}")
        return None

    for line in data.get("logs", []):
        log(line)

    if data.get("status") == "error":
        log(f"  [OCR] 识别失败: {data.get('error')}")
        return None

    return data.get("total")


# ═══════════════════ Excel 读写 ═══════════════════

def _read_order_value_from_pivot(po):
    """从透视表读取指定 PO 的订单净值（列B「求和项:订单净值」）。"""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        if TARGET_SHEET not in wb.sheetnames:
            wb.close()
            return None
        ws = wb[TARGET_SHEET]
        po_str = str(po).strip()
        for row in range(2, ws.max_row + 1):
            if str(ws.cell(row, 1).value or "").strip() == po_str:
                val = ws.cell(row, 2).value
                wb.close()
                return str(val) if val is not None else None
        wb.close()
    except Exception:
        pass
    return None


def _write_to_pivot_excel(po, project_name=None, order_qty=None,
                          support_files=None, total_amount=None, order_diff=None):
    """
    将提取数据写回「装潢透视表」Sheet，按采购凭证行对齐。
    """
    import openpyxl

    if not os.path.exists(EXCEL_FILE):
        return f"Excel 文件不存在: {EXCEL_FILE}"

    wb = openpyxl.load_workbook(EXCEL_FILE)
    if TARGET_SHEET not in wb.sheetnames:
        wb.close()
        return f"Sheet「{TARGET_SHEET}」不存在"

    ws = wb[TARGET_SHEET]

    # 1. 扫描第 1 行定位扩展列
    ext_cols = {}
    for col in range(1, ws.max_column + 1):
        val = str(ws.cell(1, col).value or "").strip()
        if val == "E2E项目名":
            ext_cols["E2E项目名"] = col
        elif val == "E2E订单数量":
            ext_cols["E2E订单数量"] = col
        elif val == "项目总金额(审批)":
            ext_cols["项目总金额(审批)"] = col
        elif val == "订单差异(单独计算)":
            ext_cols["订单差异(单独计算)"] = col
        elif val == "支持文件名(FLD 审批)":
            ext_cols["支持文件名(FLD 审批)"] = col

    if not ext_cols:
        wb.close()
        return "未找到扩展列，请先点击「添加扩展列」按钮"

    # 2. 扫描列A 找匹配 PO 的行
    po_str = str(po).strip()
    target_rows = []
    for row in range(2, ws.max_row + 1):
        val = str(ws.cell(row, 1).value or "").strip()
        if val == po_str:
            target_rows.append(row)

    if not target_rows:
        wb.close()
        return f"在列A中未找到采购凭证 {po_str}"

    # 3. 写入数据
    filled = []
    for tr in target_rows:
        if project_name and "E2E项目名" in ext_cols:
            ws.cell(tr, ext_cols["E2E项目名"], project_name)
            filled.append("E2E项目名")
        if order_qty and "E2E订单数量" in ext_cols:
            try:
                ws.cell(tr, ext_cols["E2E订单数量"], float(str(order_qty).replace(",", "")))
            except ValueError:
                ws.cell(tr, ext_cols["E2E订单数量"], order_qty)
            filled.append("E2E订单数量")
        if total_amount is not None and "项目总金额(审批)" in ext_cols:
            try:
                ws.cell(tr, ext_cols["项目总金额(审批)"], float(str(total_amount).replace(",", "")))
            except ValueError:
                ws.cell(tr, ext_cols["项目总金额(审批)"], total_amount)
            filled.append("项目总金额(审批)")
        if order_diff is not None and "订单差异(单独计算)" in ext_cols:
            try:
                ws.cell(tr, ext_cols["订单差异(单独计算)"], float(str(order_diff).replace(",", "")))
            except ValueError:
                ws.cell(tr, ext_cols["订单差异(单独计算)"], order_diff)
            filled.append("订单差异(单独计算)")
        if support_files and "支持文件名(FLD 审批)" in ext_cols:
            ws.cell(tr, ext_cols["支持文件名(FLD 审批)"], support_files)
            filled.append("支持文件名(FLD 审批)")

    wb.save(EXCEL_FILE)
    wb.close()
    return f"已写入 {len(target_rows)} 行（PO {po_str}）：{', '.join(filled)}"


# ═══════════════════ 单 PO 处理 ═══════════════════

def _process_one_po(po, context, log):
    """
    处理单个 PO 的完整浏览器流程：
    搜索 → 点击请求ID → 提取项目名称/数量 → 下载附件。
    返回 dict 或 None。
    """
    result = {"po": po}

    try:
        po_page = context.new_page()
        po_page.set_default_timeout(30000)

        # ── 1. 导航到搜索页 ──
        po_page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=20000)
        time.sleep(1.5)

        # ── 2. 填入 PO 号 ──
        po_input = _locate_po_input(po_page, log)
        if po_input is None:
            log(f"  ⚠ 无法定位 PO 输入框，跳过")
            return None

        po_input.click()
        po_input.fill("")
        time.sleep(0.2)
        po_input.type(po, delay=80)
        time.sleep(0.3)

        # ── 3. 点击查找 ──
        search_btn = _locate_search_button(po_page, log)
        if search_btn is None:
            # 兜底：按 Enter 键触发搜索
            log("  未找到「查找」按钮，尝试按 Enter 键提交...")
            po_input.press("Enter")
        else:
            search_btn.click()
        log("  已触发搜索...")

        if not _wait_for_search_results(po_page, log):
            log(f"  ⚠ 搜索结果未出现，跳过")
            return None

        # ── 4. 点击请求ID → 详情页 ──
        rid_text, detail_page = _click_request_id(po_page, context, log)
        result["request_id"] = rid_text

        if detail_page is None:
            log(f"  ⚠ 未能进入详情页")
            result["project_name"] = None
            result["order_qty"] = None
            result["fld_files"] = []
            result["non_fld_files"] = []
        else:
            log(f"  详情页已打开，等待加载...")
            time.sleep(2)
            try:
                detail_page.wait_for_load_state("domcontentloaded", timeout=10000)
            except Exception:
                pass

            # 写入 page_scan.txt 供调试
            try:
                detail_text = detail_page.inner_text("body")
                with open(PAGE_SCAN_FILE, "w", encoding="utf-8") as f:
                    f.write(detail_text)
            except Exception:
                pass

            # ── 5. 下载全部附件（按 PO 号分文件夹，FLD/非FLD 分类）──
            dl_result = _download_support_files(detail_page, log, po)
            result["fld_files"] = dl_result.get("fld_files", [])
            result["non_fld_files"] = dl_result.get("non_fld_files", [])
            log(f"  附件下载: FLD {len(result['fld_files'])} 个, 非FLD {len(result['non_fld_files'])} 个")

            has_fld = len(result["fld_files"]) > 0

            # ── 6. 有 FLD 附件时：提取项目名称和梯台明细数量（用于 Excel 写回）──
            if has_fld:
                result["project_name"] = _extract_project_name(detail_page, log)
                # 兜底：网页提取失败时从 FLD 文件名提取项目名称
                if not result["project_name"] and result["fld_files"]:
                    result["project_name"] = _extract_project_from_fld_name(
                        result["fld_files"][0], log
                    )
                result["order_qty"] = _extract_order_qty(detail_page, log)
            else:
                log(f"  ⓘ PO {po} 无 FLD 附件，跳过项目名称/数量提取")
                result["project_name"] = None
                result["order_qty"] = None

            # 关闭详情页
            if detail_page != po_page:
                try:
                    detail_page.close()
                except Exception:
                    pass

        # ★ 保留搜索页给下一个 PO 复用，不关闭页面
        log(f"  ✓ PO {po} 网页阶段完成")
        return result

    except Exception as e:
        log(f"  ✗ PO {po} 异常: {e}")
        return None


def _finalize_po_result(result, log):
    """
    本地处理：PDF解析 + 差异计算 + Excel写回。
    不依赖浏览器，避免 PDF 解析阻塞后续 PO 查询。
    """
    po = result.get("po")

    # ── FLD .msg 解析 → 提取审批价格 ──
    total_amount = None
    fld_files = result.get("fld_files", [])
    if fld_files:
        log("  解析 FLD 附件...")
        fld_target = os.path.join(FLD_DIR, str(po))

        # 找第一个 FLD .msg 文件
        fld_msg = None
        for fname in fld_files:
            if fname.lower().endswith(".msg"):
                fld_msg = fname
                break

        if fld_msg:
            fpath = os.path.join(fld_target, fld_msg)
            if os.path.exists(fpath):
                total_amount = _extract_approval_price_from_msg(fpath, log)
            else:
                log(f"  FLD .msg 文件不存在: {fld_msg}")
        else:
            log("  FLD 附件中未找到 .msg 文件")

        if total_amount:
            log(f"  审批价格: {total_amount}")
        else:
            log("  未从 FLD .msg 中解析到审批价格")
    else:
        log("  无 FLD 附件，跳过审批价格解析")
    result["total_amount"] = total_amount

    # ── 读取订单净值 → 计算差异 ──
    order_value = _read_order_value_from_pivot(po)
    result["order_value"] = order_value
    order_diff = None

    if total_amount is not None and order_value is not None:
        try:
            ta = float(str(total_amount).replace(",", ""))
            ov = float(str(order_value).replace(",", ""))
            order_diff = round(ta - ov, 2)
            log(f"  订单差异: {order_diff} (={ta} - {ov})")
        except (ValueError, TypeError):
            log("  ⚠ 差异计算失败")
    result["order_diff"] = order_diff

    # ── 统计非 PDF 附件 ──
    non_pdf = [f for f in fld_files if not f.lower().endswith(".pdf")]
    result["non_pdf_count"] = len(non_pdf)

    # ── 写回 Excel ──
    support_files_str = "; ".join(fld_files) if fld_files else ""
    diff_str = str(order_diff) if order_diff is not None else None
    total_str = str(total_amount) if total_amount is not None else None

    excel_msg = _write_to_pivot_excel(
        po,
        project_name=result.get("project_name"),
        order_qty=result.get("order_qty"),
        support_files=support_files_str,
        total_amount=total_str,
        order_diff=diff_str,
    )
    log(f"  {excel_msg}")
    log(f"  ✓ PO {po} 本地处理完成")
    return result


# ═══════════════════ 主入口 ═══════════════════

def open_website_and_search(status_callback=None):
    """
    主流程：
      1. 从透视表提取全部 PO 号
      2. 启动 Edge → 登录 → 循环处理每个 PO（搜索/提取/下载）
      3. 所有 PO 网页阶段完成后 → 本地解析 PDF → 写回 Excel
      4. 汇总结果
    """
    with open(RUN_LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"自动化运行日志 {time.strftime('%Y-%m-%d %H:%M:%S')}\n")

    def log(msg):
        if status_callback:
            status_callback(msg)
        try:
            with open(RUN_LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"{time.strftime('%H:%M:%S')} {msg}\n")
        except Exception:
            pass

    # ── 1. 提取全部 PO ──
    log("正在从透视表提取全部采购凭证号...")
    all_pos = extract_all_pos_from_pivot()
    if not all_pos:
        return False, "透视表中未找到任何采购凭证号。\n请确认透视表已生成。"
    total = len(all_pos)
    log(f"共提取到 {total} 个采购凭证号")

    # ── 2. 启动浏览器 ──
    log("正在启动 Edge 浏览器...")
    from playwright.sync_api import sync_playwright
    playwright = sync_playwright().start()

    try:
        browser = playwright.chromium.launch(
            channel="msedge",
            headless=False,
            args=["--start-maximized"],
        )

        context_kwargs = {"no_viewport": True, "accept_downloads": True}
        if os.path.exists(AUTH_FILE):
            log("检测到已保存的登录状态，尝试自动登录...")
            try:
                context = browser.new_context(**{**context_kwargs, "storage_state": AUTH_FILE})
                log("已加载登录状态。")
            except Exception:
                os.remove(AUTH_FILE)
                context = browser.new_context(**context_kwargs)
                log("登录状态文件已过期，需要重新登录。")
        else:
            context = browser.new_context(**context_kwargs)

        page = context.new_page()

        # ── 3. 登录检测 ──
        log("正在打开 TKE VIEW...")
        page.goto(HOME_URL, wait_until="domcontentloaded", timeout=30000)
        time.sleep(2)

        if _is_on_login_page(page):
            log("════════════════════════════════")
            log("⚠ 需要登录 — 请在浏览器中完成 Microsoft 账户验证")
            log("  登录完成后程序将自动继续...")
            log("════════════════════════════════")
            try:
                page.wait_for_url(
                    lambda url: "login.php" not in url.lower()
                    and "login.microsoftonline" not in url.lower(),
                    timeout=300000
                )
                log("✅ 登录成功！正在保存登录状态...")
                context.storage_state(path=AUTH_FILE)
                log("登录状态已保存。")
                time.sleep(2)
            except Exception:
                return False, "登录超时（超过 5 分钟），请重试。"
        else:
            log("✅ 已登录，会话有效。")

        try:
            page.close()
        except Exception:
            pass

        # ── 4. 循环处理每个 PO ──
        if TEST_MODE:
            test_pos = all_pos[:TEST_LIMIT]
            total = len(test_pos)
            log(f"[测试模式] 仅处理前 {total} 个 PO")
        else:
            test_pos = all_pos
            total = len(test_pos)
            log(f"[正式模式] 处理全部 {total} 个 PO")

        success_count = 0
        fail_count = 0
        failed_pos = []
        web_results = []
        crashed = False

        for i, po in enumerate(test_pos):
            log("")
            log(f"━━━ [{i+1}/{total}] PO: {po} ━━━")
            try:
                result = _process_one_po(po, context, log)
                if result is not None:
                    success_count += 1
                    web_results.append(result)
                    log(f"✓ [{i+1}/{total}] PO {po} 网页阶段完成")
                else:
                    fail_count += 1
                    failed_pos.append(po)
                    log(f"⚠ [{i+1}/{total}] PO {po} 未完成")
            except Exception as e:
                log(f"  ✗ 浏览器异常: {e}")
                fail_count += 1
                failed_pos.append(po)
                error_str = str(e).lower()
                if any(kw in error_str for kw in ("epipe", "closed", "target closed", "connection")):
                    log("⚠ 浏览器连接断开，停止后续处理。")
                    crashed = True
                    break

            # PO 之间短暂休息
            time.sleep(0.5)

        # 保留浏览器供查看
        if not crashed:
            log("网页阶段完成，浏览器窗口保留供查看。")

        # ── 5. 本地解析 PDF + 写回 Excel ──
        finalize_success = 0
        finalize_fail = 0
        no_amount_pos = []
        if web_results:
            log("")
            log(f"开始本地解析并写回 Excel（共 {len(web_results)} 个 PO）...")
            for j, result in enumerate(web_results, 1):
                po = result.get("po")
                log("")
                log(f"━━━ [本地 {j}/{len(web_results)}] PO: {po} ━━━")
                try:
                    if not result.get("fld_files"):
                        log(f"  ⏭ PO {po} 无 FLD 附件，跳过本地解析")
                        continue
                    _finalize_po_result(result, log)
                    finalize_success += 1
                    if result.get("total_amount") is None:
                        no_amount_pos.append(result.get("po", "?"))
                except Exception as e:
                    finalize_fail += 1
                    log(f"  ✗ PO {po} 本地处理异常: {e}")

        # ── 6. 汇总 ──
        log("")
        if crashed:
            log(f"⚠ 浏览器异常中断！网页阶段完成: {success_count} / 失败: {fail_count}")
        else:
            log(f"✅ 全部完成！网页阶段: 成功 {success_count} / 失败 {fail_count} / 总数 {total}")
        log(f"本地解析写回: 成功 {finalize_success} / 失败 {finalize_fail}")

        auth_info = "登录状态已保存，下次自动登录。" if os.path.exists(AUTH_FILE) else ""
        summary = (
            f"✅ 全部查询完成！\n\n"
            f"网页查询成功: {success_count} / 失败: {fail_count} / 总数: {total}\n"
            f"本地解析写回成功: {finalize_success} / 失败: {finalize_fail}"
        )
        if no_amount_pos:
            summary += f"\n\n⚠ {len(no_amount_pos)} 个项目未获取到审批金额："
            for p in no_amount_pos:
                summary += f"\n  - {p}"
        no_fld_download_pos = [r.get("po") for r in web_results if not r.get("fld_files")]
        if no_fld_download_pos:
            summary += f"\n\nℹ {len(no_fld_download_pos)} 个PO无FLD附件（仅下载非FLD文件至「无FLD文件」目录）："
            for p in no_fld_download_pos:
                summary += f"\n  - {p}"
        if failed_pos:
            summary += f"\n\n失败 PO ({len(failed_pos)} 个):\n"
            for fp in failed_pos[:20]:
                summary += f"  - {fp}\n"
            if len(failed_pos) > 20:
                summary += f"  ... 等共 {len(failed_pos)} 个\n"
        summary += f"\n\n{auth_info}"
        return True, summary

    except Exception as e:
        return False, f"网页操作异常：\n{e}\n\n{traceback.format_exc()}"
