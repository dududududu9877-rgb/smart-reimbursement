#!/usr/bin/env python3
"""报销导出：读 state.json → 填模板 → 输出 Excel 到桌面

用法:
  python3 export_expense.py 4月     # 指定月份
  python3 export_expense.py 4       # 纯数字也接受
  python3 export_expense.py         # 默认当前月份
"""

import json
import os
import sys
import base64
import tempfile
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XlImage
from openpyxl.utils import get_column_letter

APP_DIR = Path(__file__).parent
TEMPLATE = APP_DIR / "template.xlsx"
DESKTOP = Path.home() / "Desktop"
EXPORTS_DIR = APP_DIR / "exports"   # 统一存放导出文件
STATE_FILE = APP_DIR / "state.json"
DB_PATH = APP_DIR / "data.db"
SHEET_NAME = "费用报销清单"


def _resolve_app_dir():
    """PyInstaller 打包后资源在 _MEIPASS 或 .app/Contents/MacOS 下；
    数据文件（data.db/exports/invoices）需写入用户可写目录。
    返回 (resource_dir, data_dir)"""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # PyInstaller onefile：资源在 _MEIPASS
        resource_dir = Path(sys._MEIPASS)
        # 数据写入用户家目录
        data_dir = Path.home() / ".reimbursement"
        data_dir.mkdir(exist_ok=True)
        return resource_dir, data_dir
    return APP_DIR, APP_DIR


def get_paths():
    """获取所有路径（resource 用于只读资源，data 用于读写数据）"""
    resource_dir, data_dir = _resolve_app_dir()
    return {
        'template': resource_dir / "template.xlsx",
        'exports': data_dir / "exports",
        'state': data_dir / "state.json",
        'db': data_dir / "data.db",
    }


def get_user_name(db_path=None):
    """从 settings 表读取用户名，用于命名导出文件。"""
    try:
        import sqlite3
        p = db_path or get_paths()['db']
        db = sqlite3.connect(str(p))
        row = db.execute("SELECT value FROM settings WHERE key='user_name'").fetchone()
        db.close()
        return (row[0] if row else "") or "报销用户"
    except Exception:
        return "报销用户"

DATA_START_ROW = 3   # 模板第 3 行开始填数据
MAX_DATA_ROWS = 12   # 模板预留 12 行
MAX_ROWS = 500       # 安全上限

# ── Excel 列映射：state.json 字段 → 列号（1-based） ──
TEXT_COL_MAP = {
    "发生日期":     2,   # B
    "费用类别":     3,   # C
    "招待客户名称": 4,   # D
    "事前申请编号": 5,   # E
    "项目编号":     6,   # F
    "招待人同乘人": 7,   # G
    "人数":         8,   # H
    "金额":         9,   # I
    "备注":        12,   # L
}

PAY_IMG_COL  = 10   # J — 付款截图
RECV_IMG_COL = 11   # K — 水单截图

# 图片嵌入尺寸（像素）
IMG_CELL_WIDTH  = 80
IMG_CELL_HEIGHT = 110


# ═══════════════════════════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════════════════════════

def load_state(state_file=None) -> List[Dict]:
    """读取 state.json；无文件或无数据时抛 ValueError（不再 sys.exit）"""
    p = Path(state_file) if state_file else get_paths()['state']
    if not p.exists():
        raise ValueError("state.json 不存在，请先在系统中上传识别报销截图")
    with open(p, encoding="utf-8") as f:
        data = json.load(f)
    if not data:
        raise ValueError("当前批次无报销记录")
    return data


def materialize_image(value: str, temp_files: List[Path]) -> Optional[str]:
    """将 base64 data URI 或本地路径转为临时文件路径

    返回临时文件路径，调用方负责在函数结束后清理 temp_files 列表中的文件。
    如果 value 为空或无法解析，返回 None。
    """
    if not value:
        return None

    # 本地文件路径
    if os.path.exists(value):
        return value

    # base64 data URI
    if value.startswith("data:image/"):
        try:
            header, payload = value.split(",", 1)
            ext = "png"
            if "jpeg" in header or "jpg" in header:
                ext = "jpg"
            elif "webp" in header:
                ext = "webp"
            elif "gif" in header:
                ext = "gif"

            fd, tmp_name = tempfile.mkstemp(prefix="reimb_", suffix=f".{ext}")
            os.close(fd)
            tmp_path = Path(tmp_name)
            tmp_path.write_bytes(base64.b64decode(payload))
            temp_files.append(tmp_path)
            return str(tmp_path)
        except Exception as exc:
            print(f"⚠ 无法解析图片数据: {exc}")
            return None

    return None


def embed_image(ws, img_path: str, col: int, row: int):
    """向指定单元格嵌入图片（openpyxl 锚定方式）"""
    try:
        img = XlImage(img_path)
        # 缩放到单元格合适大小
        img.width = IMG_CELL_WIDTH
        img.height = IMG_CELL_HEIGHT
        cell_ref = f"{get_column_letter(col)}{row}"
        img.anchor = cell_ref
        ws.add_image(img)
        # 调整行高列宽以适配图片
        if ws.row_dimensions[row].height is None or ws.row_dimensions[row].height < IMG_CELL_HEIGHT * 0.75:
            ws.row_dimensions[row].height = IMG_CELL_HEIGHT * 0.75
    except Exception as exc:
        print(f"⚠ 嵌入图片失败 ({cell_ref}): {exc}")


# ═══════════════════════════════════════════════════════════
# 导出主逻辑
# ═══════════════════════════════════════════════════════════

def export(month=None, records_override=None, return_result=False, exports_dir=None, db_path=None):
    """
    导出报销明细 Excel。
    - month: 月份字符串（"4" / "4月" / None=当月）
    - records_override: 直接传入记录列表（优先于 state.json），serve.py 调用时用
    - return_result: True 时返回 dict（不 print/sys.exit），供 serve.py 调用
    - exports_dir: 自定义导出目录（默认用 get_paths()['exports']）
    - db_path: 自定义 db 路径（用于读取 user_name）

    返回（return_result=True 时）：
      {"ok": True, "path": "...", "filename": "...", "rows": N, "month": "4"}
      或 {"ok": False, "error": "..."}
    """
    def _fail(msg):
        if return_result:
            return {"ok": False, "error": msg}
        print(f"❌ {msg}")
        sys.exit(1)

    paths = get_paths()
    # 外部传入的目录优先（serve.py 调用时传入，保证路径一致）
    use_exports_dir = Path(exports_dir) if exports_dir else paths['exports']
    use_db_path = Path(db_path) if db_path else paths['db']
    # ── 读取记录 ──
    try:
        records = records_override if records_override is not None else load_state(paths['state'])
    except ValueError as e:
        return _fail(str(e))

    if len(records) > MAX_ROWS:
        records = records[:MAX_ROWS]

    # ── 确定月份和输出文件名 ──
    if not month:
        month = datetime.now().strftime("%m")
    month = str(month).replace("月", "").strip()
    if not month.isdigit():
        return _fail(f"无效月份: {month}")

    uname = get_user_name(use_db_path)
    output_name = f"{uname}{month}月费用报销明细.xlsx"
    use_exports_dir.mkdir(exist_ok=True)
    output_path = use_exports_dir / output_name

    # 文件已存在则加时间戳后缀
    if output_path.exists():
        ts = datetime.now().strftime("%H%M%S")
        output_name = f"{uname}{month}月费用报销明细_{ts}.xlsx"
        output_path = use_exports_dir / output_name

    # ── 复制模板 ──
    tpl = paths['template']
    if not tpl.exists():
        desktop_tpl = DESKTOP / "费用报销单模板.xlsx"
        if desktop_tpl.exists():
            shutil.copyfile(desktop_tpl, output_path)
        else:
            return _fail(f"模板文件不存在: {tpl}")
    else:
        shutil.copyfile(tpl, output_path)

    os.chmod(output_path, 0o644)

    # ── 打开工作簿 ──
    try:
        wb = load_workbook(output_path)
    except Exception as e:
        return _fail(f"打开模板失败: {e}")

    if SHEET_NAME not in wb.sheetnames:
        return _fail(f"模板中找不到 sheet「{SHEET_NAME}」，现有: {wb.sheetnames}")

    ws = wb[SHEET_NAME]

    # ── 清空模板预留数据（A 列序号保留，跳过合并单元格）──
    data_end = DATA_START_ROW + MAX_DATA_ROWS - 1
    merged_coords = set()
    for mr in ws.merged_cells.ranges:
        for row in range(mr.min_row, mr.max_row + 1):
            for col in range(mr.min_col, mr.max_col + 1):
                merged_coords.add((row, col))

    for r in range(DATA_START_ROW, data_end + 1):
        for c in range(2, 13):  # B~L 列，跳过序号列(A)
            if (r, c) not in merged_coords:
                ws.cell(row=r, column=c).value = None

    # ── 填入文本字段 ──
    temp_files: List[Path] = []
    actual_rows = min(len(records), MAX_DATA_ROWS)

    for i, record in enumerate(records[:actual_rows]):
        row_num = DATA_START_ROW + i
        for field, col_idx in TEXT_COL_MAP.items():
            value = record.get(field)
            if value is not None and value != "":
                ws.cell(row=row_num, column=col_idx).value = value

    # ── 嵌入图片（J 列付款截图、K 列水单截图） ──
    for i, record in enumerate(records[:actual_rows]):
        row_num = DATA_START_ROW + i
        pay_src = record.get("付款截图路径", "")
        pay_path = materialize_image(pay_src, temp_files)
        if pay_path:
            embed_image(ws, pay_path, PAY_IMG_COL, row_num)
        recv_src = record.get("水单截图路径", "")
        recv_path = materialize_image(recv_src, temp_files)
        if recv_path:
            embed_image(ws, recv_path, RECV_IMG_COL, row_num)

    # ── 更新合计行公式 ──
    sum_end_row = DATA_START_ROW + actual_rows - 1
    sum_formula = f"=SUM(I{DATA_START_ROW}:I{sum_end_row})"
    total_row = DATA_START_ROW + MAX_DATA_ROWS
    ws.cell(row=total_row, column=9).value = sum_formula
    count_formula = f"=SUM(H{DATA_START_ROW}:H{sum_end_row})"
    ws.cell(row=total_row, column=8).value = count_formula

    # ── 保存 ──
    try:
        wb.save(output_path)
    except Exception as e:
        return _fail(f"保存 Excel 失败: {e}")
    finally:
        wb.close()

    # ── 清理临时图片文件 ──
    for tmp in temp_files:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass

    result = {
        "ok": True,
        "path": str(output_path),
        "filename": output_path.name,
        "rows": actual_rows,
        "month": month,
        # 当输入记录数超过模板容量（12 行），告知调用方有截断
        "truncated": max(0, len(records) - MAX_DATA_ROWS),
        "total_input": len(records),
    }
    if return_result:
        return result
    # 命令行模式：打印
    print(f"✅ 已导出: {output_path}")
    print(f"   写入范围: B{DATA_START_ROW}:L{sum_end_row}")
    print(f"   记录数: {actual_rows} 条  |  月份: {month}月  |  合计公式: {sum_formula}")
    return result


def main():
    month = sys.argv[1] if len(sys.argv) > 1 else None
    export(month)


if __name__ == "__main__":
    main()
