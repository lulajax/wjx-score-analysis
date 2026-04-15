#!/usr/bin/env python3
"""
展鹏公考 - 新学员登记
将学员信息追加到 新学员登记表.xlsx 的下一空行。

用法:
  python3 register_student.py --xlsx 新学员登记表.xlsx --data '{...}'
"""

import json, sys, io, os, argparse
from datetime import datetime

# Windows 终端中文输出
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

try:
    import openpyxl
except ImportError:
    print("正在安装 openpyxl...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
    import openpyxl


# 用户字段 → Excel 表头 的映射
FIELD_MAP = {
    "协议是否签署":       "协议是否签署",
    "学员姓名":           "学员姓名",
    "性别":               "性别",
    "年龄":               "年龄",
    "专业":               "专业",
    "学历和学位":         "学历和学位",
    "学校":               "学校",
    "从事何种行业":       "从事何种行业",
    "备考地区":           "备考地区",
    "目标考试":           "目标考试",
    "报考年份":           "报考年份",
    "参考经历":           "参考经历（0基础/备考几月）",
    "每日学习时长":       "每日学习时长",
    "政治面貌":           "政治面貌",
    "毕业时间":           "毕业时间",
    "邮寄地址":           "邮寄地址",
    "电话":               "电话",
    "报名课程产品名称":   "报名课程产品名称（全名称）",
    "实付金额":           "实付金额",
    "业绩归属":           "业绩归属",
    "报班时间":           "报班时间",
    "学员特殊情况备注":   "备注",
}

# 需要按日期处理的字段
DATE_FIELDS = {"报班时间", "毕业时间"}


def parse_date(text):
    """尝试将日期文本解析为 datetime 对象"""
    text = text.strip()
    for fmt in ("%Y年%m月%d日", "%Y-%m-%d", "%Y年%m月", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return text  # 无法解析则原样写入


def find_next_row(ws):
    """找到第一个空行（A 列为空）"""
    for r in range(2, ws.max_row + 2):
        if ws.cell(r, 1).value is None or str(ws.cell(r, 1).value).strip() in ("", "\n"):
            return r
    return ws.max_row + 1


def get_header_map(ws):
    """读取第 1 行，返回 {表头名: 列号}"""
    hmap = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v:
            hmap[v.strip()] = c
    return hmap


def main():
    ap = argparse.ArgumentParser(description="新学员登记")
    ap.add_argument("--xlsx", required=True, help="Excel 文件路径")
    ap.add_argument("--data", required=True, help="学员信息 JSON 字符串")
    args = ap.parse_args()

    data = json.loads(args.data)
    xlsx_path = os.path.abspath(args.xlsx)
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    header_map = get_header_map(ws)
    next_col = max(header_map.values()) + 1 if header_map else 1
    row = find_next_row(ws)

    filled = []
    for user_field, value in data.items():
        excel_header = FIELD_MAP.get(user_field, user_field)

        # 如果表头不存在，新增列
        if excel_header not in header_map:
            header_map[excel_header] = next_col
            ws.cell(1, next_col, excel_header)
            print(f"  新增列 {openpyxl.utils.get_column_letter(next_col)}: {excel_header}")
            next_col += 1

        col = header_map[excel_header]

        # 日期字段特殊处理
        if user_field in DATE_FIELDS:
            value = parse_date(str(value))
            if isinstance(value, datetime):
                ws.cell(row, col, value).number_format = "YYYY-MM-DD"
                filled.append(f"  {excel_header} ({openpyxl.utils.get_column_letter(col)}): {value.strftime('%Y-%m-%d')}")
                continue

        # 实付金额转数字
        if user_field == "实付金额":
            try:
                value = float(str(value).replace(",", "").replace("元", ""))
            except ValueError:
                pass

        ws.cell(row, col, value)
        filled.append(f"  {excel_header} ({openpyxl.utils.get_column_letter(col)}): {value}")

    wb.save(xlsx_path)
    print(f"\n已写入第 {row} 行，共 {len(filled)} 个字段:")
    for f in filled:
        print(f)
    print(f"\n文件已保存: {xlsx_path}")


if __name__ == "__main__":
    main()
