"""新学员登记：从消息文本中识别学员字段，追加到 Excel 登记表"""

import os
import re
from datetime import datetime

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

# 字段显示顺序
FIELD_ORDER = list(FIELD_MAP.keys())

# 消息文本中可能出现的关键词 → 标准字段名
_FIELD_ALIASES = {}
for key in FIELD_MAP:
    _FIELD_ALIASES[key] = key

# --- 完整表头（含括号）---
_FIELD_ALIASES["参考经历（0基础/备考几月）"] = "参考经历"
_FIELD_ALIASES["参考经历(0基础/备考几月)"] = "参考经历"
_FIELD_ALIASES["参考经历（0基础）"] = "参考经历"
_FIELD_ALIASES["报名课程产品名称（全名称）"] = "报名课程产品名称"
_FIELD_ALIASES["报名课程产品名称(全名称)"] = "报名课程产品名称"

# --- 姓名 ---
_FIELD_ALIASES["姓名"] = "学员姓名"
_FIELD_ALIASES["名字"] = "学员姓名"
_FIELD_ALIASES["名称"] = "学员姓名"

# --- 电话 ---
_FIELD_ALIASES["手机"] = "电话"
_FIELD_ALIASES["手机号"] = "电话"
_FIELD_ALIASES["手机号码"] = "电话"
_FIELD_ALIASES["联系电话"] = "电话"
_FIELD_ALIASES["联系方式"] = "电话"
_FIELD_ALIASES["电话号码"] = "电话"
_FIELD_ALIASES["tel"] = "电话"

# --- 学历 ---
_FIELD_ALIASES["学历"] = "学历和学位"
_FIELD_ALIASES["学位"] = "学历和学位"
_FIELD_ALIASES["学历学位"] = "学历和学位"
_FIELD_ALIASES["最高学历"] = "学历和学位"

# --- 行业 ---
_FIELD_ALIASES["行业"] = "从事何种行业"
_FIELD_ALIASES["职业"] = "从事何种行业"
_FIELD_ALIASES["工作"] = "从事何种行业"
_FIELD_ALIASES["从事行业"] = "从事何种行业"
_FIELD_ALIASES["目前职业"] = "从事何种行业"
_FIELD_ALIASES["从事工作"] = "从事何种行业"

# --- 课程 ---
_FIELD_ALIASES["课程名称"] = "报名课程产品名称"
_FIELD_ALIASES["课程产品"] = "报名课程产品名称"
_FIELD_ALIASES["课程"] = "报名课程产品名称"
_FIELD_ALIASES["报名课程"] = "报名课程产品名称"
_FIELD_ALIASES["产品名称"] = "报名课程产品名称"

# --- 金额 ---
_FIELD_ALIASES["金额"] = "实付金额"
_FIELD_ALIASES["付款金额"] = "实付金额"
_FIELD_ALIASES["缴费金额"] = "实付金额"
_FIELD_ALIASES["费用"] = "实付金额"
_FIELD_ALIASES["学费"] = "实付金额"

# --- 地址 ---
_FIELD_ALIASES["地址"] = "邮寄地址"
_FIELD_ALIASES["收件地址"] = "邮寄地址"
_FIELD_ALIASES["快递地址"] = "邮寄地址"
_FIELD_ALIASES["收货地址"] = "邮寄地址"

# --- 协议 ---
_FIELD_ALIASES["协议"] = "协议是否签署"
_FIELD_ALIASES["签署协议"] = "协议是否签署"
_FIELD_ALIASES["是否签署协议"] = "协议是否签署"
_FIELD_ALIASES["合同"] = "协议是否签署"

# --- 备注 ---
_FIELD_ALIASES["备注"] = "学员特殊情况备注"
_FIELD_ALIASES["特殊情况备注"] = "学员特殊情况备注"
_FIELD_ALIASES["特殊情况"] = "学员特殊情况备注"
_FIELD_ALIASES["特殊备注"] = "学员特殊情况备注"

# --- 其他 ---
_FIELD_ALIASES["毕业院校"] = "学校"
_FIELD_ALIASES["院校"] = "学校"
_FIELD_ALIASES["大学"] = "学校"
_FIELD_ALIASES["学习时长"] = "每日学习时长"
_FIELD_ALIASES["学习时间"] = "每日学习时长"
_FIELD_ALIASES["每天学习时长"] = "每日学习时长"
_FIELD_ALIASES["每天学习时间"] = "每日学习时长"
_FIELD_ALIASES["考试地区"] = "备考地区"
_FIELD_ALIASES["地区"] = "备考地区"
_FIELD_ALIASES["考试"] = "目标考试"
_FIELD_ALIASES["目标"] = "目标考试"
_FIELD_ALIASES["考试经历"] = "参考经历"
_FIELD_ALIASES["备考经历"] = "参考经历"
_FIELD_ALIASES["基础"] = "参考经历"
_FIELD_ALIASES["业绩"] = "业绩归属"
_FIELD_ALIASES["归属"] = "业绩归属"
_FIELD_ALIASES["报班日期"] = "报班时间"
_FIELD_ALIASES["报名时间"] = "报班时间"
_FIELD_ALIASES["报名日期"] = "报班时间"
_FIELD_ALIASES["毕业日期"] = "毕业时间"

# 按长度降序排列，优先匹配长关键词
_SORTED_ALIASES = sorted(_FIELD_ALIASES.keys(), key=len, reverse=True)

# 行首编号/标点前缀：1. 或 18: 或 18： 或 ① 或 - 或 · 等
_LINE_PREFIX_RE = re.compile(r"^\s*(?:\d{1,2}\s*[.．、:：)\]）】]\s*|[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]\s*|[-·•]\s+)")

# 噪音行模式：聊天语气、打招呼等（跳过）
_NOISE_RE = re.compile(
    r"^(老师|你好|您好|嗨|hi|hello|麻烦|请|谢谢|感谢|辛苦|好的|收到|以下是|信息如下|学员信息|新学员|报名信息)",
    re.IGNORECASE,
)

# 值的模式识别正则
_PHONE_RE = re.compile(r"^1[3-9]\d{9}$")
_MONEY_RE = re.compile(r"^\d{3,6}(?:\.\d{1,2})?(?:元)?$")
_DATE_FULL_RE = re.compile(r"^\d{4}年\d{1,2}月\d{1,2}日$")
_GENDER_RE = re.compile(r"^[男女]$")
_AGE_RE = re.compile(r"^\d{1,2}岁?$")


def _parse_one(lines):
    """从一组行中解析一个学员的字段"""
    result = {}
    unmatched = []  # 未匹配的行，留给模式识别

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 跳过噪音行
        plain = _LINE_PREFIX_RE.sub("", line)
        if _NOISE_RE.match(plain):
            continue

        line = plain
        matched_field = None
        rest = line

        for alias in _SORTED_ALIASES:
            if line.startswith(alias):
                matched_field = _FIELD_ALIASES[alias]
                rest = line[len(alias):]
                break

        if matched_field:
            # 去掉字段名后的分隔符（：: = - 空格）
            rest = re.sub(r"^[\s:：=\-]+", "", rest).strip()
            result[matched_field] = rest
        else:
            unmatched.append(line)

    # 对未匹配的行做值模式识别
    for line in unmatched:
        # 可能是纯值（没有字段名），尝试按内容猜字段
        val = re.sub(r"^[\s:：=\-]+", "", line).strip()
        if not val:
            continue
        if "电话" not in result and _PHONE_RE.match(val):
            result["电话"] = val
        elif "性别" not in result and _GENDER_RE.match(val):
            result["性别"] = val
        elif "年龄" not in result and _AGE_RE.match(val):
            result["年龄"] = val
        elif "实付金额" not in result and _MONEY_RE.match(val.replace(",", "").replace("元", "")):
            result["实付金额"] = val

    return result


def parse_message(text):
    """从消息文本中解析学员字段。支持多人批量粘贴，返回列表 [{字段名: 值}, ...]"""
    lines = text.strip().splitlines()

    # 检测是否包含多人：寻找"学员姓名"出现的行位置
    name_positions = []
    for i, line in enumerate(lines):
        stripped = _LINE_PREFIX_RE.sub("", line.strip())
        for alias in ("学员姓名", "姓名"):
            if stripped.startswith(alias):
                name_positions.append(i)
                break

    if len(name_positions) <= 1:
        # 单人
        result = _parse_one(lines)
        return [result] if result else []

    # 多人：按"学员姓名"行分割
    groups = []
    for idx, pos in enumerate(name_positions):
        end = name_positions[idx + 1] if idx + 1 < len(name_positions) else len(lines)
        # 往前找属于当前人的行（从上一个人结束到当前姓名行）
        start = name_positions[idx - 1] if idx > 0 else 0
        if idx > 0:
            # 在 start..pos 之间找分界：编号重新从1开始、或空行
            block_start = pos
            for j in range(name_positions[idx - 1] + 1, pos):
                stripped = lines[j].strip()
                if not stripped:
                    block_start = j + 1
                    break
                m = re.match(r"^\s*1\s*[.．、:：]", stripped)
                if m:
                    block_start = j
                    break
            start = block_start
        groups.append(lines[start:end])

    results = []
    for group in groups:
        parsed = _parse_one(group)
        if parsed:
            results.append(parsed)
    return results

DATE_FIELDS = {"报班时间", "毕业时间"}


def _parse_date(text):
    text = text.strip()
    for fmt in ("%Y年%m月%d日", "%Y-%m-%d", "%Y年%m月", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return text


def _find_next_row(ws):
    for r in range(2, ws.max_row + 2):
        if ws.cell(r, 1).value is None or str(ws.cell(r, 1).value).strip() in ("", "\n"):
            return r
    return ws.max_row + 1


def _get_header_map(ws):
    hmap = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v:
            hmap[v.strip()] = c
    return hmap


def register(xlsx_path, data):
    """将学员数据追加到 Excel 文件，返回写入结果"""
    xlsx_path = os.path.abspath(xlsx_path)

    if not os.path.exists(xlsx_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        # 按 FIELD_MAP 的 value 创建表头
        for i, header in enumerate(FIELD_MAP.values(), 1):
            ws.cell(1, i, header)
        wb.save(xlsx_path)

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    header_map = _get_header_map(ws)
    next_col = max(header_map.values()) + 1 if header_map else 1
    row = _find_next_row(ws)

    filled = []
    for user_field, value in data.items():
        if not value:
            continue
        excel_header = FIELD_MAP.get(user_field, user_field)

        if excel_header not in header_map:
            header_map[excel_header] = next_col
            ws.cell(1, next_col, excel_header)
            next_col += 1

        col = header_map[excel_header]

        if user_field in DATE_FIELDS:
            value = _parse_date(str(value))
            if isinstance(value, datetime):
                ws.cell(row, col, value).number_format = "YYYY-MM-DD"
                filled.append({"field": excel_header, "value": value.strftime("%Y-%m-%d")})
                continue

        if user_field == "实付金额":
            try:
                value = float(str(value).replace(",", "").replace("元", ""))
            except ValueError:
                pass

        ws.cell(row, col, value)
        filled.append({"field": excel_header, "value": str(value)})

    wb.save(xlsx_path)
    return {"row": row, "count": len(filled), "fields": filled, "path": xlsx_path}
