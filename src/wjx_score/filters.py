"""问卷星筛选参数编码：ec (extra condition) / qc (question condition)

通过逆向页面 JS getQueryCond() 得到的编码规则。
"""

from datetime import date, timedelta
from urllib.parse import quote

EC_SEP = "\u250b"   # ┋ - ec 字段内分隔符
EC_JOIN = "\u3012"  # 〒 - 多个 ec 条件间分隔符
QC_SEP = "\u00a7"   # § - qc 字段内分隔符
QC_JOIN = "\u3012"  # 〒 - 多个 qc 条件间分隔符

# 动态日期码 (ec type=1, operator=4)
DYNAMIC_DATES = {
    "today":       "-1",  "今日": "-1",  "当日": "-1",
    "yesterday":   "-2",  "昨日": "-2",
    "this-week":   "-3",  "本周": "-3",
    "last-week":   "-4",  "上周": "-4",
    "this-month":  "-5",  "本月": "-5",
    "last-month":  "-6",  "上月": "-6",
    "this-year":   "-7",  "本年": "-7",
    "last-year":   "-8",  "去年": "-8",
    "last-7-days": "-9",  "近7天": "-9",
    "last-30-days":"-10", "近30天": "-10",
}

# 前端下拉选项（用于 Web UI）
DYNAMIC_DATE_OPTIONS = [
    ("-1",  "当日"),
    ("-2",  "昨日"),
    ("-3",  "本周"),
    ("-4",  "上周"),
    ("-5",  "本月"),
    ("-6",  "上月"),
    ("-7",  "本年"),
    ("-8",  "去年"),
    ("-9",  "最近7天"),
    ("-10", "最近30天"),
]

DEFAULT_ACTIVITY_ID = "331434168"
BASE_URL = "https://www.wjx.cn/wjx/activitystat/viewstatsummary.aspx"


def build_ec_condition(ec_type, operator, val1, val2=""):
    """构建单个 ec 条件字符串"""
    return f"{ec_type}{EC_SEP}{operator}{EC_SEP}{val1}{EC_SEP}{val2}"


def build_ec_date_dynamic(alias_or_code):
    """动态日期筛选 (当日/昨日/本周/...)"""
    code = DYNAMIC_DATES.get(alias_or_code, alias_or_code)
    return build_ec_condition(1, 4, code)


def build_ec_date_exact(date_str):
    """精确日期筛选"""
    return build_ec_condition(1, 0, date_str)


def build_ec_date_range(start, end):
    """日期区间筛选"""
    return build_ec_condition(1, 3, start, end)


def build_ec_date_before_days(days):
    """前天等需要计算实际日期的情况"""
    d = date.today() - timedelta(days=days)
    return build_ec_date_exact(d.strftime("%Y-%m-%d"))


def build_ec_score_range(min_val=None, max_val=None):
    """分数范围筛选 (ec type=5)"""
    if min_val is not None and max_val is not None:
        return build_ec_condition(5, 3, str(min_val), str(max_val))
    elif min_val is not None:
        return build_ec_condition(5, 1, str(min_val))
    elif max_val is not None:
        return build_ec_condition(5, 2, str(max_val))
    return ""


def build_qc_name(name, operator=2):
    """姓名筛选 (qc, questionValue=10000, 默认 operator=2 包含)"""
    return f"10000{QC_SEP}{name}{QC_SEP}{operator}"


def build_filter_url(activity_id, ec_conditions=None, qc_conditions=None):
    """组装带筛选条件的完整 URL"""
    ec_str = ""
    qc_str = ""
    if ec_conditions:
        ec_str = EC_JOIN.join(c for c in ec_conditions if c)
    if qc_conditions:
        qc_str = QC_JOIN.join(c for c in qc_conditions if c)
    return f"{BASE_URL}?activity={activity_id}&qc={quote(qc_str)}&ec={quote(ec_str)}"
