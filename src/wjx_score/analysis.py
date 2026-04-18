"""成绩分析：题目分组配置 + 正确率统计"""

import re

QUESTION_GROUPS = [
    ("changshi",   "常识判断", "常识判断",            list(range(1, 9)),      8),
    ("zhengzhi",   "政治理论", "政治理论",            list(range(9, 15)),     6),
    ("xuanci",     "言语理解", "选词填空",            [15, 16],               2),
    ("pianduan",   "言语理解", "片段阅读",            [17, 18, 19, 20, 23, 24], 6),
    ("yuju",       "言语理解", "语句表达",            [21, 22],               2),
    ("jisuan",     "数量关系", "计算问题",            [25, 26],               2),
    ("zuibuli",    "数量关系", "最不利原则",          [27],                   1),
    ("gongcheng",  "数量关系", "工程问题",            [28],                   1),
    ("shulun",     "数量关系", "数论基础",            [29],                   1),
    ("tuxing",     "判断推理", "图形推理",            list(range(30, 33)),    3),
    ("dingyi",     "判断推理", "定义判断",            [33, 34],               2),
    ("leibi",      "判断推理", "类比推理",            [35, 36],               2),
    ("luoji",      "判断推理", "逻辑判断",            list(range(37, 40)),    3),
    ("chazhao",    "资料分析", "查找比较",            [40],                   1),
    ("jiandan",    "资料分析", "简单计算",            [41],                   1),
    ("pingjun",    "资料分析", "判断平均数变化方向",  [42],                   1),
    ("bizhong",    "资料分析", "判断比重变化方向",    [43],                   1),
    ("zengzhang",  "资料分析", "增长量",              [44],                   1),
    ("shenlun",    "申论",     "申论",                [45],                   1),
]

MAJOR_SUBJECTS = ["常识判断", "政治理论", "言语理解", "数量关系", "判断推理", "资料分析", "申论"]
MAJOR_TOTALS = {"常识判断": 8, "政治理论": 6, "言语理解": 10, "数量关系": 5, "判断推理": 10, "资料分析": 5, "申论": 1}


def parse_qnum(title):
    m = re.match(r'\*?\s*(\d+)\.', title)
    return int(m.group(1)) if m else None


def analyze(questions):
    qr = {}
    for q in questions:
        n = parse_qnum(q.get('title', ''))
        if n is not None and q.get('is_correct') is not None:
            qr[n] = q['is_correct']
    groups = {}
    for key, major, sub, nums, total in QUESTION_GROUPS:
        correct = sum(1 for n in nums if qr.get(n, False))
        groups[key] = {
            "major": major, "sub": sub,
            "correct": correct, "total": total,
            "wrong": total - correct,
            "rate": round(correct / total * 100) if total else 0,
        }
    majors = {}
    for s in MAJOR_SUBJECTS:
        c = sum(g["correct"] for g in groups.values() if g["major"] == s)
        t = MAJOR_TOTALS[s]
        majors[s] = {"correct": c, "total": t, "rate": round(c / t * 100) if t else 0}
    return groups, majors
