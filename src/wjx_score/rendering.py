"""HTML 成绩单渲染：模板填充、图表生成"""

import math
from pathlib import Path

from .analysis import MAJOR_SUBJECTS

_TEMPLATE_PATH = Path(__file__).parent / "template.html"


def default_template():
    return _TEMPLATE_PATH.read_text(encoding="utf-8")


def rate_color(rate):
    if rate >= 80: return "#10b981"
    if rate >= 60: return "#6366f1"
    if rate >= 40: return "#f59e0b"
    return "#ef4444"


def rate_label(rate):
    if rate >= 80: return "优秀"
    if rate >= 60: return "良好"
    if rate >= 40: return "中等"
    return "薄弱"


def build_svg_radar(majors):
    r = 55
    pad = 58
    vb = (r + pad) * 2
    cx, cy = vb // 2, vb // 2
    subs = [s for s in MAJOR_SUBJECTS if s != "申论"]
    n = len(subs)
    ang = [i * 2 * math.pi / n - math.pi / 2 for i in range(n)]
    p = lambda a, rd: (cx + rd * math.cos(a), cy + rd * math.sin(a))

    svg = [f'<svg viewBox="0 0 {vb} {vb}" xmlns="http://www.w3.org/2000/svg" style="display:block;width:100%;height:100%">']
    for lv in [.2, .4, .6, .8, 1.]:
        pts = " ".join(f"{p(a, r * lv)[0]:.1f},{p(a, r * lv)[1]:.1f}" for a in ang)
        op = 0.08 if lv < 1 else 0.15
        svg.append(f'<polygon points="{pts}" fill="rgba(100,116,139,{op})" stroke="#94a3b8" stroke-width="0.5"/>')
    for a in ang:
        x2, y2 = p(a, r)
        svg.append(f'<line x1="{cx}" y1="{cy}" x2="{x2:.1f}" y2="{y2:.1f}" stroke="#cbd5e1" stroke-width="0.5"/>')
    rates = [majors[s]["rate"] / 100 for s in subs]
    pts = " ".join(f"{p(ang[i], r * max(rates[i], .05))[0]:.1f},{p(ang[i], r * max(rates[i], .05))[1]:.1f}" for i in range(n))
    svg.append(f'<polygon points="{pts}" fill="rgba(99,102,241,0.25)" stroke="#6366f1" stroke-width="2"/>')
    for i in range(n):
        x, y = p(ang[i], r * max(rates[i], .05))
        svg.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="#6366f1"/>')
    svg.append('</svg>')

    label_dist = r + 20
    labels = []
    for i, s in enumerate(subs):
        a = ang[i]
        lx = cx + label_dist * math.cos(a)
        ly = cy + label_dist * math.sin(a)
        lx_pct = lx / vb * 100
        ly_pct = ly / vb * 100
        labels.append(
            f'<div class="radar-label" style="left:{lx_pct:.1f}%;top:{ly_pct:.1f}%">'
            f'<span class="radar-label__name">{s}</span>'
            f'<span class="radar-label__rate">{majors[s]["rate"]}%</span>'
            f'</div>'
        )

    return '<div class="radar-wrap">' + "\n".join(svg) + "\n".join(labels) + '</div>'


def build_summary_table(majors):
    headers = "".join(f"<th>{s}</th>" for s in MAJOR_SUBJECTS)
    total_c = sum(majors[s]["correct"] for s in MAJOR_SUBJECTS)
    total_t = sum(majors[s]["total"] for s in MAJOR_SUBJECTS)
    total_r = round(total_c / total_t * 100) if total_t else 0
    row_correct = "".join(f"<td>{majors[s]['correct']}</td>" for s in MAJOR_SUBJECTS)
    row_total = "".join(f"<td>{majors[s]['total']}</td>" for s in MAJOR_SUBJECTS)
    row_rate = "".join(f'<td class="rate-cell" style="color:{rate_color(majors[s]["rate"])}">{majors[s]["rate"]}%</td>' for s in MAJOR_SUBJECTS)
    return f'''<table class="summary-table">
<thead><tr><th></th>{headers}<th>总计</th></tr></thead>
<tbody>
<tr><td>答对</td>{row_correct}<td><b>{total_c}</b></td></tr>
<tr><td>总数</td>{row_total}<td>{total_t}</td></tr>
<tr><td>正确率</td>{row_rate}<td class="rate-cell" style="color:{rate_color(total_r)}">{total_r}%</td></tr>
</tbody></table>'''


def build_detail_tables(groups, majors):
    all_subs = []
    for subj in MAJOR_SUBJECTS:
        sub_groups = [(k, g) for k, g in groups.items() if g["major"] == subj]
        if sub_groups:
            all_subs.append((subj, sub_groups))

    mid = (len(all_subs) + 1) // 2
    table_groups = [all_subs[:mid], all_subs[mid:]]
    tables = []
    for tg in table_groups:
        if not tg:
            continue
        h1 = ""
        h2 = ""
        row_correct = ""
        row_total = ""
        row_rate = ""
        for subj, subs in tg:
            span = len(subs)
            h1 += f'<th colspan="{span}">{subj}</th>'
            for _, g in subs:
                h2 += f"<th>{g['sub']}</th>"
                row_correct += f"<td>{g['correct']}</td>"
                row_total += f"<td>{g['total']}</td>"
                c = rate_color(g["rate"])
                row_rate += f'<td class="rate-cell" style="color:{c}">{g["rate"]}%</td>'
        tables.append(f'''<table class="detail-htable">
<thead><tr><th></th>{h1}</tr><tr><th>题目类型</th>{h2}</tr></thead>
<tbody>
<tr><td>答对</td>{row_correct}</tr>
<tr><td>总数</td>{row_total}</tr>
<tr><td>正确率</td>{row_rate}</tr>
</tbody></table>''')
    return "\n".join(tables)


def build_bar_chart(majors):
    subs = [s for s in MAJOR_SUBJECTS if s != "申论"]
    bars = []
    for s in subs:
        m = majors[s]
        pct = m["rate"]
        if pct >= 80: bg = "linear-gradient(90deg,#6366f1,#4f46e5)"
        elif pct >= 60: bg = "linear-gradient(90deg,#818cf8,#6366f1)"
        elif pct >= 40: bg = "linear-gradient(90deg,#a5b4fc,#818cf8)"
        else: bg = "linear-gradient(90deg,#c7d2fe,#a5b4fc)"
        pct_label = f'<span class="hbar__pct">{pct}%</span>' if pct >= 25 else ''
        bars.append(f'''<div class="hbar">
  <span class="hbar__label">{s}</span>
  <div class="hbar__track"><div class="hbar__fill" style="width:{max(pct, 2)}%;background:{bg}">{pct_label}</div></div>
  <span class="hbar__val">{m["correct"]}/{m["total"]}</span>
</div>''')
    return "\n".join(bars)


def generate_ai_analysis(majors, groups):
    total_c = sum(m["correct"] for m in majors.values())
    total_t = sum(m["total"] for m in majors.values())
    rate = round(total_c / total_t * 100) if total_t else 0

    non_shenlun = {k: v for k, v in majors.items() if k != "申论"}
    strong = [s for s, m in non_shenlun.items() if m["rate"] >= 70]
    weak = [s for s, m in non_shenlun.items() if m["rate"] < 50]

    if rate >= 80: level = "表现优秀"
    elif rate >= 60: level = "基础良好"
    elif rate >= 40: level = "基础尚可"
    else: level = "基础薄弱"

    parts = [f"正确率{rate}%，{level}。"]
    if strong:
        parts.append(f'{"、".join(strong)}为优势科目，保持即可。')
    if weak:
        parts.append(f'{"、".join(weak)}较薄弱，建议专项突破。')
    if not weak and not strong:
        parts.append("各科较均衡，建议全面巩固提升。")

    return "<p>" + "".join(parts) + "</p>"


def build_ai_section(ai_text):
    if not ai_text:
        return ""
    return f'''<div class="ai-section">
  <div class="ai-section__header">
    <span class="ai-section__icon">&#x1f4a1;</span>
    <span class="ai-section__title">学情分析</span>
  </div>
  <div class="ai-section__body">{ai_text}</div>
</div>'''


def render_html(template, person, groups, majors, ai_analysis="", exam_name="学前测"):
    total_correct = sum(g["correct"] for g in groups.values())
    total_questions = 45
    data = {
        "name": person["name"],
        "exam_name": exam_name,
        "submit_time": person["submit_time"],
        "total_score": person["total_score"],
        "total_questions": str(total_questions),
        "total_correct": str(total_correct),
        "overall_rate": str(round(total_correct / total_questions * 100)),
        "subject_cards": build_summary_table(majors),
        "radar_svg": build_svg_radar(majors),
        "bar_chart": build_bar_chart(majors),
        "detail_tables": build_detail_tables(groups, majors),
        "ai_section": build_ai_section(ai_analysis),
    }
    html = template
    for key, val in data.items():
        html = html.replace("{{" + key + "}}", val)
    return html
