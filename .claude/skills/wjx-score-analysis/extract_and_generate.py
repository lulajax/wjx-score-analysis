#!/usr/bin/env python3
"""
展鹏教育 - 问卷星学前测成绩提取 & HTML 成绩单生成
通过 CDP 从 wjx.cn 提取所有学员的逐题作答数据，
用 template.html 模板生成专业 HTML 成绩分析报告。

用法: python3 extract_and_generate.py <activity_id> [--output-dir ./成绩单]
如果 Chrome 未开启远程调试，脚本会自动拉起。
"""

import json, asyncio, csv, math, re, os, sys, time, argparse, urllib.request, subprocess, shutil
sys.stdout.reconfigure(line_buffering=True)
import websockets

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ═══════════════════════════════════════════
#  配置
# ═══════════════════════════════════════════
CDP_HOST = "localhost"
CDP_PORT = 9222
PAGE_SIZE = 100
SELECTED_CHAR = 59103

QUESTION_GROUPS = [
    ("changshi",   "常识判断", "常识判断",            list(range(1, 9)),      8),
    ("zhengzhi",   "政治理论", "政治理论",            list(range(9, 15)),     6),
    ("xuanci",     "言语理解", "选词填空",            [15, 16],               2),
    ("pianduan",   "言语理解", "片段阅读",            [17,18,19,20,23,24],    6),
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
MAJOR_TOTALS   = {"常识判断":8, "政治理论":6, "言语理解":10, "数量关系":5, "判断推理":10, "资料分析":5, "申论":1}

# ═══════════════════════════════════════════
#  CDP 通信
# ═══════════════════════════════════════════
_msg_counter = 1000

async def cdp_send(ws, method, params=None, timeout=15):
    global _msg_counter; _msg_counter += 1; mid = _msg_counter
    await ws.send(json.dumps({"id": mid, "method": method, **({"params": params} if params else {})}))
    dl = time.time() + timeout
    while time.time() < dl:
        resp = json.loads(await asyncio.wait_for(ws.recv(), timeout=max(dl - time.time(), 1)))
        if resp.get("id") == mid: return resp
    raise TimeoutError(f"CDP timeout: {method}")

async def cdp_eval(ws, expr, timeout=15):
    r = (await cdp_send(ws, "Runtime.evaluate",
         {"expression": expr, "returnByValue": True, "awaitPromise": True}, timeout)
        ).get("result",{}).get("result",{})
    return r.get("value","") if r.get("type") == "string" else r.get("value")

async def wait_ready(ws, sel, max_wait=20):
    for _ in range(max_wait):
        try:
            if await cdp_eval(ws, "document.readyState", 5) == "complete":
                if (await cdp_eval(ws, f"document.querySelectorAll('{sel}').length", 5) or 0) > 0:
                    return True
        except Exception: pass
        await asyncio.sleep(1)
    return False

# ═══════════════════════════════════════════
#  登录
# ═══════════════════════════════════════════
LOGIN_URL = "https://www.wjx.cn/Login.aspx?returnUrl=%2fnewwjx%2fmanage%2fmyquestionnaires.aspx%3frandomt%3d1774682900"

async def check_logged_in(ws):
    """检查是否已登录：访问管理页，看是否被重定向到登录页"""
    await cdp_send(ws, "Page.navigate", {"url": "https://www.wjx.cn/newwjx/manage/myquestionnaires.aspx"})
    await asyncio.sleep(4)
    for _ in range(10):
        try:
            if await cdp_eval(ws, "document.readyState", 5) == "complete": break
        except Exception: pass
        await asyncio.sleep(1)
    url = await cdp_eval(ws, "window.location.href") or ""
    return "login.aspx" not in url.lower()

async def try_login(ws, username, password):
    """尝试自动登录问卷星"""
    print("导航到登录页...")
    await cdp_send(ws, "Page.navigate", {"url": LOGIN_URL})
    await asyncio.sleep(3)
    for _ in range(10):
        try:
            if await cdp_eval(ws, "document.readyState", 5) == "complete": break
        except Exception: pass
        await asyncio.sleep(1)

    login_js = """(function(){
        var ui = document.querySelector('#txtAccount') ||
                 document.querySelector('input[name*="Account"]') ||
                 document.querySelector('input[placeholder*="手机"]') ||
                 document.querySelector('input[placeholder*="账号"]') ||
                 document.querySelector('input[type="text"]');
        if(!ui) return JSON.stringify({status:'error',msg:'找不到账号输入框'});
        var pi = document.querySelector('#txtPsw') ||
                 document.querySelector('input[name*="Psw"]') ||
                 document.querySelector('input[type="password"]');
        if(!pi) return JSON.stringify({status:'error',msg:'找不到密码输入框'});
        var s=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set;
        s.call(ui,'__USER__'); ui.dispatchEvent(new Event('input',{bubbles:true}));
        ui.dispatchEvent(new Event('change',{bubbles:true}));
        s.call(pi,'__PASS__'); pi.dispatchEvent(new Event('input',{bubbles:true}));
        pi.dispatchEvent(new Event('change',{bubbles:true}));
        var btn = document.querySelector('#btnLogin') ||
                  document.querySelector('input[type="submit"]') ||
                  document.querySelector('button[type="submit"]');
        if(!btn) return JSON.stringify({status:'error',msg:'找不到登录按钮'});
        btn.click(); return JSON.stringify({status:'clicked'});
    })()""".replace('__USER__', username).replace('__PASS__', password)

    result = await cdp_eval(ws, login_js)
    if result:
        info = json.loads(result)
        if info['status'] == 'clicked':
            print("已提交登录表单，等待响应...")
            await asyncio.sleep(5)
            url = await cdp_eval(ws, "window.location.href") or ""
            if "login.aspx" not in url.lower():
                return True
            print("仍在登录页，可能需要验证码或密码错误")
            return False
        else:
            print(f"登录失败: {info.get('msg','未知错误')}")
    return False

async def ensure_logged_in(ws, username, password):
    """确保已登录，失败则等待人工协助"""
    print("检查登录状态...")
    if await check_logged_in(ws):
        print("已登录问卷星")
        return True

    print("未登录，尝试自动登录...")
    if await try_login(ws, username, password):
        print("自动登录成功")
        return True

    print("\n" + "="*50)
    print("自动登录失败，可能需要验证码或人工操作")
    print("请在浏览器中手动完成登录，完成后按 Enter 继续...")
    print("="*50)
    loop = asyncio.get_running_loop()
    await loop.run_in_executor(None, lambda: input("按 Enter 继续..."))

    if await check_logged_in(ws):
        print("登录确认成功")
        return True
    print("仍未登录，退出。")
    return False

# ═══════════════════════════════════════════
#  汇总页
# ═══════════════════════════════════════════
TBL = "#ctl02_ContentPlaceHolder1_ViewStatSummary1_tbSummary"

async def set_page_size(ws, size=100):
    r = await cdp_eval(ws, f"""(function(){{
        var s=document.getElementById('ctl02_ContentPlaceHolder1_ViewStatSummary1_ddlPageCount');
        if(!s)return'no';s.value='{size}';s.dispatchEvent(new Event('change',{{bubbles:true}}));return'ok';}})()""")
    if r == 'ok': await asyncio.sleep(5); await wait_ready(ws, TBL)
    return r

async def page_info(ws):
    return json.loads(await cdp_eval(ws, """(function(){
        var t=document.querySelector('#ctl02_ContentPlaceHolder1_ViewStatSummary1_lbTotal'),
            p=document.querySelector('#ctl02_ContentPlaceHolder1_ViewStatSummary1_lbPage');
        return JSON.stringify({total:t?t.innerText.trim():'0',page:p?p.innerText.trim():'1/1'});})()"""))

async def next_page(ws):
    r = await cdp_eval(ws, """(function(){
        var b=document.getElementById('ctl02_ContentPlaceHolder1_ViewStatSummary1_btnNext');
        if(!b||b.disabled)return'disabled';b.click();return'clicked';})()""")
    if r == 'clicked': await asyncio.sleep(5); await wait_ready(ws, TBL)
    return r

async def summary_rows(ws):
    return json.loads(await cdp_eval(ws, """(function(){
        var t=document.getElementById('ctl02_ContentPlaceHolder1_ViewStatSummary1_tbSummary');
        if(!t)return'[]';var rs=[];
        for(var r=1;r<t.rows.length;r++){var w=t.rows[r];
        rs.push({joinid:w.getAttribute('jid'),seq:w.cells[4].innerText.trim(),
        name:w.cells[5].innerText.trim(),submit_time:w.cells[6].innerText.trim(),
        duration:w.cells[7].innerText.trim(),source:w.cells[8].innerText.trim(),
        ip_location:w.cells[10].innerText.trim(),total_score:w.cells[11].innerText.trim()});}
        return JSON.stringify(rs);})()"""))

# ═══════════════════════════════════════════
#  详情页
# ═══════════════════════════════════════════
DETAIL_JS = """(function(){var items=document.querySelectorAll('.data__items'),qs=[];
for(var i=0;i<items.length;i++){var it=items[i],q={};
q.topic=it.getAttribute('topic')||'';
var ti=it.querySelector('.data__tit_cjd');q.title=ti?ti.innerText.trim():'';
var sc=it.querySelector('.score-val-ques');q.max_score=sc?sc.innerText.trim().replace('分值','').replace('分',''):'';
var opts=it.querySelectorAll('.ulradiocheck > div'),ua=[];
opts.forEach(function(o){var ic=o.querySelector('i.icon'),sp=o.querySelector('span');
if(ic&&ic.textContent.charCodeAt(0)===59103)ua.push(sp?sp.innerText.trim():o.innerText.trim());});
q.user_answer=ua.join('; ');
if(!opts.length){var kd=it.querySelector('.data__key');if(kd){var cl=kd.cloneNode(true);
cl.querySelectorAll('.judge_ques_right,.judge_ques_false,.answer-ansys').forEach(function(j){j.remove();});
q.user_answer=cl.innerText.trim();}}
var jf=it.querySelector('.judge_ques_false'),jr=it.querySelector('.judge_ques_right:not(.judge_ques_false)');
if(jr){q.is_correct=true;var f=jr.querySelector('font');q.earned=f?f.innerText.trim().replace('+','').replace('分',''):'';}
else if(jf){q.is_correct=false;var f=jf.querySelector('font');q.earned=f?f.innerText.trim().replace('+','').replace('分',''):'';}
else{q.is_correct=null;q.earned='';}
var ad=it.querySelector('.answer-ansys');
if(ad){var m=ad.innerText.match(/正确答案[：:]([\\s\\S]*?)(?:答案解析|$)/);q.correct_answer=m?m[1].trim():'';}
else q.correct_answer='';qs.push(q);}return JSON.stringify(qs);})()"""

async def extract_detail(ws, activity_id, joinid):
    url = f"https://www.wjx.cn/Modules/Wjx/ViewCeShiJoinActivity.aspx?activity={activity_id}&joinid={joinid}&v=&pWidth=1"
    await cdp_send(ws, "Page.navigate", {"url": url})
    await asyncio.sleep(2)
    if not await wait_ready(ws, ".data__items", 15):
        print(f"  [WARN] 超时 joinid={joinid}"); return []
    r = await cdp_eval(ws, DETAIL_JS, 20)
    return json.loads(r) if r else []

# ═══════════════════════════════════════════
#  分析
# ═══════════════════════════════════════════
def parse_qnum(title):
    m = re.match(r'\*?\s*(\d+)\.', title)
    return int(m.group(1)) if m else None

def analyze(questions):
    qr = {}
    for q in questions:
        n = parse_qnum(q.get('title',''))
        if n is not None and q.get('is_correct') is not None: qr[n] = q['is_correct']
    groups = {}
    for key, major, sub, nums, total in QUESTION_GROUPS:
        correct = sum(1 for n in nums if qr.get(n, False))
        groups[key] = {"major":major,"sub":sub,"correct":correct,"total":total,
                       "wrong":total-correct,"rate":round(correct/total*100) if total else 0}
    majors = {}
    for s in MAJOR_SUBJECTS:
        c = sum(g["correct"] for g in groups.values() if g["major"]==s)
        t = MAJOR_TOTALS[s]
        majors[s] = {"correct":c,"total":t,"rate":round(c/t*100) if t else 0}
    return groups, majors

# ═══════════════════════════════════════════
#  HTML 渲染（模板驱动）
# ═══════════════════════════════════════════
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
    """生成雷达图：SVG 画图形，标签用 HTML 绝对定位（兼容 html2canvas）"""
    r = 55             # 雷达半径
    pad = 58           # 留给 HTML 标签的空间
    vb = (r + pad) * 2  # viewBox 尺寸 = 226
    cx, cy = vb // 2, vb // 2
    subs = [s for s in MAJOR_SUBJECTS if s != "申论"]
    n = len(subs)
    ang = [i*2*math.pi/n - math.pi/2 for i in range(n)]
    p = lambda a,rd: (cx+rd*math.cos(a), cy+rd*math.sin(a))

    # SVG 图形部分（不含文字）
    svg = [f'<svg viewBox="0 0 {vb} {vb}" xmlns="http://www.w3.org/2000/svg" style="display:block;width:100%;height:100%">']
    for lv in [.2,.4,.6,.8,1.]:
        pts=" ".join(f"{p(a,r*lv)[0]:.1f},{p(a,r*lv)[1]:.1f}" for a in ang)
        op=0.08 if lv<1 else 0.15
        svg.append(f'<polygon points="{pts}" fill="rgba(100,116,139,{op})" stroke="#94a3b8" stroke-width="0.5"/>')
    for a in ang:
        x2,y2=p(a,r); svg.append(f'<line x1="{cx}" y1="{cy}" x2="{x2:.1f}" y2="{y2:.1f}" stroke="#cbd5e1" stroke-width="0.5"/>')
    rates=[majors[s]["rate"]/100 for s in subs]
    pts=" ".join(f"{p(ang[i],r*max(rates[i],.05))[0]:.1f},{p(ang[i],r*max(rates[i],.05))[1]:.1f}" for i in range(n))
    svg.append(f'<polygon points="{pts}" fill="rgba(99,102,241,0.25)" stroke="#6366f1" stroke-width="2"/>')
    for i in range(n):
        x,y=p(ang[i],r*max(rates[i],.05)); svg.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="#6366f1"/>')
    svg.append('</svg>')

    # HTML 标签：将 SVG 坐标转为容器百分比
    label_dist = r + 20  # 标签到圆心的距离（SVG 坐标）
    labels = []
    for i,s in enumerate(subs):
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

    return (
        '<div class="radar-wrap">'
        + "\n".join(svg)
        + "\n".join(labels)
        + '</div>'
    )

def build_summary_table(majors):
    """生成 Excel 风格的科目汇总横表"""
    headers = "".join(f"<th>{s}</th>" for s in MAJOR_SUBJECTS)
    total_c = sum(majors[s]["correct"] for s in MAJOR_SUBJECTS)
    total_t = sum(majors[s]["total"] for s in MAJOR_SUBJECTS)
    total_r = round(total_c/total_t*100) if total_t else 0
    row_correct = "".join(f"<td>{majors[s]['correct']}</td>" for s in MAJOR_SUBJECTS)
    row_total   = "".join(f"<td>{majors[s]['total']}</td>" for s in MAJOR_SUBJECTS)
    row_rate    = "".join(f'<td class="rate-cell" style="color:{rate_color(majors[s]["rate"])}">{majors[s]["rate"]}%</td>' for s in MAJOR_SUBJECTS)
    return f'''<table class="summary-table">
<thead><tr><th></th>{headers}<th>总计</th></tr></thead>
<tbody>
<tr><td>答对</td>{row_correct}<td><b>{total_c}</b></td></tr>
<tr><td>总数</td>{row_total}<td>{total_t}</td></tr>
<tr><td>正确率</td>{row_rate}<td class="rate-cell" style="color:{rate_color(total_r)}">{total_r}%</td></tr>
</tbody></table>'''

def build_detail_tables(groups, majors):
    """生成 Excel 风格的细分横表，按大科目分组做 colspan"""
    # 将所有子类型按大科目分组，分两行表格展示
    all_subs = []
    for subj in MAJOR_SUBJECTS:
        sub_groups = [(k,g) for k,g in groups.items() if g["major"]==subj]
        if sub_groups:
            all_subs.append((subj, sub_groups))

    # 分成两行：前半 + 后半
    mid = (len(all_subs) + 1) // 2
    table_groups = [all_subs[:mid], all_subs[mid:]]
    tables = []
    for tg in table_groups:
        if not tg: continue
        # 第一行: 大科目名（colspan）
        h1 = ""
        # 第二行: 子类型名
        h2 = ""
        # 数据行
        row_correct = ""
        row_total   = ""
        row_rate    = ""
        for subj, subs in tg:
            span = len(subs)
            h1 += f'<th colspan="{span}">{subj}</th>'
            for _,g in subs:
                h2 += f"<th>{g['sub']}</th>"
                row_correct += f"<td>{g['correct']}</td>"
                row_total   += f"<td>{g['total']}</td>"
                c = rate_color(g["rate"])
                row_rate    += f'<td class="rate-cell" style="color:{c}">{g["rate"]}%</td>'
        tables.append(f'''<table class="detail-htable">
<thead><tr><th></th>{h1}</tr><tr><th>题目类型</th>{h2}</tr></thead>
<tbody>
<tr><td>答对</td>{row_correct}</tr>
<tr><td>总数</td>{row_total}</tr>
<tr><td>正确率</td>{row_rate}</tr>
</tbody></table>''')
    return "\n".join(tables)

def build_bar_chart(majors):
    """生成水平柱状图，紫色渐变色调"""
    subs = [s for s in MAJOR_SUBJECTS if s != "申论"]
    # 按正确率映射紫色深浅：高正确率深紫，低正确率浅紫
    bars = []
    for s in subs:
        m = majors[s]
        pct = m["rate"]
        # 紫色系渐变：低分浅紫，高分深紫
        if pct >= 80: bg = "linear-gradient(90deg,#6366f1,#4f46e5)"
        elif pct >= 60: bg = "linear-gradient(90deg,#818cf8,#6366f1)"
        elif pct >= 40: bg = "linear-gradient(90deg,#a5b4fc,#818cf8)"
        else: bg = "linear-gradient(90deg,#c7d2fe,#a5b4fc)"
        pct_label = f'<span class="hbar__pct">{pct}%</span>' if pct >= 25 else ''
        bars.append(f'''<div class="hbar">
  <span class="hbar__label">{s}</span>
  <div class="hbar__track"><div class="hbar__fill" style="width:{max(pct,2)}%;background:{bg}">{pct_label}</div></div>
  <span class="hbar__val">{m["correct"]}/{m["total"]}</span>
</div>''')
    return "\n".join(bars)

def generate_ai_analysis(majors, groups):
    """生成精简学情分析，控制在50~80字"""
    total_c = sum(m["correct"] for m in majors.values())
    total_t = sum(m["total"] for m in majors.values())
    rate = round(total_c / total_t * 100) if total_t else 0

    non_shenlun = {k: v for k, v in majors.items() if k != "申论"}
    strong = [s for s, m in non_shenlun.items() if m["rate"] >= 70]
    weak   = [s for s, m in non_shenlun.items() if m["rate"] < 50]

    if rate >= 80:   level = "表现优秀"
    elif rate >= 60: level = "基础良好"
    elif rate >= 40: level = "基础尚可"
    else:            level = "基础薄弱"

    parts = [f"正确率{rate}%，{level}。"]
    if strong:
        parts.append(f'{"、".join(strong)}为优势科目，保持即可。')
    if weak:
        parts.append(f'{"、".join(weak)}较薄弱，建议专项突破。')
    if not weak and not strong:
        parts.append("各科较均衡，建议全面巩固提升。")

    return "<p>" + "".join(parts) + "</p>"

def build_ai_section(ai_text):
    if not ai_text: return ""
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
        "name":           person["name"],
        "exam_name":      exam_name,
        "submit_time":    person["submit_time"],
        "total_score":    person["total_score"],
        "total_questions": str(total_questions),
        "total_correct":  str(total_correct),
        "overall_rate":   str(round(total_correct/total_questions*100)),
        "subject_cards":  build_summary_table(majors),
        "radar_svg":      build_svg_radar(majors),
        "bar_chart":      build_bar_chart(majors),
        "detail_tables":  build_detail_tables(groups, majors),
        "ai_section":     build_ai_section(ai_analysis),
    }
    html = template
    for key, val in data.items():
        html = html.replace("{{" + key + "}}", val)
    return html

# ═══════════════════════════════════════════
#  主流程
# ═══════════════════════════════════════════
DEFAULT_ACTIVITY_ID = "331434168"
# ec 参数格式: 1┋4┋{offset}┋  其中 offset: -1=今日, -2=昨日, -3=前天 ...
DAY_ALIASES = {"today": "-1", "今日": "-1", "yesterday": "-2", "昨日": "-2", "前天": "-3"}

def build_day_url(activity_id, day_offset):
    """根据 activity_id 和日期偏移构造带筛选的 URL"""
    sep = "\u250b"  # ┋
    ec = f"1{sep}4{sep}{day_offset}{sep}"
    from urllib.parse import quote
    return f"https://www.wjx.cn/wjx/activitystat/viewstatsummary.aspx?activity={activity_id}&&qc=&ec={quote(ec)}"

def parse_url(url_or_id):
    """从完整 URL 或纯 activity_id 提取信息，返回 (summary_url, activity_id)"""
    if url_or_id.startswith("http"):
        from urllib.parse import urlparse, parse_qs
        parsed = urlparse(url_or_id)
        qs = parse_qs(parsed.query)
        aid = qs.get("activity", [""])[0]
        return url_or_id, aid
    else:
        return f"https://www.wjx.cn/wjx/activitystat/viewstatsummary.aspx?activity={url_or_id}", url_or_id


async def main():
    ap = argparse.ArgumentParser(description="问卷星成绩提取 & HTML 成绩单生成")
    ap.add_argument("url", nargs="?", default=None, help="问卷星活动完整 URL 或 activity_id（可选，使用 --day 时可省略）")
    ap.add_argument("--day", default=None, help="日期快捷方式: today/今日, yesterday/昨日, 前天, 或直接传偏移量如 -1, -2, -3")
    ap.add_argument("--activity", default=DEFAULT_ACTIVITY_ID, help=f"activity_id（默认 {DEFAULT_ACTIVITY_ID}）")
    ap.add_argument("--output-dir", default="./成绩单")
    ap.add_argument("--cdp-port", type=int, default=9222)
    ap.add_argument("--limit", type=int, default=0, help="只提取前 N 人 (0=全部)")
    ap.add_argument("--exam-name", default="学前测")
    ap.add_argument("--template", default=os.path.join(SCRIPT_DIR, "template.html"), help="HTML 模板路径")
    ap.add_argument("--username", default="15930272873", help="问卷星登录账号")
    ap.add_argument("--password", default="20250715zp", help="问卷星登录密码")
    args = ap.parse_args()
    global CDP_PORT; CDP_PORT = args.cdp_port

    if args.url:
        summary_url, activity_id = parse_url(args.url)
        day_offset = -1  # URL 模式默认当日
    elif args.day:
        offset = DAY_ALIASES.get(args.day, args.day)  # 支持别名或直接传 -1/-2
        summary_url = build_day_url(args.activity, offset)
        activity_id = args.activity
        day_offset = int(offset)
    else:
        # 默认今日
        summary_url = build_day_url(args.activity, "-1")
        activity_id = args.activity
        day_offset = -1

    # 根据日期偏移计算目标日期，用于文件名
    from datetime import datetime, timedelta
    target_date = (datetime.now() + timedelta(days=day_offset + 1)).strftime("%Y-%m-%d")

    print(f"汇总页: {summary_url}")
    print(f"Activity ID: {activity_id}")
    print(f"目标日期: {target_date}")

    # 读取模板
    with open(args.template, encoding="utf-8") as f:
        template = f.read()
    print(f"模板: {args.template}")

    # 尝试连接 Chrome CDP，如果未启动则自动拉起
    tabs = None
    for attempt in range(2):
        try:
            tabs = json.loads(urllib.request.urlopen(f"http://{CDP_HOST}:{CDP_PORT}/json", timeout=3).read())
            break
        except Exception:
            if attempt == 0:
                print("Chrome 远程调试未就绪，正在自动启动...")
                chrome_bin = shutil.which("google-chrome") or shutil.which("google-chrome-stable") or shutil.which("chromium-browser") or shutil.which("chromium") or "google-chrome"
                subprocess.Popen(
                    [chrome_bin, f"--remote-debugging-port={CDP_PORT}", "--no-first-run", "--no-default-browser-check"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                )
                # 等待 Chrome 启动完成
                for i in range(15):
                    time.sleep(1)
                    try:
                        tabs = json.loads(urllib.request.urlopen(f"http://{CDP_HOST}:{CDP_PORT}/json", timeout=2).read())
                        break
                    except Exception:
                        pass
                if tabs:
                    print("Chrome 已自动启动并就绪")
                    break
                else:
                    print("ERROR: 无法启动 Chrome，请手动运行: google-chrome --remote-debugging-port=9222")
                    sys.exit(1)

    pts = [t for t in tabs if t.get("type")=="page"]
    if not pts: print("ERROR: 需要至少 1 个浏览器标签页"); sys.exit(1)

    ws = await websockets.connect(pts[0]["webSocketDebuggerUrl"], max_size=10*1024*1024)
    try:
        await cdp_send(ws, "Page.enable")

        # ── 登录检查 ──
        if not await ensure_logged_in(ws, args.username, args.password):
            return

        # ── 阶段1: 翻页收集所有人的汇总信息 ──
        print("加载汇总页...")
        await cdp_send(ws, "Page.navigate", {"url": summary_url})
        await asyncio.sleep(3); await wait_ready(ws, TBL)

        print(f"设置每页 {PAGE_SIZE} 条...")
        await set_page_size(ws, PAGE_SIZE)

        info = None
        data_ready = False
        for i in range(10):
            info = await page_info(ws)
            if '/' in info.get('page','') and info.get('total','0')!='0':
                data_ready = True; break
            print(f"  等待... ({i+1})"); await asyncio.sleep(2); await wait_ready(ws, TBL)

        if not data_ready or not info or info.get('total','0') == '0':
            print("该日期范围内没有数据，请检查筛选条件。")
            return

        total = int(info['total'])
        page_str = info.get('page', '1/1')
        _cp, tp = map(int, page_str.split('/')) if '/' in page_str else (1, 1)
        print(f"共 {total} 条, {tp} 页")

        all_people = []
        for pn in range(1, tp+1):
            if pn > 1:
                if await next_page(ws) == 'disabled': break
            rows = await summary_rows(ws)
            all_people.extend(rows)
            print(f"  第 {pn}/{tp} 页: +{len(rows)} (累计 {len(all_people)})")

        if args.limit: all_people = all_people[:args.limit]
        print(f"\n汇总完成，共 {len(all_people)} 人")

        # ── 阶段2: 逐个导航到详情页提取数据（复用同一个标签页）──
        print(f"开始提取详情...\n")
        os.makedirs(args.output_dir, exist_ok=True)

        all_results = []
        for idx, p in enumerate(all_people):
            print(f"[{idx+1}/{len(all_people)}] {p['name']} (总分={p['total_score']})", end=" ", flush=True)
            qs = await extract_detail(ws, activity_id, p['joinid'])
            groups, majors = analyze(qs)
            tc = sum(g["correct"] for g in groups.values())
            print(f"答对 {tc}/45")

            all_results.append({"person":p, "questions":qs, "groups":groups, "majors":majors})

            ai_text = generate_ai_analysis(majors, groups)
            html = render_html(template, p, groups, majors, ai_analysis=ai_text, exam_name=args.exam_name)
            with open(os.path.join(args.output_dir, f"{target_date}-{args.exam_name}成绩单-{p['name']}.html"), "w", encoding="utf-8") as f:
                f.write(html)

        # JSON 备份
        jp = os.path.join(args.output_dir, f"{target_date}-全部成绩详情.json")
        with open(jp,"w",encoding="utf-8") as f: json.dump(all_results,f,ensure_ascii=False,indent=2)

        # CSV 汇总
        cp_ = os.path.join(args.output_dir, f"{target_date}-成绩汇总.csv")
        with open(cp_,"w",newline="",encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["序号","姓名","提交时间","用时","总分"]+MAJOR_SUBJECTS)
            for e in all_results:
                ep = e["person"]
                w.writerow([ep['seq'],ep['name'],ep['submit_time'],ep['duration'],ep['total_score']]
                           +[f"{e['majors'][s]['correct']}/{e['majors'][s]['total']}" for s in MAJOR_SUBJECTS])

        print(f"\n{'='*50}\n完成! {len(all_results)} 份成绩单 → {args.output_dir}/\n{'='*50}")
    finally:
        await ws.close()

if __name__ == "__main__":
    asyncio.run(main())
