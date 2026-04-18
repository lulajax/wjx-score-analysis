"""问卷星数据提取：汇总页翻页 + 详情页逐题数据"""

import json
import asyncio

from .cdp import cdp_send, cdp_eval, wait_ready

PAGE_SIZE = 100
TBL = "#ctl02_ContentPlaceHolder1_ViewStatSummary1_tbSummary"

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


async def check_logged_in(ws):
    """检查是否已登录：访问管理页，检测未登录的各种表现形式"""
    await cdp_send(ws, "Page.navigate", {"url": "https://www.wjx.cn/newwjx/manage/myquestionnaires.aspx"})
    for _ in range(15):
        await asyncio.sleep(1)
        try:
            if await cdp_eval(ws, "document.readyState", 5) == "complete":
                break
        except Exception:
            pass

    url = (await cdp_eval(ws, "window.location.href") or "").lower()
    if "login" in url:
        return False

    # wjx.cn 未登录时不重定向，而是弹 layer.alert('您还未登录')，
    # 页面 title 变为 "系统提示"，body 内含 Login.aspx 跳转脚本
    title = (await cdp_eval(ws, "document.title") or "").strip()
    if title == "系统提示":
        return False

    body = await cdp_eval(ws, "document.body?document.body.innerHTML.substring(0,800):''") or ""
    if "未登录" in body or "Login.aspx" in body:
        return False

    return True


async def login(ws, username, password):
    """自动登录问卷星"""
    await cdp_send(ws, "Page.navigate", {"url": "https://www.wjx.cn/login.aspx"})
    await asyncio.sleep(3)
    await wait_ready(ws, "#LoginButton")

    # 填入用户名和密码
    await cdp_eval(ws, f"""(function(){{
        var u=document.getElementById('UserName');
        var p=document.getElementById('Password');
        u.value='';p.value='';
        u.focus();u.value={json.dumps(username)};u.dispatchEvent(new Event('input',{{bubbles:true}}));
        p.focus();p.value={json.dumps(password)};p.dispatchEvent(new Event('input',{{bubbles:true}}));
        var cb=document.getElementById('checkxiexi');
        if(cb&&!cb.checked)cb.click();
    }})()""")
    await asyncio.sleep(0.5)

    # 点击登录
    await cdp_eval(ws, "document.getElementById('LoginButton').click()")
    await asyncio.sleep(4)

    # 检查结果
    for _ in range(10):
        try:
            if await cdp_eval(ws, "document.readyState", 5) == "complete":
                break
        except Exception:
            pass
        await asyncio.sleep(1)

    url = await cdp_eval(ws, "window.location.href") or ""
    if "login.aspx" in url.lower():
        # 获取错误提示
        err = await cdp_eval(ws, """(function(){
            var e=document.querySelector('.vc_error, .err-tip, .error-msg, #lblErr');
            return e?e.innerText.trim():'';
        })()""")
        return {"success": False, "error": err or "登录失败，请检查用户名和密码"}
    return {"success": True}


async def set_page_size(ws, size=100):
    r = await cdp_eval(ws, f"""(function(){{
        var s=document.getElementById('ctl02_ContentPlaceHolder1_ViewStatSummary1_ddlPageCount');
        if(!s)return'no';s.value='{size}';s.dispatchEvent(new Event('change',{{bubbles:true}}));return'ok';}})()""")
    if r == 'ok':
        await asyncio.sleep(5)
        await wait_ready(ws, TBL)
    return r


async def page_info(ws):
    return json.loads(await cdp_eval(ws, """(function(){
        var t=document.querySelector('#ctl02_ContentPlaceHolder1_ViewStatSummary1_lbTotal'),
            p=document.querySelector('#ctl02_ContentPlaceHolder1_ViewStatSummary1_lbPage');
        return JSON.stringify({total:t?t.innerText.trim():'0',page:p?p.innerText.trim():'1/1'});})()"""))


async def next_page(ws):
    r = await cdp_eval(ws, """(function(){
        var b=document.getElementById('ctl02_ContentPlaceHolder1_ViewStatSummary1_btnNext');
        if(!b||b.hasAttribute('disabled'))return'disabled';b.click();return'clicked';})()""")
    if r == 'clicked':
        await asyncio.sleep(5)
        await wait_ready(ws, TBL)
    return r


async def prev_page(ws):
    r = await cdp_eval(ws, """(function(){
        var b=document.getElementById('ctl02_ContentPlaceHolder1_ViewStatSummary1_btnPre');
        if(!b||b.hasAttribute('disabled'))return'disabled';b.click();return'clicked';})()""")
    if r == 'clicked':
        await asyncio.sleep(5)
        await wait_ready(ws, TBL)
    return r


async def current_page_data(ws):
    """获取当前页的学员数据和分页信息"""
    info = await page_info(ws)
    rows = await summary_rows(ws)
    total = int(info.get('total', '0'))
    page_str = info.get('page', '1/1')
    if '/' in page_str:
        cur, total_pages = map(int, page_str.split('/'))
    else:
        cur, total_pages = 1, 1
    return {
        "students": rows,
        "total": total,
        "page": cur,
        "total_pages": total_pages,
    }


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


async def extract_detail(ws, activity_id, joinid):
    url = f"https://www.wjx.cn/Modules/Wjx/ViewCeShiJoinActivity.aspx?activity={activity_id}&joinid={joinid}&v=&pWidth=1"
    await cdp_send(ws, "Page.navigate", {"url": url})
    await asyncio.sleep(2)
    if not await wait_ready(ws, ".data__items", 15):
        print(f"  [WARN] 超时 joinid={joinid}")
        return []
    r = await cdp_eval(ws, DETAIL_JS, 20)
    return json.loads(r) if r else []


_EMPTY = {"students": [], "total": 0, "page": 1, "total_pages": 1}

# 无数据时页面可能显示的提示元素
_NO_DATA_CHECK = """(function(){
    var body = document.body ? document.body.innerText : '';
    if (/暂无数据|没有.*数据|no.*data|无记录/.test(body)) return true;
    var tbl = document.querySelector('%s');
    if (tbl && tbl.rows.length <= 1) return true;
    return false;
})()""" % TBL


async def load_summary_page(ws, summary_url):
    """导航到汇总页并设置每页条数，返回首页数据"""
    await cdp_send(ws, "Page.navigate", {"url": summary_url})
    await asyncio.sleep(3)

    # 等页面加载完成（不要求表格一定存在）
    for _ in range(15):
        try:
            if await cdp_eval(ws, "document.readyState", 5) == "complete":
                break
        except Exception:
            pass
        await asyncio.sleep(1)

    # 快速检查：表格不存在或无数据行 → 直接返回空
    has_tbl = await cdp_eval(ws, f"document.querySelectorAll('{TBL}').length", 5)
    if not has_tbl:
        no_data = await cdp_eval(ws, _NO_DATA_CHECK, 5)
        if no_data:
            return _EMPTY
        # 表格还没出来，再等一轮
        if not await wait_ready(ws, TBL, 10):
            return _EMPTY

    # 检查 total，为 0 直接返回
    info = await page_info(ws)
    if info.get('total', '0') == '0':
        return _EMPTY

    await set_page_size(ws, PAGE_SIZE)
    return await current_page_data(ws)


async def fetch_all_students(ws, summary_url):
    """加载汇总页并翻页提取所有学员信息"""
    first = await load_summary_page(ws, summary_url)
    if first["total"] == 0:
        return []

    all_people = list(first["students"])
    for _ in range(first["total_pages"] - 1):
        if await next_page(ws) == 'disabled':
            break
        rows = await summary_rows(ws)
        all_people.extend(rows)
    return all_people
