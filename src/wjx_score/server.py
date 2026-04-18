"""Flask Web 服务：提供 API 接口和前端页面"""

import asyncio
import csv
import json
import os
import queue
import threading
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory, Response

from . import cdp, extraction, filters, analysis, rendering

STATIC_DIR = Path(__file__).parent / "static"

app = Flask(__name__, static_folder=None)

# 全局状态
_state = {
    "ws": None,
    "cdp_host": "localhost",
    "cdp_port": 9222,
    "activity_id": filters.DEFAULT_ACTIVITY_ID,
    "output_dir": "./成绩单",
    "exam_name": "学前测",
    "template": None,
    "current_url": None,       # 当前加载的汇总页 URL
    "cached_students": [],     # 当前页学员数据缓存
    "cached_page": "",
    "cached_total": 0,
    "reports": {},             # joinid -> {name, html} 生成的报告缓存
}

# 生成进度队列
_progress_queue = queue.Queue()

# 持久 event loop（在后台线程中运行，所有 async CDP 操作都提交到这里）
_loop = asyncio.new_event_loop()
_loop_thread = threading.Thread(target=_loop.run_forever, daemon=True)
_loop_thread.start()


def _run_async(coro):
    """将 async 协程提交到持久 event loop 并等待结果"""
    future = asyncio.run_coroutine_threadsafe(coro, _loop)
    return future.result(timeout=120)


async def _ensure_ws():
    """确保 CDP websocket 连接可用，断连时自动重连"""
    ws = _state["ws"]
    if ws is not None:
        try:
            await cdp.cdp_send(ws, "Runtime.evaluate",
                               {"expression": "1"}, timeout=5)
        except Exception:
            _state["ws"] = None
    if _state["ws"] is None:
        _state["ws"] = await cdp.connect(_state["cdp_host"], _state["cdp_port"])
    return _state["ws"]


@app.route("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.route("/api/check-login")
def check_login():
    try:
        ws = _run_async(_ensure_ws())
        logged_in = _run_async(extraction.check_logged_in(ws))
        return jsonify({"logged_in": logged_in})
    except Exception as e:
        return jsonify({"logged_in": False, "error": str(e)}), 500


@app.route("/api/login", methods=["POST"])
def do_login():
    data = request.get_json()
    username = data.get("username", "")
    password = data.get("password", "")
    if not username or not password:
        return jsonify({"success": False, "error": "请输入用户名和密码"}), 400
    try:
        ws = _run_async(_ensure_ws())
        result = _run_async(extraction.login(ws, username, password))
        return jsonify(result)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/students")
def get_students():
    """查询学员列表，支持筛选和分页"""
    try:
        ws = _run_async(_ensure_ws())
    except Exception as e:
        return jsonify({"error": f"CDP 连接失败: {e}"}), 500

    # 构建筛选条件
    ec_conditions = []
    qc_conditions = []

    day = request.args.get("day", "")
    date_start = request.args.get("date_start", "")
    date_end = request.args.get("date_end", "")
    name = request.args.get("name", "")
    name_op = int(request.args.get("name_op", "2"))  # 默认"包含"

    # 日期筛选
    if day:
        ec_conditions.append(filters.build_ec_date_dynamic(day))
    elif date_start and date_end:
        ec_conditions.append(filters.build_ec_date_range(date_start, date_end))
    elif date_start:
        ec_conditions.append(filters.build_ec_date_exact(date_start))

    # 姓名筛选
    if name:
        qc_conditions.append(filters.build_qc_name(name, operator=name_op))

    activity_id = request.args.get("activity", _state["activity_id"])
    summary_url = filters.build_filter_url(activity_id, ec_conditions, qc_conditions)

    try:
        data = _run_async(extraction.load_summary_page(ws, summary_url))
        _state["cached_students"] = list(data["students"])
        _state["current_url"] = summary_url
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": f"数据提取失败: {e}"}), 500


@app.route("/api/students/next")
def students_next():
    """翻到下一页"""
    try:
        ws = _run_async(_ensure_ws())
        result = _run_async(extraction.next_page(ws))
        if result == 'disabled':
            return jsonify({"error": "已是最后一页"}), 400
        data = _run_async(extraction.current_page_data(ws))
        _state["cached_students"].extend(data["students"])
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/students/prev")
def students_prev():
    """翻到上一页"""
    try:
        ws = _run_async(_ensure_ws())
        result = _run_async(extraction.prev_page(ws))
        if result == 'disabled':
            return jsonify({"error": "已是第一页"}), 400
        data = _run_async(extraction.current_page_data(ws))
        # 合并已加载学员（避免重复）
        existing_ids = {s["joinid"] for s in _state["cached_students"]}
        for s in data["students"]:
            if s["joinid"] not in existing_ids:
                _state["cached_students"].append(s)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/generate", methods=["POST"])
def generate_reports():
    """为选中的学员生成成绩单"""
    data = request.get_json()
    joinids = data.get("joinids", [])
    exam_name = data.get("exam_name", _state["exam_name"])
    activity_id = data.get("activity_id", _state["activity_id"])

    if not joinids:
        return jsonify({"error": "未选择学员"}), 400

    # 从缓存中找到对应的学员信息
    selected = [s for s in _state["cached_students"] if s["joinid"] in joinids]
    if not selected:
        return jsonify({"error": "未找到选中的学员数据"}), 400

    # 在后台线程中执行生成
    def _generate():
        try:
            ws = _run_async(_ensure_ws())
            template = _state["template"] or rendering.default_template()
            output_dir = _state["output_dir"]
            os.makedirs(output_dir, exist_ok=True)

            all_results = []
            report_list = []
            for idx, person in enumerate(selected):
                _progress_queue.put({
                    "type": "progress",
                    "current": idx + 1,
                    "total": len(selected),
                    "name": person["name"],
                })

                qs = _run_async(extraction.extract_detail(ws, activity_id, person["joinid"]))
                groups, majors = analysis.analyze(qs)
                all_results.append({"person": person, "questions": qs, "groups": groups, "majors": majors})

                ai_text = rendering.generate_ai_analysis(majors, groups)
                html = rendering.render_html(template, person, groups, majors, ai_analysis=ai_text, exam_name=exam_name)

                # 存内存（供预览）
                _state["reports"][person["joinid"]] = {"name": person["name"], "html": html}
                report_list.append({"joinid": person["joinid"], "name": person["name"],
                                    "score": person["total_score"]})

                # 存本地文件
                filepath = os.path.join(output_dir, f"{exam_name}成绩单-{person['name']}.html")
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(html)

            # JSON 备份
            jp = os.path.join(output_dir, "全部成绩详情.json")
            with open(jp, "w", encoding="utf-8") as f:
                json.dump(all_results, f, ensure_ascii=False, indent=2)

            # CSV 汇总
            cp = os.path.join(output_dir, "成绩汇总.csv")
            with open(cp, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["序号", "姓名", "提交时间", "用时", "总分"] + analysis.MAJOR_SUBJECTS)
                for e in all_results:
                    ep = e["person"]
                    w.writerow(
                        [ep['seq'], ep['name'], ep['submit_time'], ep['duration'], ep['total_score']]
                        + [f"{e['majors'][s]['correct']}/{e['majors'][s]['total']}" for s in analysis.MAJOR_SUBJECTS]
                    )

            _progress_queue.put({
                "type": "done",
                "count": len(all_results),
                "output_dir": os.path.abspath(output_dir),
                "reports": report_list,
            })
        except Exception as e:
            _progress_queue.put({"type": "error", "error": str(e)})

    thread = threading.Thread(target=_generate, daemon=True)
    thread.start()

    return jsonify({"status": "started", "count": len(selected)})


@app.route("/api/report/<joinid>")
def view_report(joinid):
    """预览生成的成绩单"""
    report = _state["reports"].get(joinid)
    if not report:
        return "报告未找到", 404
    return Response(report["html"], mimetype="text/html")


@app.route("/api/progress")
def progress():
    """SSE 端点：推送生成进度"""
    def stream():
        while True:
            try:
                msg = _progress_queue.get(timeout=30)
                yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
                if msg.get("type") in ("done", "error"):
                    break
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'heartbeat'})}\n\n"

    return Response(stream(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/api/config")
def get_config():
    """返回当前配置供前端使用"""
    return jsonify({
        "activity_id": _state["activity_id"],
        "output_dir": os.path.abspath(_state["output_dir"]),
        "exam_name": _state["exam_name"],
        "date_options": filters.DYNAMIC_DATE_OPTIONS,
    })


def configure(cdp_host, cdp_port, activity_id, output_dir, exam_name, template_path=None):
    """配置服务器参数"""
    _state["cdp_host"] = cdp_host
    _state["cdp_port"] = cdp_port
    _state["activity_id"] = activity_id
    _state["output_dir"] = output_dir
    _state["exam_name"] = exam_name
    if template_path:
        with open(template_path, encoding="utf-8") as f:
            _state["template"] = f.read()
