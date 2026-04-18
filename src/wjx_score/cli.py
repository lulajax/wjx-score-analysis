"""CLI 入口：启动 Chrome + Flask Web 服务"""

import argparse
import signal
import sys
import webbrowser

from . import chrome, server, filters


def main():
    ap = argparse.ArgumentParser(
        description="展鹏教育 - 问卷星成绩单生成工具 (Web UI)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="启动后在浏览器中操作：筛选、查询、选择学员、生成成绩单。",
    )
    ap.add_argument("--port", type=int, default=8080, help="Web UI 端口 (默认: 8080)")
    ap.add_argument("--cdp-host", default="localhost", help="Chrome CDP 主机 (默认: localhost)")
    ap.add_argument("--cdp-port", type=int, default=9222, help="Chrome CDP 端口 (默认: 9222)")
    ap.add_argument("--output-dir", default="./成绩单", help="成绩单输出目录 (默认: ./成绩单)")
    ap.add_argument("--exam-name", default="学前测", help="考试名称 (默认: 学前测)")
    ap.add_argument("--activity", default=filters.DEFAULT_ACTIVITY_ID,
                    help=f"活动 ID (默认: {filters.DEFAULT_ACTIVITY_ID})")
    ap.add_argument("--template", default=None, help="自定义 HTML 模板路径")
    ap.add_argument("--no-chrome", action="store_true",
                    help="不自动启动 Chrome (已有 Chrome 运行时使用)")
    ap.add_argument("--no-browser", action="store_true",
                    help="不自动打开浏览器")
    args = ap.parse_args()

    chrome_proc = None

    def cleanup(signum=None, frame=None):
        if chrome_proc:
            print("\n正在关闭 Chrome...")
            chrome_proc.terminate()
        sys.exit(0)

    signal.signal(signal.SIGINT, cleanup)
    signal.signal(signal.SIGTERM, cleanup)

    # 启动 Chrome
    if not args.no_chrome:
        if chrome.is_cdp_available(args.cdp_host, args.cdp_port):
            print(f"检测到 CDP 已在 {args.cdp_host}:{args.cdp_port} 运行")
        else:
            print("启动 Chrome...")
            try:
                chrome_proc = chrome.launch_chrome(args.cdp_port)
            except RuntimeError as e:
                print(f"错误: {e}")
                sys.exit(1)

            print(f"等待 CDP 就绪 (端口 {args.cdp_port})...")
            if not chrome.wait_for_cdp(args.cdp_host, args.cdp_port):
                print("错误: Chrome CDP 启动超时")
                cleanup()
    else:
        if not chrome.is_cdp_available(args.cdp_host, args.cdp_port):
            print(f"错误: CDP 未就绪 ({args.cdp_host}:{args.cdp_port})")
            print(f"请先启动 Chrome: google-chrome --remote-debugging-port={args.cdp_port}")
            sys.exit(1)

    # 配置 Flask 服务
    server.configure(
        cdp_host=args.cdp_host,
        cdp_port=args.cdp_port,
        activity_id=args.activity,
        output_dir=args.output_dir,
        exam_name=args.exam_name,
        template_path=args.template,
    )

    url = f"http://localhost:{args.port}"
    print(f"\n{'='*50}")
    print(f"  Web UI 已就绪: {url}")
    print(f"  输出目录: {args.output_dir}")
    print(f"  按 Ctrl+C 退出")
    print(f"{'='*50}\n")

    # 打开浏览器
    if not args.no_browser:
        webbrowser.open(url)

    # 启动 Flask
    try:
        server.app.run(host="0.0.0.0", port=args.port, debug=False, threaded=True)
    finally:
        cleanup()
