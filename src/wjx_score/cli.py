"""CLI 入口：启动 Chrome + Flask Web 服务"""

import argparse
import atexit
import os
import signal
import socket
import sys
import webbrowser

from . import chrome, server, filters

# PID 文件路径：放在用户目录下，避免权限问题
_PID_DIR = os.path.join(os.path.expanduser("~"), ".zhanpeng-toolbox")
_PID_FILE = os.path.join(_PID_DIR, "server.pid")


def _is_port_in_use(port):
    """检测端口是否被占用"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(("0.0.0.0", port))
            return False
        except OSError:
            return True


def _read_pid():
    """读取 PID 文件，返回 (pid, port) 或 None"""
    try:
        with open(_PID_FILE, "r") as f:
            parts = f.read().strip().split(",")
            return int(parts[0]), int(parts[1]) if len(parts) > 1 else 0
    except (FileNotFoundError, ValueError):
        return None


def _write_pid(port):
    """写入当前进程 PID 和端口"""
    os.makedirs(_PID_DIR, exist_ok=True)
    with open(_PID_FILE, "w") as f:
        f.write(f"{os.getpid()},{port}")


def _remove_pid():
    """清理 PID 文件"""
    try:
        os.remove(_PID_FILE)
    except FileNotFoundError:
        pass


def _is_process_alive(pid):
    """检查进程是否还在运行"""
    if sys.platform == "win32":
        import ctypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
        if handle:
            kernel32.CloseHandle(handle)
            return True
        return False
    else:
        try:
            os.kill(pid, 0)
            return True
        except (ProcessLookupError, PermissionError):
            return False


def _kill_process(pid):
    """终止进程"""
    try:
        if sys.platform == "win32":
            import subprocess
            subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            os.kill(pid, signal.SIGTERM)
        return True
    except Exception:
        return False


def _cleanup_old_instance(port):
    """清理旧实例：检查 PID 文件和端口占用"""
    info = _read_pid()
    if info:
        old_pid, old_port = info
        if _is_process_alive(old_pid):
            print(f"检测到旧服务 (PID={old_pid}, 端口={old_port})，正在关闭...")
            _kill_process(old_pid)
            # 等待进程退出
            for _ in range(30):  # 最多等 3 秒
                if not _is_process_alive(old_pid):
                    break
                import time
                time.sleep(0.1)
            if _is_process_alive(old_pid):
                print(f"  警告: 旧进程 {old_pid} 未能正常关闭")
            else:
                print("  旧服务已关闭")
        _remove_pid()

    # 再检查端口是否仍被占用（可能是非本程序的进程占用）
    if _is_port_in_use(port):
        print(f"错误: 端口 {port} 被其他程序占用")
        if sys.platform != "win32":
            print(f"  排查: lsof -i:{port}")
        else:
            print(f"  排查: netstat -ano | findstr :{port}")
        print(f"  或使用其他端口: --port {port + 1}")
        sys.exit(1)


def main():
    ap = argparse.ArgumentParser(
        description="展鹏教育工具箱 (Web UI)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="启动后在浏览器中操作：成绩单生成、学员登记等。",
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
    ap.add_argument("--xlsx", default="./新学员登记表.xlsx",
                    help="学员登记表 Excel 路径 (默认: ./新学员登记表.xlsx)")
    args = ap.parse_args()

    chrome_proc = None

    # 清理旧实例
    _cleanup_old_instance(args.port)

    # 写入 PID 文件 + 注册退出清理
    _write_pid(args.port)
    atexit.register(_remove_pid)

    def cleanup(signum=None, frame=None):
        _remove_pid()
        if chrome_proc:
            print("\n正在关闭 Chrome...")
            chrome_proc.terminate()
        sys.exit(0)

    signal.signal(signal.SIGINT, cleanup)
    if sys.platform != "win32":
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
        xlsx_path=args.xlsx,
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
