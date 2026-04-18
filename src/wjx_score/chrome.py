"""Chrome 进程管理：启动/检测/关闭"""

import shutil
import subprocess
import time
import urllib.request

CHROME_CANDIDATES = [
    "google-chrome",
    "google-chrome-stable",
    "chromium",
    "chromium-browser",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
]


def find_chrome():
    """查找系统中可用的 Chrome/Chromium 路径"""
    for name in CHROME_CANDIDATES:
        path = shutil.which(name)
        if path:
            return path
    return None


def is_cdp_available(host="localhost", port=9222):
    """检查 CDP 端口是否可用"""
    try:
        urllib.request.urlopen(f"http://{host}:{port}/json", timeout=2)
        return True
    except Exception:
        return False


def launch_chrome(port=9222):
    """启动 Chrome 并启用远程调试，返回 Popen 对象"""
    chrome = find_chrome()
    if not chrome:
        raise RuntimeError(
            "未找到 Chrome/Chromium。请安装 Google Chrome 或 Chromium，"
            "或使用 --no-chrome 手动启动: google-chrome --remote-debugging-port=9222"
        )

    cmd = [
        chrome,
        f"--remote-debugging-port={port}",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return proc


def wait_for_cdp(host="localhost", port=9222, timeout=15):
    """轮询直到 CDP 就绪"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if is_cdp_available(host, port):
            return True
        time.sleep(1)
    return False
