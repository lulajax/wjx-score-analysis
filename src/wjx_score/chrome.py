"""Chrome 进程管理：启动/检测/关闭"""

import os
import shutil
import subprocess
import sys
import time
import urllib.request

# Linux / macOS 候选
_UNIX_CANDIDATES = [
    "google-chrome",
    "google-chrome-stable",
    "chromium",
    "chromium-browser",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
    # macOS
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
]

# Windows 常见安装路径
_WIN_CANDIDATES = [
    os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"),
                 r"Google\Chrome\Application\chrome.exe"),
    os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
                 r"Google\Chrome\Application\chrome.exe"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""),
                 r"Google\Chrome\Application\chrome.exe"),
]


def find_chrome():
    """查找系统中可用的 Chrome/Chromium 路径"""
    # 先尝试 PATH
    for name in ("google-chrome", "google-chrome-stable", "chromium", "chrome"):
        path = shutil.which(name)
        if path:
            return path

    candidates = _WIN_CANDIDATES if sys.platform == "win32" else _UNIX_CANDIDATES
    for p in candidates:
        if p and os.path.isfile(p):
            return p

    # Windows: 尝试从注册表读取
    if sys.platform == "win32":
        try:
            import winreg
            for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                try:
                    key = winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe")
                    val, _ = winreg.QueryValueEx(key, None)
                    winreg.CloseKey(key)
                    if val and os.path.isfile(val):
                        return val
                except OSError:
                    pass
        except ImportError:
            pass

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
    import os, tempfile
    chrome = find_chrome()
    if not chrome:
        raise RuntimeError(
            "未找到 Chrome/Chromium。请安装 Google Chrome 或 Chromium，"
            "或使用 --no-chrome 手动启动: google-chrome --remote-debugging-port=9222"
        )

    # Chrome 已有实例运行时，必须指定独立 user-data-dir 才能启用远程调试
    user_data_dir = os.path.join(
        os.path.expanduser("~"), ".zhanpeng-toolbox", "chrome-profile"
    )
    os.makedirs(user_data_dir, exist_ok=True)

    cmd = [
        chrome,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={user_data_dir}",
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
