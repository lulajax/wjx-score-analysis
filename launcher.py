"""展鹏教育工具箱 - 独立启动入口（PyInstaller 用）"""

import sys
import os

# PyInstaller 打包后，资源文件在 _MEIPASS 目录下
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))

from wjx_score.cli import main

if __name__ == "__main__":
    main()
