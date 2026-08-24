# Vercel Serverless 入口：把根目录加入 sys.path 后导出 WSGI 应用
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app as application  # noqa: E402,F401
