"""
WSGI 入口文件 - Gunicorn 的启动点
"""
from main import app

if __name__ == "__main__":
    app.run()
