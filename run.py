#!/usr/bin/env python
import subprocess
import sys
import os
from pathlib import Path
from dotenv import load_dotenv

def main():
    # 检查并创建.env文件
    project_root = Path(__file__).parent
    env_file = project_root / ".env"
    env_example = project_root / ".env.example"

    if not env_file.exists() and env_example.exists():
        import shutil
        shutil.copy(env_example, env_file)
        print("✓ 已从 .env.example 创建 .env 文件")

    # 加载.env配置
    load_dotenv(env_file)
    ADMIN_KEY = os.getenv("ADMIN_KEY", "admin123456")
    SERVER_PORT = os.getenv("SERVER_PORT", "8000")
    SYSTEM_NAME = os.getenv("SYSTEM_NAME", "共享表格")

    # 切换到backend目录
    backend_dir = os.path.join(os.path.dirname(__file__), 'backend')
    backend_dir = os.path.abspath(backend_dir)
    os.chdir(backend_dir)

    # 添加当前目录到 Python 路径
    sys.path.insert(0, backend_dir)

    # 安装依赖
    # print("检查依赖...")
    # subprocess.run([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt', '-q'])

    # 启动服务器
    print("\n" + "="*50)
    print(f"{SYSTEM_NAME} 多人共享编辑表格")
    print("="*50)
    print("\n访问地址:")
    print(f"  - 管理后台: http://localhost:{SERVER_PORT}/admin/{ADMIN_KEY}")
    print(f"  - 表格访问: http://localhost:{SERVER_PORT}/sheet/密钥")
    print("\n按 Ctrl+C 停止服务器")
    print("="*50 + "\n")

    # 自动打开浏览器
    import webbrowser
    import threading

    def open_browser():
        import time
        time.sleep(1.5)  # 等待服务器启动
        admin_url = f"http://localhost:{SERVER_PORT}/admin/{ADMIN_KEY}"
        webbrowser.open(admin_url)

    # 在后台线程中打开浏览器
    threading.Thread(target=open_browser, daemon=True).start()

    # 使用uvicorn启动
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=int(SERVER_PORT), reload=True)

if __name__ == "__main__":
    main()
