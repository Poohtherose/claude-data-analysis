"""
SPSS ANOVA 统计分析平台 - 启动脚本
"""

import subprocess
import sys
import os

def check_dependencies():
    """检查并安装依赖"""
    print("正在检查依赖...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("依赖安装完成！")
    except subprocess.CalledProcessError as e:
        print(f"依赖安装失败: {e}")
        sys.exit(1)

def start_server():
    """启动Flask服务器"""
    print("""
╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║         SPSS ANOVA 统计分析平台                              ║
║                                                              ║
║  功能: 单因素方差分析 · LSD检验 · Duncan检验 · 方差齐性检验    ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝

正在启动服务器...
    """)

    try:
        from app import app
        print("服务器启动成功！")
        print("请访问: http://localhost:5000")
        print("\n按 Ctrl+C 停止服务器\n")
        app.run(host='0.0.0.0', port=5000, debug=True)
    except Exception as e:
        print(f"启动失败: {e}")
        sys.exit(1)

if __name__ == '__main__':
    # 检查是否在正确的目录
    if not os.path.exists('app.py'):
        print("错误：请在项目根目录运行此脚本")
        sys.exit(1)

    # 检查是否需要安装依赖
    if len(sys.argv) > 1 and sys.argv[1] == '--install':
        check_dependencies()

    start_server()
