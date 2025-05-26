#!/usr/bin/env python3
"""
Markdown to DOCX Web应用启动脚本
"""

import os
import sys
import webbrowser
import time
import threading
from app import app

def open_browser():
    """延迟打开浏览器"""
    time.sleep(1.5)
    webbrowser.open('http://localhost:5000')

def main():
    print("=" * 60)
    print("🚀 Markdown to DOCX Web应用")
    print("=" * 60)
    print("📝 功能特点:")
    print("   • 文件上传转换")
    print("   • 在线编辑器")
    print("   • 中文格式优化")
    print("   • 自定义格式选项")
    print("=" * 60)
    print("🌐 启动Web服务器...")
    print("📍 访问地址: http://localhost:5000")
    print("⏹️  按 Ctrl+C 停止服务")
    print("=" * 60)
    
    # 在新线程中打开浏览器
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()
    
    try:
        # 启动Flask应用
        app.run(
            host='0.0.0.0',
            port=5000,
            debug=False,  # 生产环境关闭debug
            use_reloader=False  # 避免重复打开浏览器
        )
    except KeyboardInterrupt:
        print("\n👋 服务已停止")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 启动失败: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()