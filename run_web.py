#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 翻译器 Web 应用启动脚本
"""

import os
import sys
from pathlib import Path

def check_dependencies():
    """检查必要的依赖是否已安装"""
    try:
        import flask
        import google.genai
        import openpyxl
        print("✅ 所有依赖已安装")
        return True
    except ImportError as e:
        print(f"❌ 缺少依赖: {e}")
        print("请运行以下命令安装依赖:")
        print("pip install -r requirements.txt")
        return False

def create_directories():
    """创建必要的目录"""
    directories = ['uploads', 'downloads', 'static/css', 'static/js', 'templates']
    
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)
    
    print("✅ 目录结构检查完成")

def main():
    """主函数"""
    print("=" * 50)
    print("Excel 中文翻译器 Web 应用")
    print("=" * 50)
    
    # 检查依赖
    if not check_dependencies():
        sys.exit(1)
    
    # 创建目录
    create_directories()
    
    # 启动应用
    try:
        from app import app
        
        print("\n🚀 启动 Web 应用...")
        print("📱 访问地址: http://localhost:5000")
        print("⚠️  按 Ctrl+C 停止服务器")
        print("-" * 50)
        
        # 开发模式启动
        app.run(
            debug=True,
            host='0.0.0.0',
            port=5000,
            threaded=True
        )
        
    except KeyboardInterrupt:
        print("\n\n👋 Web 应用已停止")
    except Exception as e:
        print(f"\n❌ 启动失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 