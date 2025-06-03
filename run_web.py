#!/usr/bin/env python
"""
Excel翻译工具 Web版 启动脚本
"""

import os
import sys
from app import app

def main():
    print("=" * 60)
    print(" Excel 翻译工具 - Web版")
    print("=" * 60)
    print(" 🌐 启动Web服务器...")
    print(" 📱 Web界面: http://localhost:5000")
    print(" 🔧 API连接测试: http://localhost:5000/test_api")
    print("=" * 60)
    print()
    
    # 检查配置文件
    if not os.path.exists('config.ini'):
        print("⚠️  警告: config.ini 文件未找到")
        print("   请确保配置文件存在并包含有效的Gemini API密钥")
        print()
    
    # 创建必要的目录
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('processed', exist_ok=True)
    
    try:
        # 启动Flask应用
        app.run(
            debug=True,
            host='0.0.0.0',
            port=5000,
            use_reloader=True
        )
    except KeyboardInterrupt:
        print("\n👋 服务器已停止")
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main() 