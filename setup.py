#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 翻译器安装脚本
自动安装依赖并验证环境
"""

import subprocess
import sys
import os
from pathlib import Path

def check_python_version():
    """检查 Python 版本"""
    print("检查 Python 版本...")
    version = sys.version_info
    
    if version.major < 3 or (version.major == 3 and version.minor < 9):
        print(f"❌ Python 版本过低: {version.major}.{version.minor}")
        print("   需要 Python 3.9 或更高版本")
        return False
    else:
        print(f"✅ Python 版本: {version.major}.{version.minor}.{version.micro}")
        return True

def install_dependencies():
    """安装依赖包"""
    print("\n安装依赖包...")
    
    requirements_file = Path("requirements.txt")
    if not requirements_file.exists():
        print("❌ requirements.txt 文件未找到")
        return False
    
    try:
        # 使用 pip 安装依赖
        result = subprocess.run([
            sys.executable, "-m", "pip", "install", "-r", "requirements.txt"
        ], capture_output=True, text=True, check=True)
        
        print("✅ 依赖包安装成功")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ 依赖包安装失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False

def verify_imports():
    """验证关键模块是否可以导入"""
    print("\n验证模块导入...")
    
    modules_to_test = [
        ("google.genai", "Google GenAI SDK"),
        ("openpyxl", "OpenPyXL"),
        ("pandas", "Pandas")
    ]
    
    all_ok = True
    
    for module_name, display_name in modules_to_test:
        try:
            __import__(module_name)
            print(f"✅ {display_name} 导入成功")
        except ImportError as e:
            print(f"❌ {display_name} 导入失败: {e}")
            all_ok = False
    
    return all_ok

def create_sample_files():
    """创建示例文件"""
    print("\n创建示例文件...")
    
    try:
        from create_sample_excel import create_sample_excel
        sample_file = create_sample_excel()
        print(f"✅ 示例文件已创建: {sample_file}")
        return True
    except Exception as e:
        print(f"❌ 创建示例文件失败: {e}")
        return False

def display_usage_info():
    """显示使用说明"""
    print("\n" + "="*50)
    print("🎉 安装完成！")
    print("="*50)
    print("\n使用方法:")
    print("1. 命令行交互模式:")
    print("   python excel_translator.py")
    print("\n2. 编程调用:")
    print("   python example_usage.py")
    print("\n3. 创建示例文件:")
    print("   python create_sample_excel.py")
    print("\n注意事项:")
    print("- 需要从 Google AI Studio 获取 Gemini API 密钥")
    print("- API 密钥获取地址: https://makersuite.google.com/app/apikey")
    print("- 支持 .xlsx 和 .xls 格式的 Excel 文件")
    print("- 程序会自动识别中文内容并翻译成英文")
    print("- 支持合并单元格和多工作表")

def main():
    """主安装流程"""
    print("=" * 50)
    print("Excel 中文翻译器 - 自动安装脚本")
    print("=" * 50)
    
    # 1. 检查 Python 版本
    if not check_python_version():
        print("\n安装失败：Python 版本不满足要求")
        sys.exit(1)
    
    # 2. 安装依赖包
    if not install_dependencies():
        print("\n安装失败：无法安装依赖包")
        sys.exit(1)
    
    # 3. 验证导入
    if not verify_imports():
        print("\n安装失败：模块导入验证失败")
        sys.exit(1)
    
    # 4. 创建示例文件
    create_sample_files()
    
    # 5. 显示使用说明
    display_usage_info()

if __name__ == "__main__":
    main() 