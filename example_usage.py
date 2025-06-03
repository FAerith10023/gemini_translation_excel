#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 翻译器使用示例
演示如何以编程方式使用 ExcelTranslator 类
"""

from excel_translator import ExcelTranslator
from create_sample_excel import create_sample_excel
import os

def example_usage():
    """演示 Excel 翻译器的使用方法"""
    
    print("=== Excel 翻译器使用示例 ===\n")
    
    # 1. 创建示例 Excel 文件
    print("1. 创建示例 Excel 文件...")
    sample_file = create_sample_excel()
    print(f"   ✓ 示例文件已创建: {sample_file}\n")
    
    # 2. 获取 API 密钥（实际使用时请替换为您的密钥）
    api_key = input("请输入您的 Gemini API 密钥: ").strip()
    if not api_key:
        print("错误: 需要提供 API 密钥才能运行示例")
        return
    
    # 3. 创建翻译器实例
    print("2. 初始化翻译器...")
    try:
        translator = ExcelTranslator(api_key=api_key)
        print("   ✓ 翻译器初始化成功\n")
    except Exception as e:
        print(f"   ✗ 翻译器初始化失败: {e}")
        return
    
    # 4. 执行翻译
    print("3. 开始翻译...")
    output_file = "sample_chinese_excel_translated.xlsx"
    
    try:
        # 使用技术领域关键词进行翻译
        translator.translate_excel(
            input_file=sample_file,
            output_file=output_file,
            keywords="技术产品"  # 专业领域关键词
        )
        print(f"   ✓ 翻译完成，结果保存至: {output_file}\n")
        
    except Exception as e:
        print(f"   ✗ 翻译失败: {e}")
        return
    
    # 5. 显示结果
    print("4. 翻译结果摘要:")
    print(f"   - 原文件: {sample_file}")
    print(f"   - 译文件: {output_file}")
    
    if os.path.exists(output_file):
        original_size = os.path.getsize(sample_file)
        translated_size = os.path.getsize(output_file)
        print(f"   - 原文件大小: {original_size} bytes")
        print(f"   - 译文件大小: {translated_size} bytes")
        print("   ✓ 翻译文件已成功生成")
    else:
        print("   ✗ 翻译文件未找到")
    
    print("\n=== 示例完成 ===")

def batch_translation_example():
    """批量翻译示例"""
    print("=== 批量翻译示例 ===\n")
    
    # 模拟多个文件的翻译
    files_to_translate = [
        ("file1.xlsx", "技术"),
        ("file2.xlsx", "医学"), 
        ("file3.xlsx", "法律")
    ]
    
    api_key = input("请输入您的 Gemini API 密钥: ").strip()
    if not api_key:
        print("错误: 需要提供 API 密钥")
        return
    
    translator = ExcelTranslator(api_key=api_key)
    
    for input_file, keyword in files_to_translate:
        if os.path.exists(input_file):
            output_file = input_file.replace('.xlsx', '_translated.xlsx')
            print(f"正在翻译: {input_file} (领域: {keyword})")
            
            try:
                translator.translate_excel(input_file, output_file, keyword)
                print(f"✓ 完成: {output_file}")
            except Exception as e:
                print(f"✗ 失败: {e}")
        else:
            print(f"⚠ 文件不存在: {input_file}")

def main():
    """主函数 - 提供交互式菜单"""
    
    while True:
        print("\n=== Excel 翻译器示例菜单 ===")
        print("1. 基本使用示例")
        print("2. 批量翻译示例")
        print("3. 退出")
        
        choice = input("\n请选择操作 (1-3): ").strip()
        
        if choice == '1':
            example_usage()
        elif choice == '2':
            batch_translation_example()
        elif choice == '3':
            print("再见！")
            break
        else:
            print("无效选择，请重试")

if __name__ == "__main__":
    main() 