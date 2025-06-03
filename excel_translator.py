import configparser
import pandas as pd
import re
from google import genai
import openpyxl
from openpyxl import load_workbook
import sys
import time

def load_config():
    """加载配置文件"""
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    return config

def has_chinese(text):
    """检查文本是否包含中文字符"""
    if not isinstance(text, str):
        return False
    chinese_pattern = re.compile(r'[\u4e00-\u9fff]')
    return bool(chinese_pattern.search(text))

def parse_excel_range(range_str):
    """解析Excel区域字符串，如 'A1:C10'"""
    try:
        if ':' in range_str:
            start_cell, end_cell = range_str.split(':')
        else:
            start_cell = end_cell = range_str
        return start_cell, end_cell
    except:
        raise ValueError("无效的Excel区域格式，请使用如 'A1:C10' 的格式")

def column_letter_to_number(letter):
    """将列字母转换为数字"""
    result = 0
    for char in letter:
        result = result * 26 + (ord(char.upper()) - ord('A') + 1)
    return result

def parse_cell_reference(cell_ref):
    """解析单元格引用，如 'A1' -> (1, 1)"""
    match = re.match(r'([A-Z]+)(\d+)', cell_ref.upper())
    if not match:
        raise ValueError(f"无效的单元格引用: {cell_ref}")
    
    col_letters, row_num = match.groups()
    col_num = column_letter_to_number(col_letters)
    return int(row_num), col_num

def translate_with_gemini(client, texts_to_translate):
    """使用Gemini API批量翻译文本"""
    if not texts_to_translate:
        return {}
    
    # 准备翻译提示
    text_list = "\n".join([f"{i+1}. {text}" for i, text in enumerate(texts_to_translate)])
    
    prompt = f"""请将以下中文文本翻译为英文，保持原文的格式和意思：

{text_list}

请按照以下格式返回翻译结果：
1. [第一句的英文翻译]
2. [第二句的英文翻译]
...

注意：
- 保持数字编号
- 如果原文中有特殊格式（如数字、符号等），请保持不变
- 只翻译中文部分，其他语言保持原样
"""

    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=prompt
        )
        
        # 解析翻译结果
        translations = {}
        response_lines = response.text.strip().split('\n')
        
        for line in response_lines:
            if line.strip():
                match = re.match(r'(\d+)\.\s*(.*)', line.strip())
                if match:
                    index = int(match.group(1)) - 1
                    translation = match.group(2)
                    if index < len(texts_to_translate):
                        translations[texts_to_translate[index]] = translation
        
        return translations
    except Exception as e:
        print(f"翻译过程中出现错误: {e}")
        return {}

def main():
    try:
        # 加载配置
        config = load_config()
        api_key = config.get('DEFAULT', 'api_key')
        input_file = config.get('FILES', 'input_excel')
        output_file = config.get('FILES', 'output_excel')
        
        # 初始化Gemini客户端
        client = genai.Client(api_key=api_key)
        print("Gemini API客户端初始化成功")
        
        # 获取用户输入的Excel区域
        print("\n请输入要翻译的Excel区域:")
        print("示例: A1:C10 (翻译A1到C10区域)")
        print("示例: B2:D20 (翻译B2到D20区域)")
        print("示例: A1 (只翻译A1单元格)")
        
        range_input = input("请输入区域: ").strip()
        if not range_input:
            print("未输入区域，程序退出")
            return
        
        # 解析区域
        start_cell, end_cell = parse_excel_range(range_input)
        start_row, start_col = parse_cell_reference(start_cell)
        end_row, end_col = parse_cell_reference(end_cell)
        
        print(f"将翻译区域: {start_cell}:{end_cell}")
        print(f"行范围: {start_row}-{end_row}, 列范围: {start_col}-{end_col}")
        
        # 加载Excel文件
        print(f"正在加载Excel文件: {input_file}")
        workbook = load_workbook(input_file)
        worksheet = workbook.active
        
        # 收集需要翻译的文本
        texts_to_translate = []
        cell_text_map = {}  # 存储单元格位置和文本的映射
        
        print("正在扫描中文内容...")
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = worksheet.cell(row=row, column=col)
                if cell.value and has_chinese(str(cell.value)):
                    text = str(cell.value).strip()
                    texts_to_translate.append(text)
                    cell_text_map[(row, col)] = text
        
        if not texts_to_translate:
            print("在指定区域内未找到中文内容")
            return
        
        print(f"找到 {len(texts_to_translate)} 个包含中文的单元格")
        
        # 显示将要翻译的内容
        print("\n将要翻译的内容:")
        for i, text in enumerate(texts_to_translate[:5]):  # 只显示前5个
            print(f"{i+1}. {text}")
        if len(texts_to_translate) > 5:
            print(f"... 以及其他 {len(texts_to_translate) - 5} 项")
        
        # 确认是否继续
        confirm = input("\n是否继续翻译? (y/n): ").strip().lower()
        if confirm != 'y':
            print("用户取消翻译")
            return
        
        # 批量翻译
        print("\n正在使用Gemini API翻译...")
        
        # 分批处理，避免请求过大
        batch_size = 10
        all_translations = {}
        
        for i in range(0, len(texts_to_translate), batch_size):
            batch = texts_to_translate[i:i+batch_size]
            print(f"正在翻译第 {i//batch_size + 1} 批 ({len(batch)} 项)...")
            
            batch_translations = translate_with_gemini(client, batch)
            all_translations.update(batch_translations)
            
            # 避免API频率限制
            if i + batch_size < len(texts_to_translate):
                time.sleep(1)
        
        # 应用翻译结果
        translated_count = 0
        for (row, col), original_text in cell_text_map.items():
            if original_text in all_translations:
                cell = worksheet.cell(row=row, column=col)
                cell.value = all_translations[original_text]
                translated_count += 1
                print(f"翻译: {original_text[:30]}... -> {all_translations[original_text][:30]}...")
        
        # 保存结果
        print(f"\n正在保存翻译结果到: {output_file}")
        workbook.save(output_file)
        
        print(f"翻译完成！")
        print(f"- 总共处理: {len(texts_to_translate)} 个单元格")
        print(f"- 成功翻译: {translated_count} 个单元格")
        print(f"- 输出文件: {output_file}")
        
    except FileNotFoundError as e:
        print(f"文件未找到: {e}")
    except Exception as e:
        print(f"程序执行出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main() 