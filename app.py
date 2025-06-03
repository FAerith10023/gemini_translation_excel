from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
import os
import tempfile
import configparser
import re
import time
from werkzeug.utils import secure_filename
from google import genai
from openpyxl import load_workbook
import uuid
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# 配置上传文件夹
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def load_config():
    """加载配置文件"""
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    return config

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def has_chinese(text):
    """检查文本是否包含中文字符"""
    if not isinstance(text, str):
        return False
    chinese_pattern = re.compile(r'[\u4e00-\u9fff]')
    return bool(chinese_pattern.search(text))

def has_english(text):
    """检查文本是否包含英文字符"""
    if not isinstance(text, str):
        return False
    english_pattern = re.compile(r'[a-zA-Z]')
    return bool(english_pattern.search(text))

def parse_excel_range(range_str):
    """解析Excel区域字符串"""
    try:
        if ':' in range_str:
            start_cell, end_cell = range_str.split(':')
        else:
            start_cell = end_cell = range_str
        return start_cell, end_cell
    except:
        raise ValueError("无效的Excel区域格式")

def column_letter_to_number(letter):
    """将列字母转换为数字"""
    result = 0
    for char in letter:
        result = result * 26 + (ord(char.upper()) - ord('A') + 1)
    return result

def parse_cell_reference(cell_ref):
    """解析单元格引用"""
    match = re.match(r'([A-Z]+)(\d+)', cell_ref.upper())
    if not match:
        raise ValueError(f"无效的单元格引用: {cell_ref}")
    
    col_letters, row_num = match.groups()
    col_num = column_letter_to_number(col_letters)
    return int(row_num), col_num

def load_prompt_file(prompt_file_path):
    """加载提示词文件"""
    try:
        with open(prompt_file_path, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except:
        return ""

def translate_with_gemini(client, texts_to_translate, source_lang="中文", target_lang="英文", custom_prompt=""):
    """使用Gemini API批量翻译文本"""
    if not texts_to_translate:
        return {}
    
    # 准备翻译提示
    text_list = "\n".join([f"{i+1}. {text}" for i, text in enumerate(texts_to_translate)])
    
    # 构建基础提示词
    base_prompt = f"""请将以下{source_lang}文本翻译为{target_lang}，保持原文的格式和意思：

{text_list}

请按照以下格式返回翻译结果：
1. [第一句的{target_lang}翻译]
2. [第二句的{target_lang}翻译]
...

注意：
- 保持数字编号
- 如果原文中有特殊格式（如数字、符号等），请保持不变
- 只翻译{source_lang}部分，其他语言保持原样"""

    # 如果有自定义提示词，添加到提示中
    if custom_prompt:
        base_prompt += f"\n\n特殊要求：\n{custom_prompt}"

    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=base_prompt
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

@app.route('/')
def index():
    """主页面"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': '未选择文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '未选择文件'}), 400
        
        if file and allowed_file(file.filename):
            # 生成唯一文件名
            file_id = str(uuid.uuid4())
            filename = secure_filename(file.filename)
            file_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_{filename}")
            file.save(file_path)
            
            # 分析Excel文件结构
            workbook = load_workbook(file_path)
            worksheet = workbook.active
            
            # 获取文件基本信息
            max_row = worksheet.max_row
            max_col = worksheet.max_column
            
            # 预览前几行数据
            preview_data = []
            for row in range(1, min(6, max_row + 1)):
                row_data = []
                for col in range(1, min(11, max_col + 1)):
                    cell = worksheet.cell(row=row, column=col)
                    value = str(cell.value) if cell.value else ""
                    row_data.append(value[:50] + "..." if len(value) > 50 else value)
                preview_data.append(row_data)
            
            return jsonify({
                'success': True,
                'file_id': file_id,
                'filename': filename,
                'max_row': max_row,
                'max_col': max_col,
                'preview': preview_data
            })
        else:
            return jsonify({'error': '文件格式不支持，请上传.xlsx或.xls文件'}), 400
            
    except Exception as e:
        return jsonify({'error': f'文件上传失败: {str(e)}'}), 500

@app.route('/translate', methods=['POST'])
def translate():
    """执行翻译"""
    try:
        data = request.get_json()
        
        file_id = data.get('file_id')
        range_str = data.get('range', 'A1:Z100')
        source_lang = data.get('source_lang', '中文')
        target_lang = data.get('target_lang', '英文')
        custom_prompt = data.get('custom_prompt', '')
        
        # 查找上传的文件
        upload_files = [f for f in os.listdir(UPLOAD_FOLDER) if f.startswith(file_id)]
        if not upload_files:
            return jsonify({'error': '文件未找到'}), 404
        
        input_file_path = os.path.join(UPLOAD_FOLDER, upload_files[0])
        
        # 加载配置和初始化客户端
        config = load_config()
        api_key = config.get('DEFAULT', 'api_key')
        client = genai.Client(api_key=api_key)
        
        # 解析区域
        start_cell, end_cell = parse_excel_range(range_str)
        start_row, start_col = parse_cell_reference(start_cell)
        end_row, end_col = parse_cell_reference(end_cell)
        
        # 加载Excel文件
        workbook = load_workbook(input_file_path)
        worksheet = workbook.active
        
        # 收集需要翻译的文本
        texts_to_translate = []
        cell_text_map = {}
        
        # 根据翻译方向选择检测函数
        text_detector = has_chinese if source_lang == '中文' else has_english
        
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = worksheet.cell(row=row, column=col)
                if cell.value and text_detector(str(cell.value)):
                    text = str(cell.value).strip()
                    texts_to_translate.append(text)
                    cell_text_map[(row, col)] = text
        
        if not texts_to_translate:
            return jsonify({'error': f'在指定区域内未找到{source_lang}内容'}), 400
        
        # 批量翻译
        batch_size = 10
        all_translations = {}
        total_batches = (len(texts_to_translate) + batch_size - 1) // batch_size
        
        for i in range(0, len(texts_to_translate), batch_size):
            batch = texts_to_translate[i:i+batch_size]
            batch_num = i // batch_size + 1
            
            batch_translations = translate_with_gemini(
                client, batch, source_lang, target_lang, custom_prompt
            )
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
        
        # 保存翻译结果
        output_filename = f"translated_{file_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_path = os.path.join(PROCESSED_FOLDER, output_filename)
        workbook.save(output_path)
        
        return jsonify({
            'success': True,
            'translated_count': translated_count,
            'total_count': len(texts_to_translate),
            'output_file': output_filename,
            'download_url': f'/download/{output_filename}'
        })
        
    except Exception as e:
        return jsonify({'error': f'翻译失败: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """下载翻译后的文件"""
    try:
        file_path = os.path.join(PROCESSED_FOLDER, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return "文件未找到", 404
    except Exception as e:
        return f"下载失败: {str(e)}", 500

@app.route('/test_api')
def test_api():
    """测试API连接"""
    try:
        config = load_config()
        api_key = config.get('DEFAULT', 'api_key')
        client = genai.Client(api_key=api_key)
        
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents="测试连接：Hello"
        )
        
        return jsonify({
            'success': True,
            'message': 'API连接正常',
            'response': response.text
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'API连接失败: {str(e)}'
        }), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 