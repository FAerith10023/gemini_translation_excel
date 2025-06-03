# Excel 中文翻译器 (使用 Gemini AI)

这是一个使用 Google Gemini AI API 自动翻译 Excel 表格中中文内容的 Python 程序。

## 功能特性

- 🔍 **自动检测中文内容**：智能识别 Excel 表格中包含中文的单元格
- 📚 **术语库匹配**：在AI翻译前，先使用术语库进行精确匹配替换
- 🌐 **批量翻译**：使用 Gemini 2.0 Flash 模型进行高质量翻译
- 📊 **合并单元格支持**：正确处理 Excel 中的合并单元格
- 🎯 **专业领域翻译**：支持输入专业领域关键词提高翻译准确性
- 📋 **多工作表处理**：支持处理包含多个工作表的 Excel 文件
- 📝 **详细日志**：提供详细的处理日志和错误信息

## 安装要求

- Python 3.9 或更高版本
- Gemini API 密钥（可从 [Google AI Studio](https://makersuite.google.com/app/apikey) 免费获取）

## 安装步骤

1. 克隆或下载项目文件
2. 安装依赖包：
   ```bash
   pip install -r requirements.txt
   ```

## 使用方法

### Web 界面运行

```bash
python run_web.py
```

然后在浏览器中访问 `http://localhost:5000`

Web界面功能：
1. **输入 Gemini API 密钥**：从 [Google AI Studio](https://makersuite.google.com/app/apikey) 获取
2. **测试 API 连接**：验证密钥是否有效
3. **上传 Excel 文件**：支持 .xlsx 和 .xls 格式
4. **术语库匹配**（可选）：使用预定义术语库进行精确匹配替换
5. **选择专业领域**（可选）：如 "医学"、"法律"、"技术" 等
6. **开始翻译**：对文件进行AI翻译
7. **下载结果**：获取翻译后的文件

### 命令行运行

```bash
python excel_translator.py
```

程序将提示您输入：
1. **Gemini API 密钥**：从 [Google AI Studio](https://makersuite.google.com/app/apikey) 获取
2. **Excel 文件路径**：要翻译的 Excel 文件路径
3. **专业领域关键词**（可选）：如 "医学"、"法律"、"技术" 等

### 程序化调用

```python
from excel_translator import ExcelTranslator

# 创建翻译器实例
translator = ExcelTranslator(api_key="your_gemini_api_key")

# 可选：先进行术语库匹配
matched_count = translator.apply_terminology_matching(
    input_file="input.xlsx",
    output_file="input_matched.xlsx"
)

# 翻译文件（可以是原文件或术语库匹配后的文件）
translator.translate_excel(
    input_file="input_matched.xlsx",
    output_file="output_translated.xlsx", 
    keywords="医学"  # 可选的专业领域关键词
)
```

## 术语库功能

### 术语库格式

术语库文件应为 Excel 格式（.xlsx），包含两列：
- 第一列：中文术语
- 第二列：对应的英文术语

示例术语库文件 `terminology_sample.xlsx` 已包含在项目中。

### 术语库匹配规则

- **精确匹配**：只有单元格内容与术语库中的中文术语完全一致（一字不差不多不少）才会被替换
- **保持格式**：替换后保持原有的Excel格式和合并单元格结构
- **处理优先级**：术语库匹配在AI翻译之前进行，确保专业术语的准确性

### 自定义术语库

1. 创建新的Excel文件
2. 第一列填入中文术语
3. 第二列填入对应的英文术语
4. 将文件命名为 `terminology_sample.xlsx` 放在项目根目录

## 工作流程

### 标准流程
1. **文件分析**：扫描 Excel 文件中的所有工作表和单元格
2. **中文识别**：使用正则表达式识别包含中文字符的单元格
3. **合并单元格处理**：记录并正确处理合并单元格的信息
4. **批量翻译**：将同一工作表的内容一起提交给 Gemini API 翻译
5. **结果应用**：将翻译结果写回对应的单元格位置
6. **文件保存**：生成带有 "_translated" 后缀的新文件

### 术语库增强流程
1. **术语库匹配**：先对文件进行术语库精确匹配替换
2. **生成中间文件**：创建术语库匹配后的文件
3. **AI翻译**：对匹配后的文件进行AI翻译
4. **最终输出**：获得更准确的翻译结果

## 技术实现细节

### 中文检测
使用 Unicode 范围 `[\u4e00-\u9fff]` 检测中文字符，涵盖：
- 中日韩统一表意文字 (CJK Unified Ideographs)
- 扩展A区、B区等

### 合并单元格处理
- 识别所有合并单元格范围
- 只在主单元格（左上角）更新翻译内容
- 保持原有的合并结构不变

### 术语库匹配算法
- 逐单元格扫描，检查内容是否在术语库字典中
- 精确字符串匹配，区分大小写
- 处理合并单元格的特殊情况

### API 调用优化
- 按工作表分批翻译，减少 API 调用次数
- 包含错误重试机制
- 添加请求间隔避免触发频率限制

### 提示词工程
根据 [Gemini API 文档](https://ai.google.dev/gemini-api/docs/quickstart?lang=python&hl=zh-tw) 的最佳实践构建提示词：
- 支持专业领域关键词
- 明确指定翻译格式和要求
- 确保翻译结果的顺序和数量匹配

## 错误处理

程序包含完善的错误处理机制：
- API 调用失败时自动切换到逐个翻译模式
- 文件读写权限检查
- 详细的错误日志记录
- 翻译失败时保留原文并标记
- 术语库加载失败时的降级处理

## 注意事项

1. **API 配额**：Gemini API 有免费配额限制，大文件可能需要付费 API
2. **文件备份**：程序不会修改原文件，会生成新的翻译文件
3. **网络连接**：需要稳定的网络连接访问 Gemini API
4. **文件格式**：支持 .xlsx 和 .xls 格式的 Excel 文件
5. **术语库管理**：建议定期更新和维护术语库以提高翻译质量

## 许可证

本项目基于 MIT 许可证开源。

## 技术支持

如有问题或建议，请创建 Issue 或提交 Pull Request。

---

基于 [Google Gemini API 快速入门指南](https://ai.google.dev/gemini-api/docs/quickstart?lang=python&hl=zh-tw) 开发 