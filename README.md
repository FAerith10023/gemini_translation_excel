# Excel 翻译工具 - 命令行版 + Web版

这是一个使用 Google Gemini-2.0-flash API 将 Excel 表格中的内容进行多语言翻译的工具，支持命令行和Web界面两种使用方式。

## 🌟 功能特点

- 🔤 使用 Google Gemini-2.0-flash API 进行高质量翻译
- 📊 支持指定 Excel 区域进行精确翻译
- 🌐 **新增Web界面**：直观的拖拽上传和可视化操作
- 🎯 支持多种语言翻译（中文、英文、日文、韩文、法文等）
- 🚀 批量处理，智能检测源语言内容
- 💾 保持原有格式，输出到新文件
- ⚡ 智能分批处理，避免 API 限制
- 🎨 **自定义提示词**：支持专业术语和特殊翻译要求

## 📦 安装依赖

```bash
pip install -r requirements.txt
```

## ⚙️ 配置设置

在 `config.ini` 文件中配置你的设置：

```ini
[DEFAULT]
api_key = 你的_GEMINI_API_密钥

[FILES]
input_excel = input.xlsx    # 输入的Excel文件
output_excel = output.xlsx  # 输出的Excel文件
glossary_excel = glossary.xlsx  # 术语表文件（可选）

[LANGUAGE]
source = 中文    # 源语言
target = 英文    # 目标语言
```

### 获取 Gemini API 密钥

1. 访问 [Google AI Studio](https://ai.google.dev/)
2. 登录你的 Google 账户
3. 创建新的 API 密钥
4. 将密钥复制到 `config.ini` 文件中

## 🚀 使用方法

### 方式一：Web界面（推荐）

1. **启动Web服务器**：
   ```bash
   python run_web.py
   ```

2. **打开浏览器**：
   - 访问：http://localhost:5000
   - 支持现代浏览器（Chrome、Firefox、Safari、Edge）

3. **使用Web界面**：
   - 📁 **上传文件**：拖拽或点击上传Excel文件
   - 👀 **预览内容**：查看文件结构和前几行数据
   - ⚙️ **配置翻译**：
     - 选择源语言和目标语言
     - 设置翻译区域（如A1:D20）
     - 添加自定义提示词
   - 🔧 **测试API**：点击测试按钮验证连接
   - ▶️ **开始翻译**：一键开始翻译过程
   - 📥 **下载结果**：翻译完成后直接下载

### 方式二：命令行版本

1. **准备文件**：
   - 将要翻译的 Excel 文件重命名为 `input.xlsx` 或修改配置文件中的文件名
   - 确保文件在项目目录中

2. **运行翻译工具**：
   ```bash
   python excel_translator.py
   ```

3. **输入区域**：
   - 程序会提示输入要翻译的 Excel 区域
   - 支持的格式：
     - `A1:C10` - 翻译 A1 到 C10 的矩形区域
     - `B2:D20` - 翻译 B2 到 D20 的矩形区域
     - `A1` - 只翻译 A1 单元格

## 🎯 自定义提示词功能

### 通过Web界面

1. 在"自定义提示词"区域输入特殊要求
2. 或上传`.txt`格式的提示词文件
3. 示例提示词（参考`example_prompt.txt`）：
   ```
   量词翻译规则：
   - "个" 翻译为 "nos"
   - "台" 翻译为 "units"
   - "套" 翻译为 "sets"
   
   技术术语保持准确性：
   - 保持品牌名称不变
   - 使用正式的商务英语
   ```

## 📊 Web界面特性

- **拖拽上传**：支持直接拖拽Excel文件
- **实时预览**：上传后立即预览文件内容
- **多语言支持**：界面支持中英双语
- **进度显示**：翻译过程实时反馈
- **响应式设计**：支持手机和平板访问
- **API状态监控**：实时检测API连接状态

## 📁 项目结构

```
excel-translator-web-gemini/
├── app.py                      # Flask Web应用主文件
├── run_web.py                  # Web启动脚本
├── excel_translator.py        # 命令行翻译脚本
├── templates/
│   └── index.html             # Web界面模板
├── static/                    # 静态文件目录
├── uploads/                   # 上传文件存储
├── processed/                 # 处理后文件存储
├── config.ini                 # 配置文件
├── requirements.txt           # Python依赖
├── example_prompt.txt         # 示例提示词文件
├── README.md                  # 说明文档
├── input.xlsx                # 输入Excel文件（命令行版）
└── output.xlsx               # 输出Excel文件（自动生成）
```

## 🌐 Web API 接口

如果你想集成到其他系统，可以使用以下API接口：

- `POST /upload` - 上传Excel文件
- `POST /translate` - 执行翻译
- `GET /download/<filename>` - 下载结果文件
- `GET /test_api` - 测试API连接

## 📱 使用示例

### Web界面操作流程

1. **启动服务**：
   ```bash
   python run_web.py
   ```

2. **打开浏览器**，访问 http://localhost:5000

3. **上传文件**：拖拽Excel文件到上传区域

4. **配置翻译**：
   - 源语言：中文
   - 目标语言：英文
   - 翻译区域：A1:D20
   - 提示词：量词"个"翻译为"nos"

5. **开始翻译**：点击"开始翻译"按钮

6. **下载结果**：翻译完成后点击"下载翻译结果"

### 命令行操作示例

```
请输入要翻译的Excel区域:
示例: A1:C10 (翻译A1到C10区域)
示例: B2:D20 (翻译B2到D20区域)
示例: A1 (只翻译A1单元格)

请输入区域: A1:D20

将翻译区域: A1:D20
行范围: 1-20, 列范围: 1-4
正在加载Excel文件: input.xlsx
正在扫描中文内容...
找到 15 个包含中文的单元格

将要翻译的内容:
1. 产品名称
2. 价格描述
3. 技术规格
4. 使用说明
5. 注意事项
... 以及其他 10 项

是否继续翻译? (y/n): y

正在使用Gemini API翻译...
正在翻译第 1 批 (10 项)...
正在翻译第 2 批 (5 项)...

翻译完成！
- 总共处理: 15 个单元格
- 成功翻译: 15 个单元格
- 输出文件: output.xlsx
```

## ⚠️ 注意事项

- 🔑 确保你的 Gemini API 密钥有效且有足够的配额
- 📝 程序会自动检测源语言内容进行翻译
- 🔄 保持原有的Excel格式和样式
- ⏱️ 大量内容翻译需要一些时间，请耐心等待
- 💰 API 调用会产生费用，请注意你的使用量
- 🌐 Web版本支持并发用户，但共享API配额

## 🔧 故障排除

### 常见问题

1. **API 密钥错误**：
   - 检查 `config.ini` 中的 API 密钥是否正确
   - 确保密钥有效且未过期

2. **文件未找到**：
   - Web版：确保上传的文件格式正确（.xlsx/.xls）
   - 命令行版：确保 `input.xlsx` 文件存在

3. **翻译失败**：
   - 检查网络连接
   - 确认 API 配额是否充足
   - 查看错误消息获取详细信息

4. **Web服务启动失败**：
   ```bash
   # 检查端口是否被占用
   netstat -an | findstr :5000
   
   # 或使用不同端口启动
   python -c "from app import app; app.run(port=5001)"
   ```

5. **依赖包问题**：
   ```bash
   pip install --upgrade -r requirements.txt
   ```

## 🚀 高级功能

### 批量处理多个文件

Web版支持逐个处理多个文件，每次上传新文件会自动清理之前的临时文件。

### 自定义翻译区域

支持复杂的Excel区域表达式：
- `A1:D20` - 矩形区域
- `A1,C3,E5` - 多个单独单元格
- `A:A` - 整列
- `1:1` - 整行

### API 限制处理

程序自动处理API频率限制：
- 分批处理大量文本
- 智能重试机制
- 错误恢复

## 🌟 技术特性

- **前端**：Bootstrap 5 + 原生JavaScript
- **后端**：Flask + Python 3.7+
- **AI模型**：Google Gemini-2.0-flash
- **文件处理**：openpyxl + pandas
- **安全性**：文件类型验证、大小限制、路径安全

## 📞 技术支持

如有问题或建议，请检查：
- [Gemini API 文档](https://ai.google.dev/gemini-api/docs/quickstart?lang=python&hl=zh-tw)
- [Google AI Studio](https://ai.google.dev/)
- [Flask 文档](https://flask.palletsprojects.com/)

## 📄 许可证

本项目遵循 MIT 许可证。 #   g e m i n i _ t r a n s l a t i o n _ e x c e l  
 