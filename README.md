# 英语试卷生成器

一个基于Python的图形化应用程序，用于自动化生成英语试卷，支持作文、完形填空、阅读理解等多种题型，并可导出为多种格式（HTML、Word、Markdown）。

## 功能特点

- 🎯 **智能试卷生成**：通过AI智能体自动生成完整的英语试卷
- 📝 **题型全面支持**：
  - 作文题
  - 完形填空
  - 阅读理解（A/B/C三篇）
- 📊 **多格式输出**：
  - HTML格式（支持浏览器查看）
  - Word文档（`.docx`格式）
  - Markdown格式（`.md`格式）
- 💾 **历史管理**：
  - 自动保存生成的试卷到SQLite数据库
  - 查看历史试卷列表
  - 查看题目和答案详情
  - 删除历史记录
- 📂 **文件导入**：支持从多种文件格式导入原文内容
  - 纯文本文件（`.txt`）
  - Word文档（`.docx`）
  - Markdown文件（`.md`）
  - HTML文件（`.html`）
- 🖥️ **用户友好界面**：
  - 现代化的图形用户界面
  - 实时生成日志显示
  - 进度条指示
  - 一键打开生成的文件

## 系统要求

- Python 3.7 或更高版本
- 操作系统：Windows / macOS / Linux

## 安装步骤

### 1. 克隆或下载项目
```bash
git clone <项目地址>
# 或直接下载ZIP文件并解压
```

### 2. 安装依赖包
```bash
pip install -r requirements.txt
```

如果没有`requirements.txt`文件，请手动安装以下依赖：
```bash
pip install requests python-docx pillow
```

## 使用方法

### 1. 配置API令牌
在代码中找到以下配置部分，将`AGENT_TOKEN`替换为您自己的API令牌：
```python
AGENT_TOKEN = "pat_DUpJqpypgHeEV904AQpVoLOYPK8oCOBr0oAlPuz96V5Dko6E4Nt2baBKZCuxKqRD"
```

### 2. 运行程序
```bash
python english_exam_generator.py
```

### 3. 生成试卷
![图形界面演示.png](https://pic1.imgdb.cn/item/6940ce9a4a4e4213d00c0fd1.png)
1. **输入试卷内容**（可选）：
   - 在左侧的文本框中输入作文题材、完形填空原文、阅读A/B/C原文
   - 或者点击"导入文件"按钮从现有文件导入内容
   - 如果留空，智能体将随机生成相关内容

2. **开始生成**：
   - 点击"开始生成试卷"按钮
   - 程序将依次调用两个智能体：
     - 试卷生成器：生成完整的试卷内容
     - 答案分离器：分离题目和答案
   - 生成过程中可以在右侧查看实时日志

3. **查看结果**：
   ![生成试卷示例](https://pic1.imgdb.cn/item/6940d0494a4e4213d00c2373.png)
   - 生成完成后，可以点击"一键打开Word文档"查看生成的试卷
   - 生成的试卷将保存为以下文件：
     - `english_comprehensive_exam.html`（HTML格式）
     - `english_comprehensive_exam.docx`（Word格式）
     - `english_comprehensive_exam.md`（Markdown格式）

### 4. 查看历史试卷
![历史试卷界面讲解](https://pic1.imgdb.cn/item/6940cfc84a4e4213d00c1efd.png)
- 点击"查看历史试卷"按钮打开历史记录窗口
- 在左侧列表中选择试卷
- 在右侧查看题目和答案详情
- 支持打开历史试卷的Word文档和HTML文件
- 可以删除不需要的历史记录

## 项目结构

```
english_exam_generator.py        # 主程序文件
exam_database.db                 # SQLite数据库文件（首次运行后生成）
english_comprehensive_exam.html  # 生成的HTML试卷文件
english_comprehensive_exam.docx  # 生成的Word试卷文件
english_comprehensive_exam.md    # 生成的Markdown试卷文件
```

## 代码模块说明

### 主要类
- `EnglishExamGeneratorGUI`：主GUI类，管理整个应用程序的界面和流程
- `ExamDatabase`：数据库操作类，负责试卷数据的存储和检索

### 核心功能函数
- `call_agent_stream_gui()`：调用智能体API，支持流式响应
- `extract_code_block_content()`：从智能体返回内容中提取代码块
- `format_content_to_html()`：将Markdown格式内容转换为HTML
- `markdown_to_word()`：将Markdown格式内容转换为Word文档
- `check_and_install_dependencies()`：检查并安装必要的Python包

## 智能体配置

程序使用两个智能体：
1. **试卷生成器**（OLD_BOT_ID）：生成完整的英语试卷
2. **答案分离器**（NEW_BOT_ID）：从试卷内容中分离题目和答案

您可以根据需要修改以下配置：
```python
OLD_BOT_ID = "7547920557141409807"  # 试卷生成器ID
NEW_BOT_ID = "7572591876910596148"  # 答案分离器ID
USER_ID = "english_test_user_001"   # 用户ID
```

## 自定义设置

### 超时设置
```python
OLD_TIMEOUT = 200  # 试卷生成器超时时间（秒）
NEW_TIMEOUT = 200  # 答案分离器超时时间（秒）
```

### 数据库设置
```python
db_path = "exam_database.db"  # 数据库文件路径
```

## 注意事项

1. **API令牌安全**：请妥善保管您的API令牌，不要分享给他人
2. **网络连接**：生成试卷需要稳定的网络连接以调用智能体API
3. **文件权限**：确保程序对当前目录有读写权限
4. **依赖安装**：首次运行时会自动安装`python-docx`包，需要网络连接
5. **文件覆盖**：每次生成的文件会覆盖同名文件，建议定期备份重要试卷

## 故障排除

### 常见问题

1. **无法生成Word文档**
   - 确保已安装`python-docx`包
   - 检查是否有写文件权限

2. **智能体调用失败**
   - 检查网络连接
   - 验证API令牌是否正确
   - 确认智能体ID是否有效

3. **界面显示异常**
   - 确保使用Python 3.7或更高版本
   - 检查系统是否支持tkinter

4. **文件导入失败**
   - 检查文件格式是否受支持
   - 确保文件编码为UTF-8（中文环境）

### 日志查看
- 程序运行日志显示在右侧文本区域
- 详细错误信息会在日志中显示
- 可以复制日志内容用于故障排查

## 支持与反馈

如有问题或建议，请：
1. 查看项目文档
2. 在GitHub仓库提交Issue
3. 联系开发者

---

**使用提示**：建议在生成试卷后，人工检查题目和答案的准确性，AI生成的内容可能需要进行适当调整和优化。
