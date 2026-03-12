# DoclingParser

这是一个小型工作流项目：使用 Docling 解析 `.xlsx` 表格，并通过 OpenAI 兼容接口（XHANG）对解析后的 Markdown 进行多轮问答。

## 项目结构

- `test.py`：解析 `template/` 目录下的第一个 `.xlsx`，输出 Markdown 到 `output/`
- `chat_with_doc.py`：将解析后的 Markdown 注入聊天上下文，启动交互式问答
- `llm.py`：OpenAI 兼容客户端封装（自动读取 `.env`）
- `template/`：输入的 Excel 文件目录
- `output/`：解析输出的 Markdown 文件目录

## 1）环境准备

建议使用你已有的 `myapp` 环境（或任意 Python 3.11+ 环境）。

```powershell
conda activate myapp
pip install docling openai
```

## 2）配置 API（.env）

在项目根目录创建 `.env`：

```env
XHANG_MODEL=xhang
XHANG_API_KEY=你的_api_key
XHANG_BASE_URL=https://xhang.buaa.edu.cn/xhang/v1
```

说明：
- `llm.py` 会自动读取项目根目录下的 `.env`
- `.env` 已被 `.gitignore` 忽略，不会被 Git 跟踪

## 3）解析 Excel 为 Markdown

将 `.xlsx` 文件放入 `template/` 后执行：

```powershell
python test.py
```

执行后会在 `output/` 生成：
- `output/<excel文件名>.md`

## 4）基于解析文件进行问答

```powershell
c
```

可选参数：

```powershell
python chat_with_doc.pycE:\projecthome\DoclingParser\output\your_file.md --max-chars 30000
```

交互方式：
- 输入问题后回车
- 输入 `/exit` 退出

## 5）上传到 GitHub 前检查

推送前建议执行：

```powershell
git status
```

确认敏感信息未被跟踪。
如果 `.env` 曾误提交到历史，请立即更换 API Key，并清理 Git 历史中的密钥。
