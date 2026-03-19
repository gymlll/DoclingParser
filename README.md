# DoclingParser

这是一个小型工作流项目：使用 Docling 解析 `.xlsx` 表格，并通过 OpenAI 兼容接口（XHANG）对解析后的 Markdown 进行多轮问答。

## 项目结构

- `test.py`：解析 `template/` 目录下的 `.xlsx`，输出到 `output/`
- `test2.py`：解析含图片的 Excel（直接处理 `template/` 下的 `.xlsx`），重写图片路径到 `images/`，输出到 `output/`
- `test3.py`：解析含嵌入式图表以及图片的 Excel（导出 ChartObjects 与 Shapes 的图表为 PNG 并合并到 Markdown），递归处理 `template/`，输出到 `output/`
- `chat_with_doc.py`：将解析后的 Markdown 注入聊天上下文，启动交互式问答
- `llm.py`：OpenAI 兼容客户端封装（自动读取 `.env`）
- `template/`：输入的 Excel 文件目录
- `output/`：解析输出的 Markdown 文件目录

## 1）环境准备

建议使用你已有的 `myapp` 环境（或任意 Python 3.11+ 环境）。

```powershell
conda activate myapp
pip install docling openai pywin32
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

执行后会在 `outputq/` 生成：

- `outputq/<excel文件名>/index.md`

### 解析含图片的 Excel（直接处理 `.xlsx`）

`test2.py` 递归处理 `template1/` 下的 `.xlsx`，用 Docling 解析为 Markdown，并重写图片到 `images/`。

```powershell
python test2.py
```

输出目录为 `outputq/<文件名>/`，包含 `index.md` 和 `images/`。

### 可选：保留图表并解析（适用于含图表的 `.xlsx`）

`test3.py` 通过 Excel COM（pywin32）导出嵌入式图表，并合并到 Markdown：

- 同时遍历 ChartObjects 与 Shapes(HasChart) 并导出为 PNG
- Docling 用于解析表格与原始图片（不涉及 PDF 转换）

```powershell
python test3.py
```

说明：

- 需要 Windows + 安装 Microsoft Excel（COM 导出依赖）
- 输出目录为 `output/<文件名>/`，包含 `index.md` 和 `images/` 下的图表 PNG

## 4）基于解析文件进行问答

```powershell
python chat_with_doc.py E:\projecthome\DoclingParser\output\your_file.md --max-chars 30000
```

交互方式：

- 输入问题后回车
- 输入 `/exit` 退出
