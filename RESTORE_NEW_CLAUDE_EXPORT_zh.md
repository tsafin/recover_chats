# 从 manifest 和 ZIP 文件还原新的 Claude 导出

本文档是一份自包含的操作说明，面向首次接触本目录的人或 LLM 智能体。文中描述了完整流程：解包新的 Claude 导出、还原对话结构与其中生成的文件、尝试重新生成 PPTX/DOCX/XLSX、更新 HTML 界面，以及最终校验。

支持 **Linux/WSL** 和 **Windows**（使用原生 Python、Node.js 以及通过 winget 安装的工具）。

## 1. 本目录的内容

工作脚本：

- `restore_export.py` — 主还原程序，同时生成 HTML 界面；
- `regenerate_python_office.py` — 安全地尝试运行已还原的 Python Office 文件生成器；
- `regenerate_node_office.py` — 尝试运行已还原的 Node.js PPTX/DOCX/XLSX 生成器；
- `regenerate_recipe_docx.py` — 复现保存在 tool calls 中的多步骤 DOCX 制作流程（TeX、绘图、Pandoc 与 `python-docx`）；
- `regenerate_known_office.py` — 复现已明确核验过的制作流程，以及针对本导出中已知 DOCX/XLSX 的精确重命名操作；
- `artifact_timestamps.py` — 根据文件展示记录和 tool calls 恢复重新生成产物的 `mtime`；
- `check_restoration.py` — 校验本地链接和所得 Office 文件的完整性。

辅助文件：

- `winenv.py` — 跨平台模块，用于在 Windows 上自动发现通过 winget 安装的工具，并在生成的源代码中正确替换路径；
- `install_deps_windows.ps1` — PowerShell 脚本，用于在 Windows 上通过 winget、pip 和 npm 自动安装所有系统与语言依赖；
- `restore_all.bat` — 一个用于在 Windows 上从首次运行到校验完整执行整个流程的批处理文件。

主产出物是通过 `--output` 指定的目录。在下文示例中，该目录被称为 `restored_export_new`。

## 2. 输入要求

请将以下内容放入同一个源目录：

- 新的 `manifest-....json` 文件；
- 与其相关的全部 `*.zip`；
- 不要重命名 ZIP 内部的任何内容。

示例：

```text
new_export/
├── manifest-xxxxxxxx.json
├── part-01.zip
├── part-02.zip
└── part-03.zip
```

重要细节：

1. `restore_export.py` 仅在 `--source` 目录的直接层级下查找 `*.zip`，不会递归查找。
2. manifest 文件的名称可以任意。当前版本的还原程序不将 manifest 作为内容来源：manifest 仅用作整套导出的清单，数据从目录下的所有 ZIP 文件中提取。
3. 解包后的数据中必须存在 `conversations.json`。通常它来自某个 ZIP 文件。
4. 不要将不同导出的 ZIP 文件混在同一个 `--source` 中：还原程序会把它们全部解包到公共的 `raw/`。
5. 每套新导出请使用一个新的、空的输出目录。这样可以避免新旧对话混淆。

## 3. 环境要求

### 3.1. 最低要求（所有平台）

- Python 3.10 或更高版本；
- 用于 HTML 渲染 Markdown 的 `markdown-it-py` 包；
- 足够的可用磁盘空间：结果可能远大于 ZIP 总大小。

### 3.2. 重新生成 Office 文件（可选）

额外的包和工具仅在需要从已还原的源码与命令重新生成 PPTX/DOCX/XLSX 时才需要。聊天内容、文本和元数据在没有它们的情况下也能还原。

**Python 库**（所有平台）：

```bash
python -m pip install markdown-it-py openpyxl python-docx python-pptx matplotlib pandas numpy pillow
```

**Node.js**（用于基于 `pptxgenjs` 和 `docx` 的 PPTX/DOCX 生成器）：

- `node` 18+ 和 `npm`；
- 包：`pptxgenjs@3.12.0`、`docx@9.5.1`、`react`、`react-dom`、`react-icons`、`sharp`。

**系统工具**（用于含 TeX/Pandoc 的 DOCX 制作流程）：

- `pandoc` — 将 Markdown/LaTeX 转换为 DOCX；
- `pdflatex`（TeX Live 或 MiKTeX）— 编译 TikZ/PGFPlots 图形；
- `pdftoppm`（Poppler）— 将 PDF 图形转换为 PNG；
- `libreoffice`/`soffice` — 对 DOCX 进行额外的可视化校验。

### 3.3. 在 Windows 上安装依赖

#### 自动安装（推荐）

在 `recover_chats` 目录下运行 PowerShell 脚本：

```powershell
powershell -ExecutionPolicy Bypass -File recover_chats\install_deps_windows.ps1
```

该脚本会以静默模式（无交互输入）通过 `winget` 安装：

| 工具 | winget ID | 用途 |
|---|---|---|
| Pandoc 3.14 | `JohnMacFarlane.Pandoc` | Markdown/LaTeX → DOCX |
| Poppler 25.07 | `oschwartz10612.Poppler` | PDF → PNG（`pdftoppm`） |
| MiKTeX 25.12 | `MiKTeX.MiKTeX` | `pdflatex`、TikZ、PGFPlots |
| LibreOffice 26.8 | `TheDocumentFoundation.LibreOffice` | 校验 DOCX → PDF |

并通过 `pip` 安装所有 Python 库。

在 `restore_export.py` 首次运行后，安装 npm 包：

```cmd
npm install --prefix restored_export_new\runtime pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp --no-audit --no-fund
```

#### 按单条命令手动安装

```cmd
winget install --id JohnMacFarlane.Pandoc --accept-package-agreements --accept-source-agreements
winget install --id oschwartz10612.Poppler --accept-package-agreements --accept-source-agreements
winget install --id MiKTeX.MiKTeX --accept-package-agreements --accept-source-agreements
winget install --id TheDocumentFoundation.LibreOffice --accept-package-agreements --accept-source-agreements
```

`--accept-package-agreements --accept-source-agreements` 标志使安装完全无交互。

#### 脚本如何在 Windows 上找到工具

通过 winget 安装的工具会放在用户级目录（`%LOCALAPPDATA%`）中，这些目录在非交互式会话里不一定会被加入 `PATH`。`winenv.py` 模块会自动：

1. 在 `%LOCALAPPDATA%\Pandoc\` 中查找 Pandoc；
2. 在 `%LOCALAPPDATA%\Programs\MiKTeX\miktex\bin\x64\` 中查找 MiKTeX；
3. 在 winget 包目录下按 `oschwartz10612.Poppler_*` 模板查找 Poppler；
4. 在 `C:\Program Files\LibreOffice\program\` 中查找 LibreOffice；
5. 将找到的目录加入当前 Python 进程的 `PATH`。

所有脚本（`regenerate_*.py`）都会导入 `winenv`，并在启动时调用 `ensure_tools_on_path()` — 无需手动配置 `PATH`。

### 3.4. 在 Linux/WSL 上安装依赖

```bash
sudo apt-get update
sudo DEBIAN_FRONTEND=noninteractive apt-get install -y \
  pandoc \
  texlive-latex-extra \
  texlive-pictures \
  texlive-fonts-recommended \
  texlive-extra-utils \
  poppler-utils \
  libreoffice-core-nogui \
  libreoffice-writer-nogui \
  default-jre-headless
```

Python：

```bash
python3 -m pip install markdown-it-py openpyxl python-docx python-pptx matplotlib pandas numpy pillow cairosvg
```

Node.js：

```bash
npm install --prefix restored_export_new/runtime pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp --no-audit --no-fund
```

各包用途：

- `pandoc` 将准备好的 LaTeX 转换为 DOCX；
- `texlive-latex-extra`、`texlive-pictures` 与 `texlive-fonts-recommended` 提供 `pdflatex`、TikZ、PGFPlots 及字体包；
- `texlive-extra-utils` 提供 `pdfcrop`；
- `poppler-utils` 提供 `pdftoppm`、`pdfinfo` 与 `pdftotext`；
- `libreoffice-*-nogui` 允许在不依赖 X11 的情况下通过转换为 PDF 来校验 DOCX；
- `default-jre-headless` 是部分 LibreOffice 组件所需的。

在服务器和 WSL 中，请安装 `libreoffice-core-nogui` 和 `libreoffice-writer-nogui`。常规的 GUI 包 `libreoffice-writer` 即便使用 `--headless` 参数也可能尝试打开 X11。

### 3.5. 脚本中已考虑的 Windows 特性

1. **路径中的转义序列。** 当把含反斜杠的 Windows 路径替换到生成器的源代码中时（如 `\anthropic` 中的 `\a`、`\flu` 中的 `\f`、`\2026` 中的 `\2`），Python 和 JavaScript 会将其解释为转义序列。所有脚本都使用 `winenv.py` 中的 `to_posix_path()`，它返回以正斜杠表示的路径。

2. **pdflatex 的文件锁定。** 在 Windows 上，`pdflatex` 可能保持 `.log` 文件打开，从而阻碍临时目录的清理（`PermissionError: [WinError 32]`）。在 LaTeX 编译之后的 `finally` 块中会调用 `kill_orphaned_tex_processes()`。

3. **`python3` 与 `python`。** 在 Windows 上，Python 可执行文件通常名为 `python`，而非 `python3`。所有脚本都使用 `sys.executable`，而不是硬编码的 `python3`。

4. **无原生 Cairo 的 `cairosvg`。** 在 Windows 上，`cairosvg` 可以通过 pip 安装，但它依赖原生库 `libcairo-2.dll`，而该库不存在。包含 SVG 图表的 DOCX 制作流程在 `regenerate_known_office.py` 中有备选路径；`cairosvg` 仅在成功导入时才会被使用。

## 4. 安全性

ZIP 文件和已还原的生成器应被视为不可信。

主还原程序：

- 校验每个 ZIP 的 CRC；
- 拒绝 ZIP 内的绝对路径和 `..` 组件；
- 在常规还原过程中不执行提取出的命令；
- 将命令与源码作为审计材料保存。

重新生成脚本确实会运行部分已还原的 Python/JavaScript 代码。它们会丢弃已知的危险构造，并将常见的输出路径重定向到 `regenerated/`，但这只是一种启发式处理，并非正式的沙箱。

`regenerate_recipe_docx.py` 也会执行已还原的中间 Python 脚本，但它会将脚本和源码复制到一次性临时目录中。只有最终经过校验的 DOCX 才会被复制到对话目录。

`regenerate_known_office.py` 包含针对本套导出已明确核验的制作流程清单。它并非通用的识别器：对于其他导出，只有在审计相应 tool calls 和输入文件之后，才应扩展该清单。

因此，重新生成应在以下条件下进行：

- 在容器、虚拟机或隔离的非特权环境中；
- 环境变量中不含密钥或工作令牌；
- 不访问重要目录；
- 最好断网；
- 仅在查看候选清单之后，或在充分理解此风险的前提下。

如果无法接受运行不可信代码，请跳过第 6 节和第 7 节。聊天内容、源码、命令、内嵌文本和元数据仍会被还原。

## 5. 第一步：解包与还原

以下所有命令均在存放本项目 Python 脚本的目录中执行。

**Linux/WSL：**

```bash
python3 restore_export.py \
  --source /absolute/path/to/new_export \
  --output /absolute/path/to/restored_export_new
```

**Windows：**

```cmd
python recover_chats\restore_export.py --source . --output restored_export_new
```

本步骤所做的事：

1. 在 `--source` 的直接层级下查找所有 `*.zip`。
2. 校验 CRC 与路径安全性。
3. 将数据解包到 `<output>/raw/`。
4. 读取 `<output>/raw/conversations.json`。
5. 为每个对话创建独立目录。
6. 从 `Write`、`Edit`、`create_file`、`str_replace`、heredoc 以及受支持的 shell 命令中还原文件状态。
7. 保存命令、内嵌脚本和提取出的附件文本。
8. 生成 `chat.md`、`chat.html`、JSON 清单以及总的 `INDEX.html`。
9. 还原项目文档和 Claude Artifacts 目录（如果导出中存在）。

正常成功的返回码为 `0`。返回码 `2` 表示解包后未找到 `raw/conversations.json`。

### 第一步之后需要检查的内容

请打开：

```text
<output>/archive_report.json
<output>/SUMMARY.json
<output>/INDEX.html
```

在 `archive_report.json` 中，每个归档的 `error` 字段应为 `null`。如果某个归档损坏，不要在未处理的情况下继续：请用正确的副本替换它，并从一个新的输出目录开始。

## 6. 尝试重新生成 Python Office 文件

**Linux/WSL：**

```bash
python3 regenerate_python_office.py /absolute/path/to/restored_export_new/conversations
```

**Windows：**

```cmd
python recover_chats\regenerate_python_office.py restored_export_new\conversations
```

该脚本查找已还原的 `*.py`，这些脚本需满足：

- 使用 `openpyxl`、`python-docx` 或 `python-pptx`；
- 声明输出 PPTX/DOCX/XLSX；
- 不包含已识别的网络、shell 或破坏性调用。

生成的文件放在：

```text
<conversation>/regenerated/
```

报告写入：

```text
<output>/python_regeneration_report.json
```

此处非零返回码并不一定意味着整体还原失败。它表示至少一个符合条件的生成器以错误结束。请查看每条记录的 `status`、`exit_code` 与 `log` 字段。典型原因：

- 原始生成器依赖缺失的图像或表格；
- 导出中没有基础二进制文件；
- 生成器在原始对话中就已存在错误；
- 缺少额外的 Python 库。

在未分析生成器的情况下，不要自动安装未知依赖。

## 7. 尝试重新生成 Node.js Office 文件

请先按第 3 节安装依赖，然后执行：

**Linux/WSL：**

```bash
python3 regenerate_node_office.py \
  restored_export_new/conversations \
  restored_export_new/runtime/node_modules
```

**Windows：**

```cmd
python recover_chats\regenerate_node_office.py restored_export_new\conversations restored_export_new\runtime\node_modules
```

该脚本查找基于 `pptxgenjs` 和 `docx` 的已还原 JavaScript 生成器，将已知输出路径重定向到 `regenerated/`，并跳过已识别为危险或重复的生成器。

报告写入：

```text
<output>/node_regeneration_report.json
```

与 Python 的情况一样，个别的 `failed`、`timeout` 或 `skipped-*` 是可接受的，应根据日志进行解释。不要因为其他生成器的错误而删除已成功生成的文件。

## 8. DOCX 制作流程与必须的 HTML 更新

### 8.1. 按保存的 tool call 制作流程重新生成 DOCX

在安装系统依赖后执行：

**Linux/WSL：**

```bash
python3 regenerate_recipe_docx.py restored_export_new/conversations
```

**Windows：**

```cmd
python recover_chats\regenerate_recipe_docx.py restored_export_new\conversations
```


1. 复制 `paper.tex`、图形文件以及已还原的 `prep.py`、`prep2.py`、`post.py`；
2. 应用后续 `sed`/Python tool calls 中记录的后期修正；
3. 通过 `pdflatex` 编译 TikZ/PGFPlots 图形；
4. 将图形和公式转换为 PNG；
5. 通过 Pandoc 组装中间 DOCX；
6. 通过 `python-docx` 进行排版；
7. 校验 ZIP/CRC 并通过 `python-docx` 重新读取；
8. 额外尝试用 LibreOffice 转换为 PDF。

最终文档放入 `<conversation>/regenerated/`，报告保存在：

```text
<output>/recipe_docx_regeneration_report.json
```

成功的记录应具有 `status: "ok"`。`validation` 字段显示 CRC、段落数、表格数、内嵌图像数以及额外 LibreOffice 校验的结果。

### 8.2. 本导出的已核验 DOCX/XLSX 制作流程

对于本仓库，额外执行：

**Linux/WSL：**

```bash
python3 regenerate_known_office.py \
  restored_export_new/conversations \
  restored_export_new/runtime/node_modules
```

**Windows：**

```cmd
python recover_chats\regenerate_known_office.py restored_export_new\conversations restored_export_new\runtime\node_modules
```

该脚本仅复现已明确映射的制作流程：从已还原的 Markdown 调用 Pandoc、调用已还原的 `md2docx.js`、包含图表的完整 Node.js 制作流程，以及来自 tool calls 的精确 `cp`/rename 操作。每个结果都作为 ZIP/CRC 进行校验，并通过 `python-docx` 或 `openpyxl` 重新打开。报告：`<output>/known_office_regeneration_report.json`。

如果所需名称仅出现在 `presented_files.json` 中，但缺少最终的源码、模板、图像或数据集，脚本不会创建伪造的替代品。此类文件会留在未还原列表中，并附可解释的原因。

### 8.3. 在所有重新生成之后更新 HTML

在任何重新生成之后，请使用相同的 `--source` 和 `--output` 再次运行主程序：

**Linux/WSL：**

```bash
python3 restore_export.py \
  --source /absolute/path/to/new_export \
  --output /absolute/path/to/restored_export_new
```

**Windows：**

```cmd
python recover_chats\restore_export.py --source . --output restored_export_new
```

本步骤的目的：

- 发现已创建在 `regenerated/` 中的文件；
- 将其与相应消息中的提及关联；
- 将其加入右侧的 artifacts 面板；
- 为其赋予来自日志的日期；
- 更新总索引和每个聊天的 HTML。

主还原程序不会删除 `regenerated/`，因此成功生成的 Office 文件会被保留。

如果完全跳过了重新生成，则第二次运行不是必需的。

### 8.4. 恢复文件时间

每个重新生成脚本在成功创建产物后都会调用 `artifact_timestamps.py`，并从对话日志恢复结果文件的修改时间（`mtime`）：

1. 优先选择文件名完全匹配的最后一次 `present_files`/`SendUserFile` 事件；
2. 如果中间文件从未展示给用户，则使用最后一条包含完整文件名的消息或 tool call；
3. 时间、来源和消息 UUID 会写入相应重新生成报告的 `restored_mtime` 或 `restored_mtimes` 字段。

选择最后一次事件是有意的：同一路径可能展示过多个版本，而重新生成的文件代表最终版本。访问时间（`atime`）保持不变。创建时间/birth time 无法跨平台可靠修改，因此 Linux 和 Windows 上保证恢复的是 `mtime`。

已有输出无需重新运行生成器，也可以补写时间：

```bash
python3 artifact_timestamps.py restored_export_new/conversations
```

Windows：

```cmd
python recover_chats\artifact_timestamps.py restored_export_new\conversations
```

## 9. 最终校验

**Linux/WSL：**

```bash
python3 check_restoration.py /absolute/path/to/restored_export_new
```

**Windows：**

```cmd
python recover_chats\check_restoration.py restored_export_new
```

成功标准：

- `broken_links` 等于 `0`；
- `corrupt_office_files` 等于 `0`；
- `html_files` 与索引、Artifacts 索引和对话数量一致；
- `office_files` 与实际重新生成的文件数量一致。

完整报告保存在：

```text
<output>/validation_report.json
```

Office 文件校验会检查 PPTX/DOCX/XLSX 的 ZIP 结构和 CRC。它不保证幻灯片、文档或公式内容的语义正确性。

## 10. 如何使用结果

在浏览器中打开：

```text
<output>/INDEX.html
```

界面包含：

- 左侧 — 按最近活动日期排序的对话列表；
- 对话的搜索与排序；
- 中间 — 选中的 HTML 聊天，含 Markdown 渲染；
- 右侧 — artifacts，含日期、版本和排序；
- PDF、图像、HTML 和文本文件的预览；
- PPTX/DOCX/XLSX 的链接；
- 已还原的文件，含创建日期、最后修改日期和修订次数；
- 已还原文件按日期、名称和修订次数排序。

如果浏览器在通过 `file://` 打开时限制嵌套的本地页面，请启动一个简单的本地 HTTP 服务器：

```bash
python3 -m http.server 8000 \
  --bind 127.0.0.1 \
  --directory /absolute/path/to/restored_export_new
```

然后打开：

```text
http://127.0.0.1:8000/INDEX.html
```

应仅将 HTTP 服务器绑定到 loopback 接口，不要对外发布，尤其是当导出包含敏感数据时。

## 11. 结果结构

```text
restored_export_new/
├── INDEX.html
├── INDEX.md
├── SUMMARY.json
├── archive_report.json
├── restoration_report.json
├── validation_report.json
├── python_regeneration_report.json
├── node_regeneration_report.json
├── recipe_docx_regeneration_report.json
├── known_office_regeneration_report.json
├── raw/
├── project_documents/
└── conversations/
    └── <date>_<title>_<uuid>/
        ├── chat.html
        ├── chat.md
        ├── conversation.json
        ├── README.md
        ├── recovered_files.json
        ├── presented_files.json
        ├── files/
        ├── commands/
        ├── inline_scripts/
        ├── attachments/
        └── regenerated/
```

并非每个目录在每段对话中都必然存在。

## 12. Claude 导出格式的限制

应区分三种情况：

1. **文件存在于 ZIP 中。** 可以精确提取。
2. **日志中保留了源码或生成器。** 可以尝试复现该文件，但不一定与原始文件逐字节一致。
3. **仅保留了路径、UUID 和 MIME 类型。** 没有外部副本则无法还原二进制文件；界面会显示仅元数据可用。

导出通常不包含 Claude 通过 `present_files` 向用户展示的文件字节。因此，聊天中出现 `report.pptx` 的名称并不能证明归档中存在该 PPTX 本身。

提取出的附件文本以 `.extracted.txt` 结尾保存，即使原始附件是 PDF 或 DOCX。这是有意为之：文本转写不应伪装成缺失的二进制原始文件。

## 13. 完整命令序列

### 13.1. Windows（单条命令）

```cmd
recover_chats\restore_all.bat . restored_export_new
```

或手动分步执行：

```cmd
powershell -ExecutionPolicy Bypass -File recover_chats\install_deps_windows.ps1

python recover_chats\restore_export.py --source . --output restored_export_new

python recover_chats\regenerate_python_office.py restored_export_new\conversations

npm install --prefix restored_export_new\runtime pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp --no-audit --no-fund

python recover_chats\regenerate_node_office.py restored_export_new\conversations restored_export_new\runtime\node_modules

python recover_chats\regenerate_recipe_docx.py restored_export_new\conversations

python recover_chats\regenerate_known_office.py restored_export_new\conversations restored_export_new\runtime\node_modules

python recover_chats\restore_export.py --source . --output restored_export_new

python recover_chats\check_restoration.py restored_export_new
```

### 13.2. Linux/WSL

```bash
python3 restore_export.py \
  --source /data/new_export \
  --output /data/restored_export_new

python3 regenerate_python_office.py \
  /data/restored_export_new/conversations

npm install \
  --prefix /data/restored_export_new/runtime \
  pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp \
  --no-audit --no-fund

python3 regenerate_node_office.py \
  /data/restored_export_new/conversations \
  /data/restored_export_new/runtime/node_modules

python3 regenerate_recipe_docx.py \
  /data/restored_export_new/conversations

python3 regenerate_known_office.py \
  /data/restored_export_new/conversations \
  /data/restored_export_new/runtime/node_modules

python3 restore_export.py \
  --source /data/new_export \
  --output /data/restored_export_new

python3 check_restoration.py \
  /data/restored_export_new
```

如果某个重新生成脚本返回了非零退出码，请阅读相应的 JSON 报告，记录错误原因，并继续使用已保留的成功结果。不要仅因个别生成器失败就宣布整个导出无法还原。

## 14. 自主 LLM 智能体的指令

如果此工作由其他 LLM 模型执行，它应：

1. 确定源目录和新的输出目录的绝对路径。
2. 确认源目录中包含同一 manifest 套件的所有 ZIP。
3. 不要将新结果与先前还原的导出混在一起。
4. 检查 Python 和 `markdown_it` 是否可用。
5. 执行首次 `restore_export.py` 运行，并审阅 `archive_report.json`。
6. 在没有适当隔离的情况下，不要运行不可信的生成器。
7. 在允许重新生成的情况下，分别运行 Python、Node、通用 DOCX 制作流程和已核验的 Office 制作流程，并保存各自的报告。
8. 在重新生成之后执行第二次 `restore_export.py` 运行。
9. 运行 `check_restoration.py`，在解释清楚损坏链接或损坏的 Office 文件之前，不要结束任务。
10. 向用户报告 `INDEX.html` 的绝对路径、对话数、文件数、artifacts 数、成功的重新生成次数，以及剩余的限制。
11. 未经用户明确指示，不要删除原始 ZIP、manifest、报告、已还原的生成器或失败的日志。
