# Python CAD 文本批量提取

本仓库归档用户提供的 52 破解论坛帖子代码。附件实际对应“Python 操作 CAD，文本批量提取”，主要功能是批量读取 DWG 图纸中的单行文字和多行文字，并导出为 Excel。

## 主要功能

- 选择一个 CAD 文件夹；
- 扫描其中的 `.dwg` 文件；
- 连接正在运行的 AutoCAD；
- 读取 `Text` 和 `MText` 对象；
- 提取文件名、序号、文字内容以及 X/Y/Z 坐标；
- 将汇总结果保存为 `.xlsx` 文件；
- 使用 PyQt5 界面显示进度，并支持取消操作。

## 批量替换功能

仓库另附 `cad_batch_replace.py`，用于按 Excel 对照表批量替换 DWG 中的 `Text` 和 `MText`。Excel 前两列分别是原文本和替换文本，例如：

| 原文本 | 替换文本 |
| --- | --- |
| x批 | 14批 |
| 项目A | 项目B |

第一行可以写成“原文本 / 替换文本”表头，也可以直接从第一行开始写规则。替换会处理字符串中的部分匹配；如果同一单元格命中多条规则，将按 Excel 从上到下的顺序执行。

命令行示例：

```powershell
python cad_batch_replace.py "D:\\图纸" "D:\\替换表.xlsx" "D:\\图纸_已替换"
```

程序会递归读取源目录下的 DWG，并在输出目录保持相同的子目录结构。原图不会被覆盖，日志写入输出目录的 `cad_batch_replace.log`。程序会启动或连接 AutoCAD；处理前请关闭同名图纸，避免文件被占用。

## 使用软件和环境

- Windows；
- AutoCAD（通过 COM/ActiveX 自动化连接）；
- Python 3；
- PyQt5；
- pyautocad；
- openpyxl。

AutoCAD 需要支持 COM 自动化。当前附件代码只检查所选目录的直接子文件，不会递归扫描更深层目录。

## 原始内容

`original-post.txt` 是用户提供的完整帖子文本，包含论坛页面信息和原始 Python 代码。为保留原始来源，本仓库没有把论坛内容擅自改写成“已验证可运行”的程序。

## 使用前准备

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install PyQt5 pyautocad openpyxl
```

然后需要从 `original-post.txt` 中提取原始界面代码保存为 `.py` 文件，再运行：

```powershell
python cad_text_extractor.py
```

批量替换功能可以直接运行仓库中的 `cad_batch_replace.py`，无需使用原始 PyQt 界面。

## DeepSeek 英文翻译版

`cad_translate_deepseek.py` 可以将 DWG 中的中文 `Text/MText` 翻译为英文并另存到新目录。它使用 DeepSeek 的 OpenAI 兼容接口：术语表中的内容优先使用 Excel 翻译，未匹配文本才调用 DeepSeek，并生成 `translation.csv` 和 `cad_translate.log`。

先复制 `.env.example` 为 `.env` 并填写自己的 Key：

```powershell
Copy-Item .env.example .env
# 编辑 .env，填写 DEEPSEEK_API_KEY
```

准备可选术语表 `glossary.xlsx`，第一列为中文，第二列为固定英文：

| 中文 | English |
| --- | --- |
| 平面图 | Floor Plan |
| 立面图 | Elevation |

运行：

```powershell
python cad_translate_deepseek.py "D:\\中文图纸" "D:\\英文图纸" --glossary "D:\\glossary.xlsx"
```

程序会保持源目录结构，不覆盖原图；API 调用结果会缓存在当前运行内存中，避免同一批任务重复翻译相同文本。翻译费用、术语准确性和图面布局需要自行检查，建议先处理一张副本图纸。

## 图形界面

运行 `python cad_translate_ui.py` 可打开图形界面。界面分为两个独立页签：

- `CAD 翻译（调用 AI）`：填写 API Key，选择 DeepSeek/ChatGPT 和目标语言；可读取术语表 Excel，也可指定保存翻译对照表的位置；单独选择 CAD 源文件夹和英文版输出文件夹。
- 翻译页默认勾选“仅生成翻译表，检查后再替换”：此模式只提取文字并生成中英对照表，不保存修改后的 DWG。检查、修正 Excel 后，到“CAD 文本替换（不调用 AI）”页读取这份确认表，再生成最终图纸。
- `CAD 文本替换（不调用 AI）`：选择 CAD 源文件夹、修改后输出文件夹和替换表 Excel；Excel 第一列是原文本，第二列是替换文本。此页只执行表格替换，不会调用 AI，也不会读取翻译页配置。

DeepSeek 默认使用 `deepseek-chat`，ChatGPT 默认使用 `gpt-4o-mini`，Base URL 和模型名称都可以修改。API Key 只保存在当前进程内存中。

帖子代码引用了 `cad_icon.png`。如果没有该图标，需要删除或注释 `setWindowIcon(QIcon("cad_icon.png"))`，否则界面可能在启动时找不到资源。

## 重要注意事项

- 代码连接并控制本机 AutoCAD，必须在 Windows 和已安装 AutoCAD 的环境运行；macOS/Linux 不能直接使用这套 COM 方案。
- 原代码会打开和关闭图纸，但没有完整记录失败文件；批量处理前应先备份图纸。
- 代码中的 `acad.iter_objects(['Text', 'MText'])`、AutoCAD 版本和 pyautocad 兼容性需要在实际环境验证。
- 这是原始论坛代码归档，未在当前环境运行，也未修改为生产级版本。
- `cad_batch_replace.py` 是在原始代码基础上补充的独立替换模块；由于当前环境没有 Windows AutoCAD，未完成真实 DWG/COM 回归测试。
- `cad_translate_deepseek.py` 同样需要 Windows AutoCAD/ZWCAD；本仓库只完成 Python 语法级检查，未在真实 CAD 和 DeepSeek 账号下验证。
- `.env`、API Key、DWG、Excel 和输出文件已加入忽略范围，不应提交到 Git。
- 论坛帖子和代码未附独立许可证；使用和再分发请确认原作者授权。

## 与“批量替换 CAD 文本”代码的区别

你前面提到的另一段代码是查找并替换 DWG 文本；本仓库附件的实际代码是提取文字并导出 Excel，两者功能不同。
