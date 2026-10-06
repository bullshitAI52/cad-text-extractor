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

然后需要从 `original-post.txt` 中提取代码保存为 `.py` 文件，再运行：

```powershell
python cad_text_extractor.py
```

帖子代码引用了 `cad_icon.png`。如果没有该图标，需要删除或注释 `setWindowIcon(QIcon("cad_icon.png"))`，否则界面可能在启动时找不到资源。

## 重要注意事项

- 代码连接并控制本机 AutoCAD，必须在 Windows 和已安装 AutoCAD 的环境运行；macOS/Linux 不能直接使用这套 COM 方案。
- 原代码会打开和关闭图纸，但没有完整记录失败文件；批量处理前应先备份图纸。
- 代码中的 `acad.iter_objects(['Text', 'MText'])`、AutoCAD 版本和 pyautocad 兼容性需要在实际环境验证。
- 这是原始论坛代码归档，未在当前环境运行，也未修改为生产级版本。
- 论坛帖子和代码未附独立许可证；使用和再分发请确认原作者授权。

## 与“批量替换 CAD 文本”代码的区别

你前面提到的另一段代码是查找并替换 DWG 文本；本仓库附件的实际代码是提取文字并导出 Excel，两者功能不同。
