"""CAD 翻译/文本替换图形界面。运行：python cad_translate_ui.py"""
from __future__ import annotations
import sys
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QFileDialog, QCheckBox, QComboBox, QFormLayout, QHBoxLayout, QLabel, QLineEdit, QMainWindow, QMessageBox, QProgressBar, QPushButton, QTabWidget, QVBoxLayout, QWidget
from cad_batch_replace import run as replace_run
from cad_translate_deepseek import run as translate_run

def chooser(edit, parent, directory=False, save=False):
    if directory: value = QFileDialog.getExistingDirectory(parent)
    elif save: value = QFileDialog.getSaveFileName(parent, "保存 Excel", "翻译对照表.xlsx", "Excel (*.xlsx)")[0]
    else: value = QFileDialog.getOpenFileName(parent, "选择 Excel", "", "Excel (*.xlsx)")[0]
    if value: edit.setText(value)

def path_row(edit, parent, directory=False, save=False):
    button = QPushButton("选择"); button.clicked.connect(lambda: chooser(edit, parent, directory, save))
    row = QHBoxLayout(); row.addWidget(edit); row.addWidget(button); widget = QWidget(); widget.setLayout(row); return widget

class Worker(QThread):
    progress = pyqtSignal(int, str); finished = pyqtSignal(bool, str)
    def __init__(self, function, values): super().__init__(); self.function = function; self.values = values
    def run(self):
        try:
            def report(index, total, name): self.progress.emit(int(index * 100 / max(total, 1)), f"正在处理 {index}/{total}：{name}")
            result = self.function(progress=report, **self.values)
            self.finished.emit(True, f"完成：处理 {result[0]} 个文件，修改 {result[1]} 个文字对象，失败 {result[2]} 个。")
        except Exception as exc: self.finished.emit(False, str(exc))

class Window(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle("CAD 图纸批处理工具"); self.resize(860, 560); self.worker = None
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #f4f7fb; color: #1f2937; font-size: 14px; }
            QTabWidget::pane { border: 1px solid #dbe4f0; background: #ffffff; border-radius: 10px; }
            QTabBar::tab { background: #e8eef7; padding: 12px 20px; margin-right: 3px; border-radius: 8px 8px 0 0; color: #475569; }
            QTabBar::tab:selected { background: #2563eb; color: white; font-weight: bold; }
            QLineEdit, QComboBox { background: white; border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px; }
            QPushButton { background: #2563eb; color: white; border: 0; border-radius: 7px; padding: 10px 18px; font-weight: bold; }
            QPushButton:hover { background: #1d4ed8; }
            QProgressBar { background: #e2e8f0; border: 0; border-radius: 6px; height: 12px; text-align: center; }
            QProgressBar::chunk { background: #22c55e; border-radius: 6px; }
            QLabel { padding: 3px; }
        """)
        tabs = QTabWidget(); tabs.addTab(self.translation_page(), "CAD 翻译（调用 AI）"); tabs.addTab(self.replace_page(), "CAD 文本替换（不调用 AI）"); tabs.addTab(self.excel_batch_page(), "CAD 按 Excel 批量自动更换"); self.setCentralWidget(tabs)
    def busy(self):
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.warning(self, "任务正在运行", "请等待当前 CAD 任务完成后，再启动其他模式。"); return True
        return False
    def translation_page(self):
        page = QWidget(); form = QFormLayout(); self.tr_provider = QComboBox(); self.tr_provider.addItems(["DeepSeek", "ChatGPT"]); self.tr_provider.currentTextChanged.connect(self.provider_changed)
        self.tr_language = QComboBox(); self.tr_language.setEditable(True); self.tr_language.addItems(["English", "Japanese", "Korean", "French", "German"]); self.tr_key = QLineEdit(); self.tr_key.setEchoMode(QLineEdit.Password); self.tr_base = QLineEdit("https://api.deepseek.com/v1"); self.tr_model = QLineEdit("deepseek-chat"); self.tr_source = QLineEdit(); self.tr_output = QLineEdit(); self.tr_glossary = QLineEdit(); self.tr_save_excel = QLineEdit()
        form.addRow("AI 服务", self.tr_provider); form.addRow("目标语言", self.tr_language); form.addRow("API Key", self.tr_key); form.addRow("Base URL", self.tr_base); form.addRow("模型名称", self.tr_model); form.addRow("CAD 源文件夹", path_row(self.tr_source, self, True)); form.addRow("英文版输出文件夹", path_row(self.tr_output, self, True)); form.addRow("读取术语表 Excel（可选）", path_row(self.tr_glossary, self)); form.addRow("保存对照表（可选）", path_row(self.tr_save_excel, self, False, True))
        self.tr_preview = QCheckBox("仅生成翻译表，检查后再替换（推荐）"); self.tr_preview.setChecked(True); self.tr_bar = QProgressBar(); self.tr_status = QLabel("翻译页：准备就绪"); start = QPushButton("开始 CAD 翻译"); start.clicked.connect(self.start_translation); layout = QVBoxLayout(); layout.addLayout(form); layout.addWidget(self.tr_preview); layout.addWidget(self.tr_bar); layout.addWidget(self.tr_status); layout.addWidget(start); page.setLayout(layout); return page
    def replace_page(self):
        page = QWidget(); form = QFormLayout(); self.r_source = QLineEdit(); self.r_output = QLineEdit(); self.r_excel = QLineEdit(); form.addRow("CAD 源文件夹", path_row(self.r_source, self, True)); form.addRow("修改后输出文件夹", path_row(self.r_output, self, True)); form.addRow("替换表 Excel", path_row(self.r_excel, self)); hint = QLabel("Excel 第 1 列=原文本，第 2 列=替换文本。此页不调用 AI，只按表格执行替换。"); self.r_bar = QProgressBar(); self.r_status = QLabel("替换页：准备就绪"); start = QPushButton("开始普通文本替换"); start.clicked.connect(self.start_replace); layout = QVBoxLayout(); layout.addLayout(form); layout.addWidget(hint); layout.addWidget(self.r_bar); layout.addWidget(self.r_status); layout.addWidget(start); page.setLayout(layout); return page
    def excel_batch_page(self):
        page = QWidget(); form = QFormLayout(); self.b_source = QLineEdit(); self.b_output = QLineEdit(); self.b_excel = QLineEdit(); form.addRow("CAD 源文件夹", path_row(self.b_source, self, True)); form.addRow("自动更换输出文件夹", path_row(self.b_output, self, True)); form.addRow("Excel 更换表", path_row(self.b_excel, self)); hint = QLabel("独立批处理模式：第 1 列=原文本，第 2 列=新文本。只按 Excel 批量执行，不调用 AI，也不读取其他页签配置。"); self.b_bar = QProgressBar(); self.b_status = QLabel("批量自动更换：准备就绪"); start = QPushButton("开始批量自动更换"); start.clicked.connect(self.start_excel_batch); layout = QVBoxLayout(); layout.addLayout(form); layout.addWidget(hint); layout.addWidget(self.b_bar); layout.addWidget(self.b_status); layout.addWidget(start); page.setLayout(layout); return page
    def provider_changed(self, name): self.tr_base.setText("https://api.openai.com/v1" if name == "ChatGPT" else "https://api.deepseek.com/v1"); self.tr_model.setText("gpt-4o-mini" if name == "ChatGPT" else "deepseek-chat")
    def start_translation(self):
        if self.busy(): return
        if not Path(self.tr_source.text()).is_dir() or not self.tr_output.text() or not self.tr_key.text().strip(): QMessageBox.warning(self, "配置不完整", "请选择 CAD 源文件夹、英文版输出文件夹并填写 API Key"); return
        values = dict(source_dir=Path(self.tr_source.text()), output_dir=Path(self.tr_output.text()), glossary_path=Path(self.tr_glossary.text()) if self.tr_glossary.text() else None, provider="chatgpt" if self.tr_provider.currentText() == "ChatGPT" else "deepseek", target_language=self.tr_language.currentText(), api_key=self.tr_key.text().strip(), base_url=self.tr_base.text().strip(), model=self.tr_model.text().strip(), report_path=Path(self.tr_save_excel.text()) if self.tr_save_excel.text() else None, preview_only=self.tr_preview.isChecked())
        self.worker = Worker(translate_run, values); self.worker.progress.connect(lambda p, s: (self.tr_bar.setValue(p), self.tr_status.setText(s))); self.worker.finished.connect(lambda ok, msg: self.done(ok, msg, self.tr_status)); self.worker.start()
    def start_replace(self):
        if self.busy(): return
        if not Path(self.r_source.text()).is_dir() or not self.r_output.text() or not Path(self.r_excel.text()).is_file(): QMessageBox.warning(self, "配置不完整", "请选择 CAD 源文件夹、修改后输出文件夹和替换表 Excel"); return
        values = dict(source_dir=Path(self.r_source.text()), output_dir=Path(self.r_output.text()), excel_path=Path(self.r_excel.text())); self.worker = Worker(replace_run, values); self.worker.progress.connect(lambda p, s: (self.r_bar.setValue(p), self.r_status.setText(s))); self.worker.finished.connect(lambda ok, msg: self.done(ok, msg, self.r_status)); self.worker.start()
    def start_excel_batch(self):
        if self.busy(): return
        if not Path(self.b_source.text()).is_dir() or not self.b_output.text() or not Path(self.b_excel.text()).is_file(): QMessageBox.warning(self, "配置不完整", "请选择 CAD 源文件夹、自动更换输出文件夹和 Excel 更换表"); return
        values = dict(source_dir=Path(self.b_source.text()), output_dir=Path(self.b_output.text()), excel_path=Path(self.b_excel.text())); self.worker = Worker(replace_run, values); self.worker.progress.connect(lambda p, s: (self.b_bar.setValue(p), self.b_status.setText(s))); self.worker.finished.connect(lambda ok, msg: self.done(ok, msg, self.b_status)); self.worker.start()
    def done(self, success, message, status): status.setText(message); QMessageBox.information(self, "完成" if success else "失败", message)

if __name__ == "__main__":
    app = QApplication(sys.argv); window = Window(); window.show(); sys.exit(app.exec_())
