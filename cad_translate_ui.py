"""CAD 翻译图形界面。运行：python cad_translate_ui.py"""
from __future__ import annotations

import sys
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QFileDialog, QComboBox, QFormLayout, QHBoxLayout, QLabel, QLineEdit, QMainWindow, QMessageBox, QProgressBar, QPushButton, QVBoxLayout, QWidget
from cad_translate_deepseek import run


class Worker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)
    def __init__(self, values):
        super().__init__(); self.values = values
    def run(self):
        try:
            def report(index, total, name): self.progress.emit(int(index * 100 / max(total, 1)), f"正在处理 {index}/{total}：{name}")
            result = run(progress=report, **self.values)
            self.finished.emit(True, f"完成：处理 {result[0]} 个文件，修改 {result[1]} 个文字对象，失败 {result[2]} 个。")
        except Exception as exc: self.finished.emit(False, str(exc))


class Window(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle("CAD 图纸智能翻译"); self.resize(720, 420); self.worker = None
        self.provider = QComboBox(); self.provider.addItems(["DeepSeek", "ChatGPT"])
        self.language = QComboBox(); self.language.setEditable(True); self.language.addItems(["English", "Japanese", "Korean", "French", "German"])
        self.key = QLineEdit(); self.key.setEchoMode(QLineEdit.Password); self.key.setPlaceholderText("只在本次运行使用")
        self.base = QLineEdit("https://api.deepseek.com"); self.model = QLineEdit("deepseek-chat")
        self.source = QLineEdit(); self.output = QLineEdit(); self.glossary = QLineEdit()
        form = QFormLayout(); form.addRow("模型供应商", self.provider); form.addRow("目标语言", self.language); form.addRow("API Key", self.key); form.addRow("Base URL", self.base); form.addRow("模型名称", self.model)
        form.addRow("CAD 源文件夹", self.row(self.source, True)); form.addRow("输出文件夹", self.row(self.output, True)); form.addRow("术语表 Excel（可选）", self.row(self.glossary, False))
        group = QWidget(); group.setLayout(form); self.bar = QProgressBar(); self.status = QLabel("准备就绪"); self.start = QPushButton("开始翻译"); self.start.clicked.connect(self.start_job)
        layout = QVBoxLayout(); layout.addWidget(group); layout.addWidget(self.bar); layout.addWidget(self.status); layout.addWidget(self.start); root = QWidget(); root.setLayout(layout); self.setCentralWidget(root)
        self.provider.currentTextChanged.connect(self.provider_changed)
    def row(self, edit, directory):
        button = QPushButton("选择")
        button.clicked.connect(lambda: edit.setText(QFileDialog.getExistingDirectory(self) if directory else QFileDialog.getOpenFileName(self, "选择 Excel", "", "Excel (*.xlsx)")[0]))
        layout = QHBoxLayout(); layout.addWidget(edit); layout.addWidget(button); widget = QWidget(); widget.setLayout(layout); return widget
    def provider_changed(self, name):
        self.base.setText("https://api.openai.com/v1" if name == "ChatGPT" else "https://api.deepseek.com"); self.model.setText("gpt-4o-mini" if name == "ChatGPT" else "deepseek-chat")
    def start_job(self):
        if not Path(self.source.text()).is_dir() or not self.output.text() or not self.key.text().strip():
            QMessageBox.warning(self, "配置不完整", "请选择 CAD 源文件夹、输出文件夹并填写 API Key"); return
        values = dict(source_dir=Path(self.source.text()), output_dir=Path(self.output.text()), glossary_path=Path(self.glossary.text()) if self.glossary.text() else None, provider="chatgpt" if self.provider.currentText() == "ChatGPT" else "deepseek", target_language=self.language.currentText(), api_key=self.key.text().strip(), base_url=self.base.text().strip(), model=self.model.text().strip())
        self.start.setEnabled(False); self.worker = Worker(values); self.worker.progress.connect(lambda p, s: (self.bar.setValue(p), self.status.setText(s))); self.worker.finished.connect(self.done); self.worker.start()
    def done(self, success, message):
        self.start.setEnabled(True); self.status.setText(message); QMessageBox.information(self, "完成" if success else "失败", message)


if __name__ == "__main__":
    app = QApplication(sys.argv); window = Window(); window.show(); sys.exit(app.exec_())
