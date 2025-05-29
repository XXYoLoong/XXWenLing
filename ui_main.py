import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout, QLineEdit, QProgressBar, QMessageBox, QDialog, QCheckBox
)
from PyQt6.QtGui import QPixmap, QFont
from PyQt6.QtCore import Qt
from merger import merge_documents

ASSETS = os.path.join(os.path.dirname(__file__), 'assets')
FONT_PATH = os.path.join(ASSETS, 'font.ttf')
BG_PATH = os.path.join(ASSETS, 'bg.jpg')
LOGO_PATH = os.path.join(ASSETS, 'logo.png')
COPYRIGHT = "© 倪家诚 2025/5/29"

INSTRUCTION_TEXT = (
    "使用说明：\n"
    "1. 先点击‘选择文档目录’选择需要合成的文档所在文件夹。\n"
    "2. 再点击‘命名合并文档’输入合成后文档的文件名。\n"
    "3. 可选择保存的目标目录（不选则默认保存到文档目录）。\n"
    "4. 最后点击‘开始合成’即可。\n"
    "5. 合成时会自动分页，保留所有内容（包括图片、表格、格式等）。\n"
    "6. 仅支持.doc和.docx文件。\n"
    "7. 所有需要合成的文档必须全部关闭，否则可能导致合成失败。\n"
    "8. 如遇异常请联系管理员。"
)

class InstructionDialog(QDialog):
    def __init__(self, parent=None, must_check=True):
        super().__init__(parent)
        self.setWindowTitle("使用说明")
        self.setFixedSize(480, 420)
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        layout = QVBoxLayout()
        label = QLabel(INSTRUCTION_TEXT)
        label.setWordWrap(True)
        label.setFont(QFont("微软雅黑", 12))
        layout.addWidget(label)
        copyright_label = QLabel(COPYRIGHT)
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        copyright_label.setStyleSheet("color:#888;font-size:12px;")
        layout.addWidget(copyright_label)
        if must_check:
            self.cb1 = QCheckBox("知道啦，不用再说了")
            self.cb2 = QCheckBox("下次使用我还要看看呢")
            self.cb1.stateChanged.connect(lambda: self.cb2.setChecked(False))
            self.cb2.stateChanged.connect(lambda: self.cb1.setChecked(False))
            layout.addWidget(self.cb1)
            layout.addWidget(self.cb2)
            self.ok_btn = QPushButton("确定")
            self.ok_btn.setEnabled(False)
            self.cb1.stateChanged.connect(self.check_ok)
            self.cb2.stateChanged.connect(self.check_ok)
            self.ok_btn.clicked.connect(self.accept)
            layout.addWidget(self.ok_btn)
            self.setWindowFlag(Qt.WindowType.WindowCloseButtonHint, False)
        else:
            close_btn = QPushButton("关闭")
            close_btn.clicked.connect(self.accept)
            layout.addWidget(close_btn)
        self.setLayout(layout)
    def check_ok(self):
        self.ok_btn.setEnabled(self.cb1.isChecked() or self.cb2.isChecked())

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("文档合璧")
        self.setFixedSize(700, 500)
        self.setObjectName("MainWindow")
        self.setStyleSheet(self.qss())
        self.init_ui()
        self.show_instruction_on_startup()

    def show_instruction_on_startup(self):
        dlg = InstructionDialog(self, must_check=True)
        dlg.exec()

    def show_manual(self):
        dlg = InstructionDialog(self, must_check=False)
        dlg.exec()

    def init_ui(self):
        # 背景
        bg_label = QLabel(self)
        bg_label.setPixmap(QPixmap(BG_PATH).scaled(self.size(), Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation))
        bg_label.setGeometry(0, 0, self.width(), self.height())
        bg_label.lower()

        # 顶部logo和标题
        logo = QLabel()
        logo.setPixmap(QPixmap(LOGO_PATH).scaled(48, 48, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        title = QLabel("文档合璧")
        title.setFont(QFont("微软雅黑", 22, QFont.Weight.Bold))
        title.setStyleSheet("color:#2563eb; letter-spacing:6px;")
        top_layout = QHBoxLayout()
        top_layout.addWidget(logo)
        top_layout.addWidget(title)
        top_layout.addStretch()
        # 使用手册按钮
        manual_btn = QPushButton("使用手册")
        manual_btn.clicked.connect(self.show_manual)
        manual_btn.setStyleSheet("border-radius:14px; background:#ff9800; color:white; font-size:15px; padding:6px 18px;")
        top_layout.addWidget(manual_btn)

        # 目录选择
        self.dir_edit = QLineEdit()
        self.dir_edit.setPlaceholderText("请选择文档目录")
        self.dir_edit.setReadOnly(True)
        dir_btn = QPushButton("选择目录")
        dir_btn.clicked.connect(self.choose_dir)
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(self.dir_edit)
        dir_layout.addWidget(dir_btn)

        # 文件名输入
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("请输入合成后文件名（不含扩展名）")

        # 保存目录
        self.save_edit = QLineEdit()
        self.save_edit.setPlaceholderText("请选择保存目标目录（可选）")
        self.save_edit.setReadOnly(True)
        save_btn = QPushButton("选择保存目录")
        save_btn.clicked.connect(self.choose_save_dir)
        save_layout = QHBoxLayout()
        save_layout.addWidget(self.save_edit)
        save_layout.addWidget(save_btn)

        # 合成按钮
        self.merge_btn = QPushButton("开始合成")
        self.merge_btn.clicked.connect(self.on_merge)
        self.merge_btn.setFixedHeight(40)
        self.merge_btn.setStyleSheet("border-radius:20px; background:#2563eb; color:white; font-size:18px;")

        # 进度条
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setFixedHeight(18)
        self.progress.setStyleSheet("QProgressBar{border-radius:9px; background:#eee;} QProgressBar::chunk{background:#2563eb; border-radius:9px;}")

        # 状态提示
        self.status = QLabel("")
        self.status.setStyleSheet("color:#ff9800; font-size:14px;")
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # 主布局
        main_layout = QVBoxLayout()
        main_layout.addLayout(top_layout)
        main_layout.addSpacing(18)
        main_layout.addLayout(dir_layout)
        main_layout.addWidget(self.name_edit)
        main_layout.addLayout(save_layout)
        main_layout.addSpacing(18)
        main_layout.addWidget(self.merge_btn)
        main_layout.addWidget(self.progress)
        main_layout.addWidget(self.status)
        main_layout.addStretch()

        # 容器
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def choose_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择文档目录")
        if d:
            self.dir_edit.setText(d)

    def choose_save_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择保存目标目录")
        if d:
            self.save_edit.setText(d)

    def on_merge(self):
        doc_dir = self.dir_edit.text().strip()
        out_name = self.name_edit.text().strip()
        save_dir = self.save_edit.text().strip() or doc_dir
        if not doc_dir or not out_name:
            self.status.setText("请先选择文档目录并输入文件名！")
            return
        self.status.setText("正在合成...")
        self.progress.setValue(0)
        try:
            out_path = merge_documents(doc_dir, out_name, save_dir, self.progress, self.status)
            self.status.setText(f"合成完成：{out_path}")
            QMessageBox.information(self, "成功", f"文档已成功合成到：\n{out_path}")
        except Exception as e:
            self.status.setText("合成失败！")
            QMessageBox.critical(self, "错误", f"合成过程中出现错误：\n{str(e)}")
        self.progress.setValue(0)

    def qss(self):
        return f"""
        QMainWindow#MainWindow {{
            background: #fff;
        }}
        QPushButton {{
            border-radius: 20px;
            background: #2563eb;
            color: white;
            font-size: 16px;
            padding: 10px 24px;
        }}
        QPushButton:hover {{
            background: #ff9800;
        }}
        QLineEdit {{
            border: 2px solid #2563eb;
            border-radius: 12px;
            padding: 8px 12px;
            font-size: 15px;
        }}
        QLabel {{
            font-family: '微软雅黑', 'SimHei', 'Arial';
        }}
        """ 