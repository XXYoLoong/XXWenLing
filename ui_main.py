import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout, QLineEdit, QProgressBar, QMessageBox, QDialog, QCheckBox, QStackedWidget
)
from PyQt6.QtGui import QPixmap, QFont, QDesktopServices
from PyQt6.QtCore import Qt, QUrl
from merger import merge_documents

ASSETS = os.path.join(os.path.dirname(__file__), 'assets')
BG_PATH = os.path.join(ASSETS, 'bg.jpg')
LOGO_PATH = os.path.join(ASSETS, 'logo.png')
GITHUB_URL = "https://github.com/XXYoLoong"

APACHE_LICENSE = (
    "Copyright 2025 Jiacheng Ni\n\n"
    "Licensed under the Apache License, Version 2.0 (the 'License');\n"
    "you may not use this file except in compliance with the License.\n"
    "You may obtain a copy of the License at\n\n"
    "    http://www.apache.org/licenses/LICENSE-2.0\n\n"
    "Unless required by applicable law or agreed to in writing, software\n"
    "distributed under the License is distributed on an 'AS IS' BASIS,\n"
    "WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.\n"
    "See the License for the specific language governing permissions and\n"
    "limitations under the License.\n\n"
    "@https://github.com/XXYoLoong"
)

INSTRUCTION_TEXT = (
    "使用说明：\n"
    "1. 先点击'选择文档目录'选择需要合成的文档所在文件夹。\n"
    "2. 再点击'命名合并文档'输入合成后文档的文件名。\n"
    "3. 可选择保存的目标目录（不选则默认保存到文档目录）。\n"
    "4. 最后点击'开始合成'即可。\n"
    "5. 合成时会自动分页，保留所有内容（包括图片、表格、格式等）。\n"
    "6. 仅支持.doc和.docx文件。\n"
    "7. 所有需要合成的文档必须全部关闭，否则可能导致合成失败。\n"
    "8. 如遇异常请联系管理员。"
)

class InstructionDialog(QDialog):
    def __init__(self, parent=None, must_check=True):
        super().__init__(parent)
        self.setWindowTitle("使用说明")
        self.setFixedSize(520, 480)
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        layout = QVBoxLayout()
        label = QLabel(INSTRUCTION_TEXT)
        label.setWordWrap(True)
        label.setFont(QFont("微软雅黑", 12))
        layout.addWidget(label)
        copyright_label = QLabel(APACHE_LICENSE)
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        copyright_label.setStyleSheet("color:#888;font-size:11px;")
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

class MainPage(QWidget):
    def __init__(self, switch_to_merge):
        super().__init__()
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # 顶部logo和标题
        logo = QLabel()
        logo.setPixmap(QPixmap(LOGO_PATH).scaled(64, 64, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        title = QLabel("小小文灵（XXWenLing）")
        title.setFont(QFont("微软雅黑", 28, QFont.Weight.Bold))
        title.setStyleSheet("color:#2563eb; letter-spacing:8px;")
        layout.addWidget(logo, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(30)
        # 文档合并按钮
        merge_btn = QPushButton("文档合并")
        merge_btn.setFixedHeight(54)
        merge_btn.setStyleSheet("border-radius:27px; background:#2563eb; color:white; font-size:22px; padding:10px 40px;")
        merge_btn.clicked.connect(switch_to_merge)
        layout.addWidget(merge_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(18)
        # 敬请期待按钮
        wait_btn = QPushButton("敬请期待")
        wait_btn.setFixedHeight(54)
        wait_btn.setStyleSheet("border-radius:27px; background:#ff9800; color:white; font-size:22px; padding:10px 40px;")
        wait_btn.clicked.connect(self.show_wait)
        layout.addWidget(wait_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(30)
        # 使用手册按钮
        manual_btn = QPushButton("使用手册")
        manual_btn.setStyleSheet("border-radius:14px; background:#43a047; color:white; font-size:15px; padding:6px 18px;")
        manual_btn.clicked.connect(self.show_manual)
        layout.addWidget(manual_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch()
        # 版权信息
        copyright_label = QLabel(APACHE_LICENSE)
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        copyright_label.setStyleSheet("color:#888;font-size:11px;")
        layout.addWidget(copyright_label)
        self.setLayout(layout)
    def show_wait(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("敬请期待")
        dlg.setFixedSize(400, 220)
        v = QVBoxLayout()
        label = QLabel("暂时没有新的功能了呢，想要看看小小游龙还做了什么有趣的东西吗？")
        label.setWordWrap(True)
        label.setFont(QFont("微软雅黑", 13))
        v.addWidget(label)
        btn = QPushButton("去GitHub看看")
        btn.setStyleSheet("border-radius:14px; background:#2563eb; color:white; font-size:15px; padding:6px 18px;")
        btn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(GITHUB_URL)))
        v.addWidget(btn, alignment=Qt.AlignmentFlag.AlignCenter)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(dlg.accept)
        v.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        # 版权
        copyright_label = QLabel(APACHE_LICENSE)
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        copyright_label.setStyleSheet("color:#888;font-size:11px;")
        v.addWidget(copyright_label)
        dlg.setLayout(v)
        dlg.exec()
    def show_manual(self):
        dlg = InstructionDialog(self, must_check=False)
        dlg.exec()

class MergePage(QWidget):
    def __init__(self, back_to_main):
        super().__init__()
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        # 顶部logo和返回
        top_layout = QHBoxLayout()
        logo = QLabel()
        logo.setPixmap(QPixmap(LOGO_PATH).scaled(40, 40, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        top_layout.addWidget(logo)
        title = QLabel("文档合并")
        title.setFont(QFont("微软雅黑", 20, QFont.Weight.Bold))
        title.setStyleSheet("color:#2563eb; letter-spacing:4px;")
        top_layout.addWidget(title)
        top_layout.addStretch()
        back_btn = QPushButton("返回主页")
        back_btn.setStyleSheet("border-radius:14px; background:#43a047; color:white; font-size:15px; padding:6px 18px;")
        back_btn.clicked.connect(back_to_main)
        top_layout.addWidget(back_btn)
        layout.addLayout(top_layout)
        layout.addSpacing(18)
        # 目录选择
        self.dir_edit = QLineEdit()
        self.dir_edit.setPlaceholderText("请选择文档目录")
        self.dir_edit.setReadOnly(True)
        dir_btn = QPushButton("选择目录")
        dir_btn.clicked.connect(self.choose_dir)
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(self.dir_edit)
        dir_layout.addWidget(dir_btn)
        layout.addLayout(dir_layout)
        # 文件名输入
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("请输入合成后文件名（不含扩展名）")
        layout.addWidget(self.name_edit)
        # 保存目录
        self.save_edit = QLineEdit()
        self.save_edit.setPlaceholderText("请选择保存目标目录（可选）")
        self.save_edit.setReadOnly(True)
        save_btn = QPushButton("选择保存目录")
        save_btn.clicked.connect(self.choose_save_dir)
        save_layout = QHBoxLayout()
        save_layout.addWidget(self.save_edit)
        save_layout.addWidget(save_btn)
        layout.addLayout(save_layout)
        layout.addSpacing(18)
        # 合成按钮
        self.merge_btn = QPushButton("开始合成")
        self.merge_btn.clicked.connect(self.on_merge)
        self.merge_btn.setFixedHeight(40)
        self.merge_btn.setStyleSheet("border-radius:20px; background:#2563eb; color:white; font-size:18px;")
        layout.addWidget(self.merge_btn)
        # 进度条
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setFixedHeight(18)
        self.progress.setStyleSheet("QProgressBar{border-radius:9px; background:#eee;} QProgressBar::chunk{background:#2563eb; border-radius:9px;}")
        layout.addWidget(self.progress)
        # 状态提示
        self.status = QLabel("")
        self.status.setStyleSheet("color:#ff9800; font-size:14px;")
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status)
        layout.addStretch()
        # 版权
        copyright_label = QLabel(APACHE_LICENSE)
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        copyright_label.setStyleSheet("color:#888;font-size:11px;")
        layout.addWidget(copyright_label)
        self.setLayout(layout)
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

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("小小文灵（XXWenLing）")
        self.setFixedSize(800, 600)
        self.setObjectName("MainWindow")
        self.setStyleSheet(self.qss())
        self.stack = QStackedWidget()
        self.main_page = MainPage(self.switch_to_merge)
        self.merge_page = MergePage(self.switch_to_main)
        self.stack.addWidget(self.main_page)
        self.stack.addWidget(self.merge_page)
        self.setCentralWidget(self.stack)
        self.show_instruction_on_startup()
    def show_instruction_on_startup(self):
        dlg = InstructionDialog(self, must_check=True)
        dlg.exec()
    def switch_to_merge(self):
        self.stack.setCurrentWidget(self.merge_page)
    def switch_to_main(self):
        self.stack.setCurrentWidget(self.main_page)
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

if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec()) 