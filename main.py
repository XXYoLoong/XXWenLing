import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout, QLineEdit, QProgressBar, QMessageBox, QDialog, QCheckBox, QStackedWidget, QStatusBar, QComboBox
)
from PyQt6.QtGui import QPixmap, QFont, QDesktopServices
from PyQt6.QtCore import Qt, QUrl
from merger import merge_documents
from format_page import FormatPage
from database.init_db import init_database

ASSETS = os.path.join(os.path.dirname(__file__), 'assets')
BG_PATH = os.path.join(ASSETS, 'bg.jpg')
LOGO_PATH = os.path.join(ASSETS, 'logo.png')
GITHUB_URL = "https://github.com/XXYoLoong"

COPYRIGHT_TEXT = '© 2025 倪家诚 | 协议: <a href="https://www.apache.org/licenses/LICENSE-2.0">Apache-2.0</a> | 联系: 2561889884@qq.com'

GLOBAL_INSTRUCTION = (
    "欢迎使用小小文灵（XXWenLing）！\n\n"
    "本软件包含以下主要功能：\n"
    "1. 文档合并：批量合并Word文档，自动分页，保留格式。\n"
    "2. 文档格式化：批量规范化Word文档排版，支持自定义模板。\n\n"
    "操作流程：\n"
    "- 选择所需功能进入对应界面，按提示操作即可。\n"
    "- 详细说明请在功能界面点击右上角'使用手册'查看。\n\n"
    "注意事项：\n"
    "- 仅支持.doc/.docx格式\n- 处理前请关闭所有相关文档\n- 如遇异常请联系作者。"
)
MERGE_INSTRUCTION = (
    "文档合并功能说明：\n"
    "1. 点击'选择文档目录'选择需要合成的文档所在文件夹。\n"
    "2. 输入合成后文档的文件名。\n"
    "3. 可选择保存的目标目录（不选则默认保存到文档目录）。\n"
    "4. 点击'开始合成'即可。\n"
    "5. 合成时会自动分页，保留所有内容（包括图片、表格、格式等）。\n"
    "6. 仅支持.doc和.docx文件。\n"
    "7. 所有需要合成的文档必须全部关闭，否则可能导致合成失败。\n"
    "8. 如遇异常请联系管理员。"
)
FORMAT_INSTRUCTION = (
    "文档格式化功能说明：\n"
    "1. 选择需要批量格式化的文档目录。\n"
    "2. 选择或新建格式模板，配置排版规则。\n"
    "3. 可选择是否覆盖原文件或另存为新文件。\n"
    "4. 点击'开始处理'，等待进度完成。\n"
    "5. 处理日志会显示成功与失败的文档。\n"
    "6. 仅支持.docx文件。\n    7. 如遇异常请联系管理员。"
)

class InstructionDialog(QDialog):
    def __init__(self, parent=None, text=None, must_check=True):
        super().__init__(parent)
        self.setWindowTitle("使用说明")
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        layout = QVBoxLayout()
        label = QLabel(text or GLOBAL_INSTRUCTION)
        label.setWordWrap(True)
        label.setFont(QFont("微软雅黑", 12))
        layout.addWidget(label)
        copyright_label = QLabel()
        copyright_label.setTextFormat(Qt.TextFormat.RichText)
        copyright_label.setText(COPYRIGHT_TEXT)
        copyright_label.setOpenExternalLinks(True)
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        copyright_label.setStyleSheet('color:#888;font-size:11px;')
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
        self.adjustSize()
    def check_ok(self):
        self.ok_btn.setEnabled(self.cb1.isChecked() or self.cb2.isChecked())

class RegisterDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("用户注册/登录")
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        layout = QVBoxLayout()
        layout.addWidget(QLabel("请输入用户名："))
        self.username_edit = QLineEdit()
        layout.addWidget(self.username_edit)
        layout.addWidget(QLabel("请输入邮箱："))
        self.email_edit = QLineEdit()
        layout.addWidget(self.email_edit)
        layout.addWidget(QLabel("请输入密码："))
        self.pwd_edit = QLineEdit()
        self.pwd_edit.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.pwd_edit)
        self.ok_btn = QPushButton("注册/登录")
        self.ok_btn.clicked.connect(self.do_register)
        layout.addWidget(self.ok_btn)
        self.status = QLabel("")
        self.status.setStyleSheet("color:#ff9800; font-size:13px;")
        layout.addWidget(self.status)
        self.setLayout(layout)
        self.user_id = None
    def do_register(self):
        from database.models import User
        username = self.username_edit.text().strip()
        email = self.email_edit.text().strip()
        pwd = self.pwd_edit.text().strip()
        if not username or not pwd:
            self.status.setText("用户名和密码不能为空！")
            return
        user_model = User()
        user = user_model.get_by_username(username)
        if user:
            # 登录
            if user['password_hash'] == pwd:
                self.user_id = user['id']
                self.accept()
            else:
                self.status.setText("密码错误！")
        else:
            # 注册
            user_id = user_model.create(username, email, pwd)
            self.user_id = user_id
            self.accept()

class MainPage(QWidget):
    def __init__(self, switch_to_merge, switch_to_format, show_manual):
        super().__init__()
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo = QLabel()
        logo.setPixmap(QPixmap(LOGO_PATH).scaled(64, 64, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        title = QLabel("小小文灵（XXWenLing）")
        title.setFont(QFont("微软雅黑", 28, QFont.Weight.Bold))
        title.setStyleSheet("color:#2563eb; letter-spacing:8px;")
        layout.addWidget(logo, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(30)
        merge_btn = QPushButton("文档合并")
        merge_btn.setFixedHeight(54)
        merge_btn.setStyleSheet("border-radius:27px; background:#2563eb; color:white; font-size:22px; padding:10px 40px;")
        merge_btn.clicked.connect(switch_to_merge)
        layout.addWidget(merge_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(18)
        format_btn = QPushButton("文档格式化")
        format_btn.setFixedHeight(54)
        format_btn.setStyleSheet("border-radius:27px; background:#2563eb; color:white; font-size:22px; padding:10px 40px;")
        format_btn.clicked.connect(switch_to_format)
        layout.addWidget(format_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(18)
        wait_btn = QPushButton("敬请期待")
        wait_btn.setFixedHeight(54)
        wait_btn.setStyleSheet("border-radius:27px; background:#ff9800; color:white; font-size:22px; padding:10px 40px;")
        wait_btn.clicked.connect(self.show_wait)
        layout.addWidget(wait_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(30)
        manual_btn = QPushButton("使用手册")
        manual_btn.setStyleSheet("border-radius:14px; background:#43a047; color:white; font-size:15px; padding:6px 18px;")
        manual_btn.clicked.connect(show_manual)
        layout.addWidget(manual_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch()
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
        dlg.setLayout(v)
        dlg.exec()

class MergePage(QWidget):
    def __init__(self, back_to_main, show_manual, show_instruction_once, user_id):
        super().__init__()
        self._shown = False
        self.show_instruction_once = show_instruction_once
        self.user_id = user_id
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
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
        manual_btn = QPushButton("使用手册")
        manual_btn.setStyleSheet("border-radius:14px; background:#43a047; color:white; font-size:15px; padding:6px 18px;")
        manual_btn.clicked.connect(show_manual)
        top_layout.addWidget(manual_btn)
        layout.addLayout(top_layout)
        layout.addSpacing(18)
        self.dir_edit = QLineEdit()
        self.dir_edit.setPlaceholderText("请选择文档目录")
        self.dir_edit.setReadOnly(True)
        dir_btn = QPushButton("选择目录")
        dir_btn.clicked.connect(self.choose_dir)
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(self.dir_edit)
        dir_layout.addWidget(dir_btn)
        layout.addLayout(dir_layout)
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("请输入合成后文件名（不含扩展名）")
        layout.addWidget(self.name_edit)
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
        self.merge_btn = QPushButton("开始合成")
        self.merge_btn.clicked.connect(self.on_merge)
        self.merge_btn.setFixedHeight(40)
        self.merge_btn.setStyleSheet("border-radius:20px; background:#2563eb; color:white; font-size:18px;")
        layout.addWidget(self.merge_btn)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setFixedHeight(18)
        self.progress.setStyleSheet("QProgressBar{border-radius:9px; background:#eee;} QProgressBar::chunk{background:#2563eb; border-radius:9px;}")
        layout.addWidget(self.progress)
        self.status = QLabel("")
        self.status.setStyleSheet("color:#ff9800; font-size:14px;")
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status)
        layout.addStretch()
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
            out_path = merge_documents(doc_dir, out_name, save_dir, self.progress, self.status, user_id=self.user_id)
            self.status.setText(f"合成完成：{out_path}")
            QMessageBox.information(self, "成功", f"文档已成功合成到：\n{out_path}")
        except Exception as e:
            self.status.setText("合成失败！")
            QMessageBox.critical(self, "错误", f"合成过程中出现错误：\n{str(e)}")
        self.progress.setValue(0)
    def showEvent(self, event):
        super().showEvent(event)
        if not self._shown:
            self._shown = True
            self.show_instruction_once()

class MainWindow(QMainWindow):
    def __init__(self, user_id):
        super().__init__()
        self.user_id = user_id
        self.setWindowTitle("小小文灵（XXWenLing）")
        self.setFixedSize(800, 600)
        self.setObjectName("MainWindow")
        self.setStyleSheet(self.qss())
        self.stack = QStackedWidget()
        self._merge_instruction_shown = False
        self._format_instruction_shown = False
        self.main_page = MainPage(self.switch_to_merge, self.switch_to_format, self.show_global_manual)
        self.merge_page = MergePage(self.switch_to_main, self.show_merge_manual, self.show_merge_instruction_once, user_id=self.user_id)
        self.format_page = FormatPage(self.switch_to_main, self.show_format_manual, self.show_format_instruction_once, user_id=self.user_id)
        self.stack.addWidget(self.main_page)
        self.stack.addWidget(self.merge_page)
        self.stack.addWidget(self.format_page)
        self.setCentralWidget(self.stack)
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        copyright_label = QLabel()
        copyright_label.setTextFormat(Qt.TextFormat.RichText)
        copyright_label.setText(COPYRIGHT_TEXT)
        copyright_label.setOpenExternalLinks(True)
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        copyright_label.setStyleSheet('color:#888;font-size:12px;')
        self.status_bar.addWidget(copyright_label)
        self.show_global_instruction()
    def show_global_instruction(self):
        dlg = InstructionDialog(self, text=GLOBAL_INSTRUCTION, must_check=True)
        dlg.exec()
    def show_global_manual(self):
        dlg = InstructionDialog(self, text=GLOBAL_INSTRUCTION, must_check=False)
        dlg.exec()
    def show_merge_manual(self):
        dlg = InstructionDialog(self, text=MERGE_INSTRUCTION, must_check=False)
        dlg.exec()
    def show_format_manual(self):
        dlg = InstructionDialog(self, text=FORMAT_INSTRUCTION, must_check=False)
        dlg.exec()
    def show_merge_instruction_once(self):
        if not self._merge_instruction_shown:
            self._merge_instruction_shown = True
            dlg = InstructionDialog(self, text=MERGE_INSTRUCTION, must_check=True)
            dlg.exec()
    def show_format_instruction_once(self):
        if not self._format_instruction_shown:
            self._format_instruction_shown = True
            dlg = InstructionDialog(self, text=FORMAT_INSTRUCTION, must_check=True)
            dlg.exec()
    def switch_to_merge(self):
        self.stack.setCurrentWidget(self.merge_page)
    def switch_to_format(self):
        self.stack.setCurrentWidget(self.format_page)
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

# 入口
if __name__ == "__main__":
    init_database()  # 启动时自动初始化数据库
    import sys
    from PyQt6.QtWidgets import QApplication
    app = QApplication(sys.argv)
    # 弹出注册/登录对话框
    reg = RegisterDialog()
    if reg.exec() == QDialog.DialogCode.Accepted:
        user_id = reg.user_id
        window = MainWindow(user_id)
        window.show()
        sys.exit(app.exec()) 