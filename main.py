import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout, QLineEdit, QProgressBar, QMessageBox, QDialog, QCheckBox, QStackedWidget, QStatusBar
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

# ... 其余代码保持不变 ...

if __name__ == "__main__":
    init_database()  # 启动时自动初始化数据库
    import sys
    from PyQt6.QtWidgets import QApplication
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec()) 