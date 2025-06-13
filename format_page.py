from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QLineEdit, QComboBox, QProgressBar, QMessageBox, QDialog,
    QFormLayout, QSpinBox, QDoubleSpinBox, QTextEdit, QFileDialog,
    QGroupBox, QRadioButton, QCheckBox, QStackedWidget, QListWidget, QListWidgetItem
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QFontDatabase
from formatter import DocumentFormatter
from database.models import FormatTemplate
import json
import os

# 中文字号与磅值映射
CHINESE_FONT_SIZES = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5
}
ALL_FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72] + list(CHINESE_FONT_SIZES.keys())

class TemplateDialog(QDialog):
    def __init__(self, parent=None, template=None):
        super().__init__(parent)
        self.setWindowTitle("格式模板配置")
        self.setFixedSize(600, 800)
        self.template = template or {
            'name': '',
            'rules': {
                'title': {'type': 'position', 'position': 0},
                'subtitle': {'type': 'position', 'position': 1},
                'body': {'type': 'default'}
            },
            'formats': {
                'title': {'font': '方正小标宋简体', 'size': 16, 'alignment': 'center'},
                'subtitle': {'font': '仿宋', 'size': 14, 'alignment': 'right'},
                'body': {'font': '仿宋', 'size': 12, 'alignment': 'left', 'indent': 2, 'line_spacing': 28}
            }
        }
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 模板名称
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("模板名称："))
        self.name_edit = QLineEdit(self.template['name'])
        name_layout.addWidget(self.name_edit)
        layout.addLayout(name_layout)
        
        # 规则配置
        rules_group = QGroupBox("段落识别规则")
        rules_layout = QFormLayout()
        
        # 标题规则
        self.title_rule_type = QComboBox()
        self.title_rule_type.addItems(['位置', '关键词', '正则表达式', '长度'])
        self.title_rule_type.currentTextChanged.connect(lambda: self.update_rule_ui('title'))
        rules_layout.addRow("标题规则类型：", self.title_rule_type)
        
        self.title_rule_stack = QStackedWidget()
        # 位置
        self.title_position = QSpinBox()
        self.title_position.setRange(0, 100)
        # 关键词
        self.title_keyword = QLineEdit()
        # 正则
        self.title_regex = QLineEdit()
        # 长度
        self.title_length = QSpinBox()
        self.title_length.setRange(1, 1000)
        self.title_rule_stack.addWidget(self.title_position)
        self.title_rule_stack.addWidget(self.title_keyword)
        self.title_rule_stack.addWidget(self.title_regex)
        self.title_rule_stack.addWidget(self.title_length)
        rules_layout.addRow("", self.title_rule_stack)
        
        # 副标题规则
        self.subtitle_rule_type = QComboBox()
        self.subtitle_rule_type.addItems(['位置', '关键词', '正则表达式', '长度'])
        self.subtitle_rule_type.currentTextChanged.connect(lambda: self.update_rule_ui('subtitle'))
        rules_layout.addRow("副标题规则类型：", self.subtitle_rule_type)
        
        self.subtitle_rule_stack = QStackedWidget()
        self.subtitle_position = QSpinBox()
        self.subtitle_position.setRange(0, 100)
        self.subtitle_keyword = QLineEdit()
        self.subtitle_regex = QLineEdit()
        self.subtitle_length = QSpinBox()
        self.subtitle_length.setRange(1, 1000)
        self.subtitle_rule_stack.addWidget(self.subtitle_position)
        self.subtitle_rule_stack.addWidget(self.subtitle_keyword)
        self.subtitle_rule_stack.addWidget(self.subtitle_regex)
        self.subtitle_rule_stack.addWidget(self.subtitle_length)
        rules_layout.addRow("", self.subtitle_rule_stack)
        
        rules_group.setLayout(rules_layout)
        layout.addWidget(rules_group)
        
        # 格式配置
        formats_group = QGroupBox("格式设置")
        formats_layout = QFormLayout()
        
        # 字体下拉框（自动读取系统字体，带搜索）
        all_fonts = QFontDatabase.families()
        # 标题格式
        formats_layout.addRow("标题格式：", QLabel(""))
        self.title_font = QComboBox()
        self.title_font.setEditable(True)
        self.title_font.addItems(all_fonts)
        self.title_font.setCurrentText(self.template['formats']['title'].get('font', '方正小标宋简体'))
        formats_layout.addRow("字体：", self.title_font)
        self.title_size = QComboBox()
        for sz in ALL_FONT_SIZES:
            self.title_size.addItem(str(sz))
        self.title_size.setEditable(True)
        self.title_size.setCurrentText(str(self.template['formats']['title'].get('size', 16)))
        formats_layout.addRow("字号：", self.title_size)
        self.title_align = QComboBox()
        self.title_align.addItems(['左对齐', '居中', '右对齐', '两端对齐'])
        self.title_align.setCurrentText(self.template['formats']['title'].get('alignment', 'center'))
        formats_layout.addRow("对齐方式：", self.title_align)
        
        # 副标题格式
        formats_layout.addRow("副标题格式：", QLabel(""))
        self.subtitle_font = QComboBox()
        self.subtitle_font.setEditable(True)
        self.subtitle_font.addItems(all_fonts)
        self.subtitle_font.setCurrentText(self.template['formats']['subtitle'].get('font', '仿宋'))
        formats_layout.addRow("字体：", self.subtitle_font)
        self.subtitle_size = QComboBox()
        for sz in ALL_FONT_SIZES:
            self.subtitle_size.addItem(str(sz))
        self.subtitle_size.setEditable(True)
        self.subtitle_size.setCurrentText(str(self.template['formats']['subtitle'].get('size', 14)))
        formats_layout.addRow("字号：", self.subtitle_size)
        self.subtitle_align = QComboBox()
        self.subtitle_align.addItems(['左对齐', '居中', '右对齐', '两端对齐'])
        self.subtitle_align.setCurrentText(self.template['formats']['subtitle'].get('alignment', 'right'))
        formats_layout.addRow("对齐方式：", self.subtitle_align)
        
        # 正文格式
        formats_layout.addRow("正文格式：", QLabel(""))
        self.body_font = QComboBox()
        self.body_font.setEditable(True)
        self.body_font.addItems(all_fonts)
        self.body_font.setCurrentText(self.template['formats']['body'].get('font', '仿宋'))
        formats_layout.addRow("字体：", self.body_font)
        self.body_size = QComboBox()
        for sz in ALL_FONT_SIZES:
            self.body_size.addItem(str(sz))
        self.body_size.setEditable(True)
        self.body_size.setCurrentText(str(self.template['formats']['body'].get('size', 12)))
        formats_layout.addRow("字号：", self.body_size)
        self.body_align = QComboBox()
        self.body_align.addItems(['左对齐', '居中', '右对齐', '两端对齐'])
        self.body_align.setCurrentText(self.template['formats']['body'].get('alignment', 'left'))
        formats_layout.addRow("对齐方式：", self.body_align)
        self.body_indent = QDoubleSpinBox()
        self.body_indent.setRange(0, 10)
        self.body_indent.setValue(self.template['formats']['body'].get('indent', 2))
        formats_layout.addRow("首行缩进(字符)：", self.body_indent)
        self.body_spacing = QSpinBox()
        self.body_spacing.setRange(1, 100)
        self.body_spacing.setValue(self.template['formats']['body'].get('line_spacing', 28))
        formats_layout.addRow("行距(磅)：", self.body_spacing)
        
        formats_group.setLayout(formats_layout)
        layout.addWidget(formats_group)
        
        # 按钮
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("保存模板")
        save_btn.clicked.connect(self.save_template)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
        
        # 初始化规则类型和内容
        self.init_rule_ui()
        
    def init_rule_ui(self):
        # 标题
        t_rule = self.template['rules'].get('title', {'type': 'position', 'position': 0})
        t_type = t_rule.get('type', 'position')
        t_map = {'position': 0, 'keyword': 1, 'regex': 2, 'length': 3}
        self.title_rule_type.setCurrentIndex(t_map.get(t_type, 0))
        if t_type == 'position':
            self.title_position.setValue(t_rule.get('position', 0))
        elif t_type == 'keyword':
            self.title_keyword.setText(t_rule.get('keyword', ''))
        elif t_type == 'regex':
            self.title_regex.setText(t_rule.get('pattern', ''))
        elif t_type == 'length':
            self.title_length.setValue(t_rule.get('max_length', 20))
        self.title_rule_stack.setCurrentIndex(t_map.get(t_type, 0))
        # 副标题
        s_rule = self.template['rules'].get('subtitle', {'type': 'position', 'position': 1})
        s_type = s_rule.get('type', 'position')
        self.subtitle_rule_type.setCurrentIndex(t_map.get(s_type, 0))
        if s_type == 'position':
            self.subtitle_position.setValue(s_rule.get('position', 1))
        elif s_type == 'keyword':
            self.subtitle_keyword.setText(s_rule.get('keyword', ''))
        elif s_type == 'regex':
            self.subtitle_regex.setText(s_rule.get('pattern', ''))
        elif s_type == 'length':
            self.subtitle_length.setValue(s_rule.get('max_length', 20))
        self.subtitle_rule_stack.setCurrentIndex(t_map.get(s_type, 0))

    def update_rule_ui(self, rule_type):
        t_map = {'位置': 0, '关键词': 1, '正则表达式': 2, '长度': 3}
        if rule_type == 'title':
            idx = t_map.get(self.title_rule_type.currentText(), 0)
            self.title_rule_stack.setCurrentIndex(idx)
        elif rule_type == 'subtitle':
            idx = t_map.get(self.subtitle_rule_type.currentText(), 0)
            self.subtitle_rule_stack.setCurrentIndex(idx)

    def save_template(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "警告", "请输入模板名称！")
            return
        # 规则保存
        t_type = self.title_rule_type.currentText()
        s_type = self.subtitle_rule_type.currentText()
        t_map = {'位置': 'position', '关键词': 'keyword', '正则表达式': 'regex', '长度': 'length'}
        s_map = t_map
        t_rule = {'type': t_map[t_type]}
        s_rule = {'type': s_map[s_type]}
        if t_type == '位置':
            t_rule['position'] = self.title_position.value()
        elif t_type == '关键词':
            t_rule['keyword'] = self.title_keyword.text().strip()
        elif t_type == '正则表达式':
            t_rule['pattern'] = self.title_regex.text().strip()
        elif t_type == '长度':
            t_rule['max_length'] = self.title_length.value()
        if s_type == '位置':
            s_rule['position'] = self.subtitle_position.value()
        elif s_type == '关键词':
            s_rule['keyword'] = self.subtitle_keyword.text().strip()
        elif s_type == '正则表达式':
            s_rule['pattern'] = self.subtitle_regex.text().strip()
        elif s_type == '长度':
            s_rule['max_length'] = self.subtitle_length.value()
        # 字号处理
        def parse_size(val):
            try:
                if val in CHINESE_FONT_SIZES:
                    return CHINESE_FONT_SIZES[val]
                return float(val)
            except:
                return 12
        template = {
            'name': name,
            'rules': {
                'title': t_rule,
                'subtitle': s_rule,
                'body': {'type': 'default'}
            },
            'formats': {
                'title': {
                    'font': self.title_font.currentText(),
                    'size': parse_size(self.title_size.currentText()),
                    'alignment': self.title_align.currentText()
                },
                'subtitle': {
                    'font': self.subtitle_font.currentText(),
                    'size': parse_size(self.subtitle_size.currentText()),
                    'alignment': self.subtitle_align.currentText()
                },
                'body': {
                    'font': self.body_font.currentText(),
                    'size': parse_size(self.body_size.currentText()),
                    'alignment': self.body_align.currentText(),
                    'indent': self.body_indent.value(),
                    'line_spacing': self.body_spacing.value()
                }
            }
        }
        self.template = template
        self.accept()

class FormatPage(QWidget):
    def __init__(self, back_to_main, show_manual, show_instruction_once, user_id):
        super().__init__()
        self._shown = False
        self.show_instruction_once = show_instruction_once
        self.user_id = user_id
        self.selected_files = []
        self.init_ui(back_to_main, show_manual)
        self.refresh_templates()
        
    def init_ui(self, back_to_main, show_manual):
        layout = QVBoxLayout()
        
        # 顶部标题和返回按钮
        top_layout = QHBoxLayout()
        title = QLabel("文档格式规范化")
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
        
        # 文件选择
        file_layout = QHBoxLayout()
        self.files_edit = QLineEdit()
        self.files_edit.setPlaceholderText("请选择要处理的文档文件（可多选）")
        self.files_edit.setReadOnly(True)
        file_btn = QPushButton("选择文件")
        file_btn.clicked.connect(self.choose_files)
        file_layout.addWidget(self.files_edit)
        file_layout.addWidget(file_btn)
        layout.addLayout(file_layout)
        
        # 输出目录
        output_layout = QHBoxLayout()
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("请选择输出目录（可选）")
        self.output_edit.setReadOnly(True)
        output_btn = QPushButton("选择输出目录")
        output_btn.clicked.connect(self.choose_output_dir)
        output_layout.addWidget(self.output_edit)
        output_layout.addWidget(output_btn)
        layout.addLayout(output_layout)
        
        # 模板选择
        template_layout = QHBoxLayout()
        template_layout.addWidget(QLabel("选择格式模板："))
        self.template_combo = QComboBox()
        self.template_combo.addItems(self.templates.keys())
        template_layout.addWidget(self.template_combo)
        
        new_template_btn = QPushButton("新建模板")
        new_template_btn.clicked.connect(self.create_template)
        template_layout.addWidget(new_template_btn)
        
        edit_template_btn = QPushButton("编辑模板")
        edit_template_btn.clicked.connect(self.edit_template)
        template_layout.addWidget(edit_template_btn)
        
        delete_template_btn = QPushButton("删除模板")
        delete_template_btn.clicked.connect(self.delete_template)
        template_layout.addWidget(delete_template_btn)
        
        layout.addLayout(template_layout)
        
        # 处理选项
        options_layout = QHBoxLayout()
        self.overwrite_check = QCheckBox("覆盖原文件")
        options_layout.addWidget(self.overwrite_check)
        layout.addLayout(options_layout)
        
        # 开始处理按钮
        self.format_btn = QPushButton("开始处理")
        self.format_btn.clicked.connect(self.start_format)
        self.format_btn.setFixedHeight(40)
        self.format_btn.setStyleSheet("border-radius:20px; background:#2563eb; color:white; font-size:18px;")
        layout.addWidget(self.format_btn)
        
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
        
        # 日志输出
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setFixedHeight(200)
        layout.addWidget(self.log)
        
        layout.addStretch()
        self.setLayout(layout)
        
    def choose_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择要处理的文档文件", "", "Word文档 (*.docx)")
        if files:
            self.selected_files = files
            self.files_edit.setText("; ".join([os.path.basename(f) for f in files]))
            
    def choose_output_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.output_edit.setText(d)
            
    def refresh_templates(self):
        self.template_combo.clear()
        ft = FormatTemplate()
        templates = ft.get_user_templates(self.user_id)
        self.templates = {t['name']: t for t in templates}
        self.template_combo.addItems(self.templates.keys())
        
    def create_template(self):
        dialog = TemplateDialog(self)
        if dialog.exec():
            template = dialog.template
            ft = FormatTemplate()
            ft.create(self.user_id, template['name'], json.dumps(template), template.get('description', ''), template.get('is_public', False))
            self.refresh_templates()
            self.template_combo.setCurrentText(template['name'])
            
    def edit_template(self):
        current = self.template_combo.currentText()
        if not current:
            QMessageBox.warning(self, "警告", "请先选择一个模板！")
            return
        template_row = self.templates[current]
        template = json.loads(template_row['config'])
        dialog = TemplateDialog(self, template)
        if dialog.exec():
            new_template = dialog.template
            ft = FormatTemplate()
            # 这里只做简单的删除再新建，实际可优化为update
            ft.create(self.user_id, new_template['name'], json.dumps(new_template), new_template.get('description', ''), new_template.get('is_public', False))
            ft.delete(template_row['id'])
            self.refresh_templates()
            self.template_combo.setCurrentText(new_template['name'])
            
    def delete_template(self):
        current = self.template_combo.currentText()
        if not current:
            QMessageBox.warning(self, "警告", "请先选择一个模板！")
            return
        template_row = self.templates[current]
        reply = QMessageBox.question(self, "确认", f"确定要删除模板 {current} 吗？",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            ft = FormatTemplate()
            ft.delete(template_row['id'])
            self.refresh_templates()
            
    def start_format(self):
        if not self.selected_files:
            QMessageBox.warning(self, "警告", "请选择要处理的文档文件！")
            return
        template_name = self.template_combo.currentText()
        if not template_name:
            QMessageBox.warning(self, "警告", "请选择格式模板！")
            return
        if not self.formatter.set_current_template(template_name):
            QMessageBox.warning(self, "警告", "模板设置失败！")
            return
        output_dir = self.output_edit.text().strip() or None
        try:
            self.format_btn.setEnabled(False)
            self.progress.setValue(0)
            self.status.setText("正在处理...")
            self.log.clear()
            result = self.formatter.format_files(
                self.selected_files,
                output_dir,
                self.progress.setValue,
                self.status.setText
            )
            self.log.append(f"处理完成！\n")
            self.log.append(f"总文件数：{result['total']}\n")
            self.log.append(f"成功处理：{result['success']}\n")
            if result['failed']:
                self.log.append("处理失败的文件：\n")
                for file, error in result['failed']:
                    self.log.append(f"- {file}: {error}\n")
            QMessageBox.information(self, "完成", "文档处理完成！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理过程中出现错误：\n{str(e)}")
        finally:
            self.format_btn.setEnabled(True)
            self.progress.setValue(0)
            self.status.setText("")

    def showEvent(self, event):
        super().showEvent(event)
        if not self._shown:
            self._shown = True
            self.show_instruction_once() 