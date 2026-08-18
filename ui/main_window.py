#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
主窗口 - Excel批量生成工具的核心界面
"""

import os
import sys
from datetime import datetime, timedelta

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QDateEdit,
    QComboBox, QListWidget, QListWidgetItem,
    QFileDialog, QProgressBar, QGroupBox,
    QMessageBox, QApplication, QShortcut
)
from PyQt5.QtCore import Qt, QDate, QTimer
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QKeySequence

from core.config_manager import ConfigManager
from core.date_calculator import DateCalculator
from core.generate_thread import GenerateThread
from core.auto_scheduler import AutoScheduler
from core.email_reporter import EmailReporter
from core.log_manager import get_logger
from models.config_models import Rule, ProductConfig
from ui.non_workdays_dialog import NonWorkdaysDialog
from ui.rule_dialog import RuleDialog
from ui.preview_dialog import PreviewDialog
from ui.auto_config_window import AutoConfigWindow


class MainWindow(QMainWindow):
    """主窗口类"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel批量生成工具")
        self.setMinimumSize(1000, 750)
        
        # ========== 启用全窗口拖入 ==========
        self.setAcceptDrops(True)
        
        self.logger = get_logger('MainWindow')
        self.logger.info("程序启动")

        self.config_manager = ConfigManager()
        self.date_calculator = DateCalculator()
        self.auto_scheduler = AutoScheduler()
        self.email_reporter = EmailReporter()

        non_workdays = self.config_manager.load_non_workdays()
        if non_workdays:
            self.date_calculator.set_non_workdays(non_workdays)

        self.date_ranges = []
        self.rules = []
        self.product_name = ''
        self.template_type = '过程'
        self.template_path = ''
        self.sheet_names = []
        self.thread = None
        
        # 模板缺失警告标记（用于状态栏）
        self._template_warning_shown = False

        self.init_ui()
        self._load_config()
        
        # ========== 绑定快捷键 ==========
        self._setup_shortcuts()
        
        # ========== 启动后延迟检查模板路径 ==========
        QTimer.singleShot(500, self._check_template_paths)


    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        # ========== 模板加载 ==========
        template_group = QGroupBox("模板文件")
        template_layout = QHBoxLayout()

        self.template_path_edit = QLineEdit()
        self.template_path_edit.setPlaceholderText("选择或拖入模板文件")
        self.template_path_edit.setReadOnly(True)
        self.template_path_edit.setAcceptDrops(True)
        self.template_path_edit.dragEnterEvent = self._drag_enter
        self.template_path_edit.dropEvent = self._drop

        btn_browse = QPushButton("浏览...")
        btn_browse.clicked.connect(self._browse_template)

        template_layout.addWidget(self.template_path_edit)
        template_layout.addWidget(btn_browse)
        template_group.setLayout(template_layout)
        main_layout.addWidget(template_group)

        # ========== 产品信息 ==========
        info_layout = QHBoxLayout()
        info_layout.addWidget(QLabel("产品名称:"))
        self.product_name_edit = QLineEdit()
        self.product_name_edit.setPlaceholderText("自动从文件名提取")
        self.product_name_edit.setFixedWidth(200)
        info_layout.addWidget(self.product_name_edit)

        info_layout.addSpacing(20)
        info_layout.addWidget(QLabel("模板类型:"))
        self.template_type_combo = QComboBox()
        self.template_type_combo.addItems(['首件', '过程', '成品'])
        self.template_type_combo.setFixedWidth(100)
        self.template_type_combo.currentTextChanged.connect(self._on_template_type_changed)
        info_layout.addWidget(self.template_type_combo)

        info_layout.addStretch()
        main_layout.addLayout(info_layout)

        # ========== 日期范围 ==========
        range_group = QGroupBox("生产工期")
        range_layout = QVBoxLayout()

        add_range_layout = QHBoxLayout()
        add_range_layout.addWidget(QLabel("起始:"))
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.setFixedWidth(120)
        add_range_layout.addWidget(self.start_date_edit)

        add_range_layout.addWidget(QLabel("结束:"))
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.setFixedWidth(120)
        add_range_layout.addWidget(self.end_date_edit)

        btn_add_range = QPushButton("添加")
        btn_add_range.clicked.connect(self._add_date_range)
        add_range_layout.addWidget(btn_add_range)

        btn_clear_ranges = QPushButton("清空所有")
        btn_clear_ranges.clicked.connect(self._clear_date_ranges)
        add_range_layout.addWidget(btn_clear_ranges)

        add_range_layout.addWidget(QLabel("提示: 选中条目后按 Delete 键删除"))
        add_range_layout.addStretch()
        range_layout.addLayout(add_range_layout)

        self.range_list = QListWidget()
        self.range_list.setMaximumHeight(120)
        range_layout.addWidget(self.range_list)

        range_group.setLayout(range_layout)
        main_layout.addWidget(range_group)

        # ========== 非工作日 + 规则 ==========
        mid_layout = QHBoxLayout()

        nonwork_group = QGroupBox("非工作日配置")
        nonwork_layout = QVBoxLayout()
        self.nonwork_list = QListWidget()
        self.nonwork_list.setMaximumHeight(100)
        self.nonwork_list.setStyleSheet("QListWidget::item { padding: 2px; }")
        nonwork_layout.addWidget(self.nonwork_list)

        btn_nonwork = QPushButton("配置非工作日")
        btn_nonwork.clicked.connect(self._configure_non_workdays)
        nonwork_layout.addWidget(btn_nonwork)

        btn_import_nw = QPushButton("导入非工作日")
        btn_import_nw.clicked.connect(self._import_non_workdays)
        btn_export_nw = QPushButton("导出非工作日")
        btn_export_nw.clicked.connect(self._export_non_workdays)
        nw_btn_layout = QHBoxLayout()
        nw_btn_layout.addWidget(btn_import_nw)
        nw_btn_layout.addWidget(btn_export_nw)
        nonwork_layout.addLayout(nw_btn_layout)

        nonwork_group.setLayout(nonwork_layout)
        mid_layout.addWidget(nonwork_group, 1)

        rule_group = QGroupBox("内容修改规则")
        rule_layout = QVBoxLayout()
        
        # ========== 规则列表优化显示 ==========
        self.rule_list = QListWidget()
        self.rule_list.setMaximumHeight(150)
        # 双击编辑
        self.rule_list.itemDoubleClicked.connect(self._edit_rule)
        rule_layout.addWidget(self.rule_list)

        rule_btn_layout = QHBoxLayout()
        btn_add_rule = QPushButton("添加规则")
        btn_add_rule.clicked.connect(self._add_rule)
        btn_edit_rule = QPushButton("编辑")
        btn_edit_rule.clicked.connect(self._edit_rule)
        btn_del_rule = QPushButton("删除")
        btn_del_rule.clicked.connect(self._delete_rule)
        btn_import_rule = QPushButton("导入规则")
        btn_import_rule.clicked.connect(self._import_rules)
        btn_export_rule = QPushButton("导出规则")
        btn_export_rule.clicked.connect(self._export_rules)

        rule_btn_layout.addWidget(btn_add_rule)
        rule_btn_layout.addWidget(btn_edit_rule)
        rule_btn_layout.addWidget(btn_del_rule)
        rule_btn_layout.addWidget(btn_import_rule)
        rule_btn_layout.addWidget(btn_export_rule)
        rule_btn_layout.addStretch()
        rule_layout.addLayout(rule_btn_layout)

        rule_group.setLayout(rule_layout)
        mid_layout.addWidget(rule_group, 2)
        main_layout.addLayout(mid_layout)

        # ========== 底部操作 ==========
        action_layout = QHBoxLayout()

        btn_preview = QPushButton("预览")
        btn_preview.clicked.connect(self._preview)
        action_layout.addWidget(btn_preview)

        btn_generate = QPushButton("开始生成")
        btn_generate.clicked.connect(self._start_generate)
        btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        action_layout.addWidget(btn_generate)

        btn_cancel = QPushButton("中止")
        btn_cancel.clicked.connect(self._cancel_generate)
        btn_cancel.setEnabled(False)
        action_layout.addWidget(btn_cancel)

        btn_auto_config = QPushButton("自动化配置")
        btn_auto_config.clicked.connect(self._open_auto_config)
        action_layout.addWidget(btn_auto_config)

        action_layout.addStretch()

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedWidth(200)
        action_layout.addWidget(self.progress_bar)

        self.time_label = QLabel("剩余时间: --")
        action_layout.addWidget(self.time_label)

        self.status_label = QLabel("就绪")
        action_layout.addWidget(self.status_label)

        main_layout.addLayout(action_layout)

        self._update_nonwork_display()

    # ============================================================
    # 快捷键设置
    # ============================================================
    
    def _setup_shortcuts(self):
        """设置全局快捷键"""
        # Ctrl+N: 添加规则
        shortcut_new_rule = QShortcut(QKeySequence("Ctrl+N"), self)
        shortcut_new_rule.activated.connect(self._add_rule)
        
        # Ctrl+S: 保存配置
        shortcut_save = QShortcut(QKeySequence("Ctrl+S"), self)
        shortcut_save.activated.connect(self._save_rules)
        
        # Delete: 删除选中项（由 keyPressEvent 处理）
    
    def keyPressEvent(self, event):
        """处理键盘按键事件"""
        if event.key() == Qt.Key_Delete:
            # 检查当前焦点在哪个列表上
            if self.rule_list.hasFocus():
                # 删除选中的规则
                self._delete_rule()
                event.accept()
                return
            elif self.range_list.hasFocus():
                # 删除选中的日期段
                self._delete_selected_date_range()
                event.accept()
                return
        super().keyPressEvent(event)

    # ============================================================
    # 全窗口拖入
    # ============================================================
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """全窗口拖入：接受文件拖入"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith(('.xlsx', '.xls')):
                event.acceptProposedAction()
            else:
                event.ignore()
    
    def dropEvent(self, event: QDropEvent):
        """全窗口拖入：处理文件拖入"""
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith(('.xlsx', '.xls')):
                self._load_template(file_path)
                event.acceptProposedAction()
            else:
                event.ignore()

    # ============================================================
    # 启动检查模板路径
    # ============================================================
    
    def _check_template_paths(self):
        """启动时检查所有产品配置中的模板路径是否存在"""
        products = self.config_manager.get_products_list()
        if not products:
            return
        
        missing_templates = []
        for product in products:
            for template_type in ['首件', '过程', '成品']:
                config = self.config_manager.load_product_config(product, template_type)
                if config and config.template_path:
                    if not os.path.exists(config.template_path):
                        missing_templates.append(f"{product} - {template_type}")
        
        if missing_templates:
            self._template_warning_shown = True
            count = len(missing_templates)
            self.status_label.setText(f"⚠️ 检测到 {count} 个模板文件缺失，请重新加载模板")
            self.status_label.setStyleSheet("color: #FF6B00; font-weight: bold;")
            self.logger.warning(f"启动检查：{count} 个模板文件缺失: {missing_templates}")
            
            # 在状态栏显示详细信息（使用鼠标悬停提示）
            self.status_label.setToolTip("\n".join(missing_templates))

    # ============================================================
    # 非工作日显示
    # ============================================================
    
    def _update_nonwork_display(self):
        dates = self.config_manager.load_non_workdays()
        self.nonwork_list.clear()
        if dates:
            for d in sorted(dates):
                self.nonwork_list.addItem(d)
        else:
            self.nonwork_list.addItem("未配置非工作日")

    # ============================================================
    # 模板加载
    # ============================================================
    
    def _drag_enter(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def _drop(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith(('.xlsx', '.xls')):
                self._load_template(file_path)

    def _browse_template(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择模板文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self._load_template(file_path)

    def _load_template(self, file_path: str):
        """加载模板文件并保存模板路径到配置"""
        self.template_path = file_path
        self.template_path_edit.setText(file_path)
        
        self.logger.info(f"加载模板: {os.path.basename(file_path)}")

        filename = os.path.basename(file_path)
        name = os.path.splitext(filename)[0]
        for suffix in ['首件', '过程', '成品']:
            if name.endswith(suffix + '检验表模板'):
                self.product_name = name.replace(suffix + '检验表模板', '').rstrip('_')
                self.template_type = suffix
                break
            elif name.endswith(suffix + '检验表'):
                self.product_name = name.replace(suffix + '检验表', '').rstrip('_')
                self.template_type = suffix
                break
        else:
            self.product_name = name

        self.product_name_edit.setText(self.product_name)
        self.template_type_combo.setCurrentText(self.template_type)

        # 清空旧规则，防止新产品无配置时继承
        self.rules = []

        self.sheet_names = self._get_sheet_names(file_path)

        config = self.config_manager.load_product_config(self.product_name, self.template_type)
        if config is None:
            config = ProductConfig(
                product_name=self.product_name,
                template_type=self.template_type,
                rules=[]
            )
        else:
            self.rules = config.rules

        config.template_path = file_path
        self.config_manager.save_product_config(config)

        self._update_rule_list()
        self._load_config()
        self.logger.info(f"模板路径已保存: {file_path}")

    def _get_sheet_names(self, file_path: str) -> list:
        import win32com.client as win32
        import pythoncom
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            wb = excel.Workbooks.Open(os.path.abspath(file_path))
            sheets = [s.Name for s in wb.Sheets]
            wb.Close()
            return sheets
        except Exception as e:
            self.logger.warning(f"读取工作表失败: {e}")
            return []
        finally:
            try:
                if excel:
                    excel.Quit()
            except:
                pass

    # ============================================================
    # 配置加载/保存
    # ============================================================
    
    def _load_config(self):
        product_name = self.product_name_edit.text().strip()
        template_type = self.template_type_combo.currentText()
        if not product_name:
            return

        self.product_name = product_name
        self.template_type = template_type

        self.logger.info(f"加载配置: {product_name} - {template_type}")

        config = self.config_manager.load_product_config(product_name, template_type)
        if config:
            self.rules = config.rules
            self.date_ranges.clear()
            self.range_list.clear()
            for date_range in config.date_ranges:
                if len(date_range) == 2:
                    start = datetime.strptime(date_range[0], '%Y-%m-%d')
                    end = datetime.strptime(date_range[1], '%Y-%m-%d')
                    self.date_ranges.append((start, end))
                    item = QListWidgetItem(f"{start.strftime('%Y-%m-%d')} ~ {end.strftime('%Y-%m-%d')}")
                    self.range_list.addItem(item)
        else:
            self.rules = []
            self.date_ranges.clear()
            self.range_list.clear()

        self._update_rule_list()
        self._update_nonwork_display()

    def _save_rules(self):
        product_name = self.product_name_edit.text().strip()
        template_type = self.template_type_combo.currentText()
        if not product_name:
            return
        
        existing_config = self.config_manager.load_product_config(product_name, template_type)
        existing_template_path = existing_config.template_path if existing_config else ''
        existing_output_dir = existing_config.output_dir if existing_config else ''
        
        date_ranges = []
        for start, end in self.date_ranges:
            date_ranges.append([start.strftime('%Y-%m-%d'), end.strftime('%Y-%m-%d')])

        config = ProductConfig(
            product_name=product_name,
            template_type=template_type,
            template_path=existing_template_path or self.template_path,
            output_dir=existing_output_dir,
            rules=self.rules,
            date_ranges=date_ranges
        )
        self.config_manager.save_product_config(config)
        
        # Ctrl+S 保存后状态栏提示
        self.status_label.setText("配置已保存")
        QTimer.singleShot(1500, lambda: self.status_label.setText("就绪"))

    def _on_template_type_changed(self, template_type: str):
        self.template_type = template_type
        self._load_config()

    # ============================================================
    # 日期范围管理
    # ============================================================
    
    def _add_date_range(self):
        start_qdate = self.start_date_edit.date()
        end_qdate = self.end_date_edit.date()
        start_date = datetime(start_qdate.year(), start_qdate.month(), start_qdate.day())
        end_date = datetime(end_qdate.year(), end_qdate.month(), end_qdate.day())
        
        if start_date > end_date:
            QMessageBox.warning(self, "警告", "起始日期不能晚于结束日期")
            return

        self.date_ranges.append((start_date, end_date))
        item = QListWidgetItem(f"{start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")
        self.range_list.addItem(item)
        self._save_rules()

    def _delete_selected_date_range(self):
        """删除选中的单段工期"""
        current_row = self.range_list.currentRow()
        if current_row < 0:
            return
        
        # 获取要删除的条目文本用于确认
        item_text = self.range_list.currentItem().text()
        reply = QMessageBox.question(
            self, "确认删除", 
            f"确定要删除以下日期段吗？\n\n{item_text}",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.date_ranges.pop(current_row)
            self.range_list.takeItem(current_row)
            self._save_rules()
            self.logger.info(f"删除日期段: {item_text}")

    def _clear_date_ranges(self):
        if not self.date_ranges:
            return
        reply = QMessageBox.question(
            self, "确认清空", 
            "确定要清空所有生产工期吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.date_ranges.clear()
            self.range_list.clear()
            self._save_rules()
            self.logger.info("清空所有日期段")

    # ============================================================
    # 非工作日管理
    # ============================================================
    
    def _configure_non_workdays(self):
        current_dates = self.config_manager.load_non_workdays()
        dialog = NonWorkdaysDialog(current_dates, self)
        if dialog.exec_():
            dates = dialog.get_selected_dates()
            self.config_manager.save_non_workdays(dates)
            self.date_calculator.set_non_workdays(dates)
            self._update_nonwork_display()
            self.logger.info(f"更新非工作日: {len(dates)} 天")

    def _import_non_workdays(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "导入非工作日", "", "JSON文件 (*.json)")
        if file_path:
            try:
                dates = self.config_manager.import_non_workdays(file_path)
                self.config_manager.save_non_workdays(dates)
                self.date_calculator.set_non_workdays(dates)
                self._update_nonwork_display()
                self.logger.info(f"导入非工作日: {len(dates)} 天")
                QMessageBox.information(self, "成功", f"已导入 {len(dates)} 天")
            except Exception as e:
                self.logger.error(f"导入非工作日失败: {e}")
                QMessageBox.critical(self, "错误", str(e))

    def _export_non_workdays(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "导出非工作日", "", "JSON文件 (*.json)")
        if file_path:
            try:
                self.config_manager.export_non_workdays(file_path)
                self.logger.info(f"导出非工作日: {file_path}")
                QMessageBox.information(self, "成功", f"已导出到 {file_path}")
            except Exception as e:
                self.logger.error(f"导出非工作日失败: {e}")
                QMessageBox.critical(self, "错误", str(e))

    # ============================================================
    # 规则管理
    # ============================================================
    
    def _update_rule_list(self):
        """更新规则列表显示（优化格式）"""
        self.rule_list.clear()
        for rule in self.rules:
            status = "✓" if rule.enabled else "✗"
            
            # 值类型中文映射
            value_type_map = {
                'random': '随机数',
                'date': '日期替换',
                'text_with_date': '文本日期',
                'night_shift': '夜班检验员'
            }
            value_type_cn = value_type_map.get(rule.value_type, rule.value_type)
            
            # 目标类型中文映射
            target_type_map = {
                'range': '行',
                'column': '列',
                'cell': '单元格',
                'cells': '离散'
            }
            target_type_cn = target_type_map.get(rule.target_type, rule.target_type)
            
            sheet = rule.sheet_name if rule.sheet_name else "全部"
            
            # 精简显示
            if rule.value_type == 'random':
                item_text = f"{status} {rule.target} → {value_type_cn} ({rule.min_val}-{rule.max_val})  [{sheet}]"
            else:
                item_text = f"{status} {rule.target} → {value_type_cn}  [{sheet}]"
            
            item = QListWidgetItem(item_text)
            self.rule_list.addItem(item)

    def _add_rule(self):
        if not self.sheet_names:
            QMessageBox.warning(self, "警告", "请先加载模板文件以获取工作表列表")
            return
        dialog = RuleDialog(sheet_names=self.sheet_names, parent=self)
        if dialog.exec_():
            rule = dialog.get_rule()
            self.rules.append(rule)
            self._save_rules()
            self._update_rule_list()
            self.logger.info(f"添加规则: {rule.id}")

    def _edit_rule(self):
        """编辑规则（支持双击触发）"""
        current_row = self.rule_list.currentRow()
        if current_row < 0 or current_row >= len(self.rules):
            QMessageBox.warning(self, "警告", "请先选择一条规则")
            return
        rule = self.rules[current_row]
        dialog = RuleDialog(rule, sheet_names=self.sheet_names, parent=self)
        if dialog.exec_():
            new_rule = dialog.get_rule()
            self.rules[current_row] = new_rule
            self._save_rules()
            self._update_rule_list()  # 编辑后自动刷新列表
            self.logger.info(f"编辑规则: {new_rule.id}")

    def _delete_rule(self):
        """删除规则（带确认对话框）"""
        current_row = self.rule_list.currentRow()
        if current_row < 0 or current_row >= len(self.rules):
            return
        
        rule = self.rules[current_row]
        # 构建确认信息
        value_type_map = {'random': '随机数', 'date': '日期替换', 'text_with_date': '文本日期', 'night_shift': '夜班检验员'}
        rule_desc = f"{rule.target} → {value_type_map.get(rule.value_type, rule.value_type)}"
        
        reply = QMessageBox.question(
            self, "确认删除", 
            f"确定要删除以下规则吗？\n\n{rule_desc}",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            rule_id = rule.id
            self.rules.pop(current_row)
            self._save_rules()
            self._update_rule_list()
            self.logger.info(f"删除规则: {rule_id}")

    def _import_rules(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "导入规则", "", "JSON文件 (*.json)")
        if file_path:
            try:
                config = self.config_manager.import_product_config(file_path)
                if config.product_name != self.product_name or config.template_type != self.template_type:
                    reply = QMessageBox.question(
                        self, "确认", 
                        f"导入的规则属于产品'{config.product_name}'类型'{config.template_type}'，是否覆盖当前配置？",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    if reply == QMessageBox.No:
                        return
                self.rules = config.rules
                self._save_rules()
                self._update_rule_list()
                self.logger.info(f"导入规则: {len(self.rules)} 条")
                QMessageBox.information(self, "成功", f"已导入 {len(self.rules)} 条规则")
            except Exception as e:
                self.logger.error(f"导入规则失败: {e}")
                QMessageBox.critical(self, "错误", str(e))

    def _export_rules(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "导出规则", "", "JSON文件 (*.json)")
        if file_path:
            try:
                self.config_manager.export_product_config(self.product_name, self.template_type, file_path)
                self.logger.info(f"导出规则: {file_path}")
                QMessageBox.information(self, "成功", f"已导出到 {file_path}")
            except Exception as e:
                self.logger.error(f"导出规则失败: {e}")
                QMessageBox.critical(self, "错误", str(e))

    # ============================================================
    # 预览与生成
    # ============================================================
    
    def _get_target_dates(self) -> list:
        template_type = self.template_type_combo.currentText()
        if not self.date_ranges:
            QMessageBox.warning(self, "警告", "请先添加生产工期")
            return []

        dates = self.date_calculator.preview_dates(template_type, self.date_ranges)
        return dates

    def _preview(self):
        dates = self._get_target_dates()
        if not dates:
            QMessageBox.information(self, "提示", "没有需要生成的日期")
            return
        dialog = PreviewDialog(dates, self)
        dialog.exec_()

    def _start_generate(self):
        if not self.template_path or not os.path.exists(self.template_path):
            QMessageBox.warning(self, "警告", "请先选择模板文件")
            return

        product_name = self.product_name_edit.text().strip()
        if not product_name:
            QMessageBox.warning(self, "警告", "请填写产品名称")
            return

        dates = self._get_target_dates()
        if not dates:
            QMessageBox.information(self, "提示", "没有需要生成的日期")
            return

        existing_config = self.config_manager.load_product_config(product_name, self.template_type)
        default_dir = existing_config.output_dir if existing_config and existing_config.output_dir else ''
        
        output_dir = QFileDialog.getExistingDirectory(
            self, "选择输出目录", default_dir
        )
        if not output_dir:
            return

        if existing_config:
            existing_config.output_dir = output_dir
            self.config_manager.save_product_config(existing_config)
            self.logger.info(f"输出目录已记录: {output_dir}")
        else:
            new_config = ProductConfig(
                product_name=product_name,
                template_type=self.template_type,
                template_path=self.template_path,
                output_dir=output_dir,
                rules=self.rules,
                date_ranges=[[s.strftime('%Y-%m-%d'), e.strftime('%Y-%m-%d')] for s, e in self.date_ranges]
            )
            self.config_manager.save_product_config(new_config)
            self.logger.info(f"输出目录已记录: {output_dir}")

        self._set_controls_enabled(False)
        self.progress_bar.setValue(0)
        self.time_label.setText("剩余时间: 计算中...")
        self.status_label.setText("生成中...")

        self.logger.info(f"开始生成，类型: {self.template_type}，共 {len(dates)} 个文件")

        self.thread = GenerateThread()
        self.thread.setup(
            dates=dates,
            template_path=self.template_path,
            output_dir=output_dir,
            rules=self.rules,
            product_name=product_name,
            template_type=self.template_type
        )
        self.thread.progress_updated.connect(self._on_progress_updated)
        self.thread.finished_signal.connect(self._on_generate_finished)
        self.thread.error_signal.connect(self._on_generate_error)
        self.thread.time_updated.connect(self._on_time_updated)
        self.thread.start()

    def _cancel_generate(self):
        if self.thread and self.thread.isRunning():
            self.thread.cancel()
            self.status_label.setText("正在中止...")

    def _on_progress_updated(self, current: int, total: int, message: str):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.status_label.setText(message)

    def _on_time_updated(self, remaining: int):
        if remaining > 0:
            minutes = remaining // 60
            seconds = remaining % 60
            self.time_label.setText(f"剩余时间: {minutes}分{seconds}秒")
        else:
            self.time_label.setText("剩余时间: 即将完成")

    def _on_generate_finished(self, success_count: int, total: int):
        self._set_controls_enabled(True)
        self.progress_bar.setValue(total)
        self.status_label.setText(f"完成！成功 {success_count}/{total}")
        self.time_label.setText("剩余时间: --")
        self.logger.info(f"生成完成，成功 {success_count}/{total}")
        QMessageBox.information(self, "完成", f"成功生成 {success_count} 个文件")
        self.thread = None

    def _on_generate_error(self, error: str):
        self._set_controls_enabled(True)
        self.status_label.setText(f"错误: {error}")
        self.time_label.setText("剩余时间: --")
        self.logger.error(f"生成错误: {error}")
        QMessageBox.critical(self, "错误", error)
        self.thread = None

    def _set_controls_enabled(self, enabled: bool):
        for widget in self.findChildren(QPushButton):
            if widget.text() == "开始生成":
                widget.setEnabled(enabled)
            elif widget.text() == "中止":
                widget.setEnabled(not enabled)
            else:
                widget.setEnabled(enabled)

        self.template_path_edit.setEnabled(enabled)
        self.product_name_edit.setEnabled(enabled)
        self.template_type_combo.setEnabled(enabled)
        self.start_date_edit.setEnabled(enabled)
        self.end_date_edit.setEnabled(enabled)
        self.range_list.setEnabled(enabled)
        self.rule_list.setEnabled(enabled)

    # ============================================================
    # 自动化
    # ============================================================
    
    def _check_auto_mode(self):
        if '--auto' in sys.argv:
            self.logger.info("以自动化模式启动")
            QApplication.processEvents()
            self._run_auto_mode()
    
    def _open_auto_config(self):
        dialog = AutoConfigWindow(self)
        dialog.exec_()
    
    def _run_auto_mode(self):
        self.logger.info("开始执行自动化任务")
        self.status_label.setText("自动化运行中...")
        
        try:
            from core.auto_scheduler import AutoScheduler
            auto_scheduler = AutoScheduler()
            
            auto_config = self.config_manager.get_auto_config()
            if not auto_config.get('enabled', False):
                self.logger.info("自动化未启用，跳过")
                self.status_label.setText("自动化未启用")
                if '--no-gui' in sys.argv and auto_config.get('exit_after_run', True):
                    QApplication.quit()
                return
            
            products_str = auto_config.get('products', '')
            products = [p.strip() for p in products_str.split(',') if p.strip()] if products_str else None
            self.logger.info(f"products_str: '{products_str}', products: {products}")
            
            if not products:
                products = self.config_manager.get_products_list()
                self.logger.info(f"从 config_manager 获取产品列表: {products}")
            
            if not products:
                self.logger.warning("没有配置任何产品")
                self.status_label.setText("没有配置产品")
                if '--no-gui' in sys.argv and auto_config.get('exit_after_run', True):
                    QApplication.quit()
                return
            
            tasks = auto_scheduler.get_today_tasks(products)
            self.logger.info(f"tasks 结果: {tasks}")
            
            if not tasks:
                self.logger.info("今日无任务")
                self.status_label.setText("今日无任务")
                if '--no-gui' in sys.argv and auto_config.get('exit_after_run', True):
                    QApplication.quit()
                return
            
            self.logger.info(f"今日任务: {len(tasks)} 个产品")
            
            execution_results = {}
            total_files = 0
            success_files = 0
            failed_files = 0
            skipped_files = 0
            output_dirs = []
            
            for product_name, task_info in tasks.items():
                self.logger.info(f"处理产品: {product_name}")
                self.status_label.setText(f"处理: {product_name}")
                
                execution_results[product_name] = {}
                
                for template_type in ['首件', '过程', '成品']:
                    if task_info.get(template_type, False):
                        execution_results[product_name][template_type] = {
                            'status': 'pending',
                            'filename': None,
                            'error': None
                        }
                        
                        try:
                            config = self.config_manager.load_product_config(product_name, template_type)
                            
                            if config and config.template_path and os.path.exists(config.template_path):
                                template_path = config.template_path
                            else:
                                template_path = self._find_template(product_name, template_type)
                                if not template_path:
                                    self.logger.warning(f"找不到模板: {product_name}_{template_type}")
                                    execution_results[product_name][template_type]['status'] = 'failed'
                                    execution_results[product_name][template_type]['error'] = f"找不到模板: {product_name}_{template_type}"
                                    failed_files += 1
                                    continue
                            
                            if config and config.output_dir:
                                output_dir = config.output_dir
                            else:
                                output_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'output')
                                self.logger.debug(f"使用默认输出目录: {output_dir}")
                            
                            if output_dir not in output_dirs:
                                output_dirs.append(output_dir)
                            
                            os.makedirs(output_dir, exist_ok=True)
                            
                            output_filename = f"{product_name}_{template_type}检验表_{datetime.now().strftime('%Y%m%d')}.xlsx"
                            output_path = os.path.join(output_dir, output_filename)
                            
                            if os.path.exists(output_path) and not self._should_overwrite():
                                self.logger.info(f"跳过已存在: {output_filename}")
                                execution_results[product_name][template_type]['status'] = 'skipped'
                                execution_results[product_name][template_type]['filename'] = output_filename
                                skipped_files += 1
                                total_files += 1
                                continue
                            
                            from core.excel_generator import ExcelGenerator
                            gen = ExcelGenerator()
                            gen.set_template(template_path)
                            gen.set_output_dir(output_dir)
                            
                            rules = []
                            if config:
                                rules = config.rules
                            gen.set_rules(rules)
                            gen.set_product_info(product_name, template_type)
                            
                            gen.generate(datetime.now(), output_filename)
                            
                            execution_results[product_name][template_type]['status'] = 'success'
                            execution_results[product_name][template_type]['filename'] = output_filename
                            success_files += 1
                            total_files += 1
                            self.logger.info(f"✓ 已生成: {output_filename}")
                            self.logger.debug(f"  输出路径: {output_path}")
                            
                        except Exception as e:
                            error_msg = str(e)
                            self.logger.error(f"生成 {product_name} {template_type} 失败: {error_msg}")
                            self.logger.exception("详细异常信息:")
                            execution_results[product_name][template_type]['status'] = 'failed'
                            execution_results[product_name][template_type]['error'] = error_msg
                            failed_files += 1
                            total_files += 1
            
            email_config = self.config_manager.get_email_config()
            if email_config.get('enabled', False):
                self._send_auto_report(execution_results, total_files, success_files, failed_files, skipped_files, output_dirs)
            
            self.status_label.setText(f"自动化完成: 成功 {success_files}, 失败 {failed_files}, 跳过 {skipped_files}")
            self.logger.info(f"自动化完成: 成功 {success_files}, 失败 {failed_files}, 跳过 {skipped_files}")
            
        except Exception as e:
            self.logger.error(f"自动化运行异常: {e}")
            self.logger.exception("详细异常信息:")
            self.status_label.setText(f"自动化异常: {e}")
        
        finally:
            if '--no-gui' in sys.argv:
                auto_config = self.config_manager.get_auto_config()
                if auto_config.get('exit_after_run', True):
                    self.logger.info("无界面模式，退出程序")
                    QApplication.quit()
    
    def _find_template(self, product_name: str, template_type: str) -> str:
        config = self.config_manager.load_product_config(product_name, template_type)
        if config and config.template_path and os.path.exists(config.template_path):
            return config.template_path
        
        possible_names = [
            f"{product_name}_{template_type}检验表模板.xlsx",
            f"{product_name}_{template_type}检验表.xlsx",
            f"{product_name}_{template_type}模板.xlsx",
        ]
        for name in possible_names:
            path = os.path.join(os.path.dirname(os.path.dirname(__file__)), name)
            if os.path.exists(path):
                return path
        
        templates_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'templates')
        for name in possible_names:
            path = os.path.join(templates_dir, name)
            if os.path.exists(path):
                return path
        
        return ''
    
    def _should_overwrite(self) -> bool:
        return False
    
    def _send_auto_report(self, execution_results: dict, total: int, success: int, failed: int, skipped: int, output_dirs: list):
        try:
            self.email_reporter.email_config = self.config_manager.get_email_config()
            
            stats = {
                'total': total,
                'success': success,
                'failed': failed,
                'skipped': skipped,
                'duration': '自动化任务',
                'output_dirs': output_dirs
            }
            
            tasks_results = {}
            for product_name, product_results in execution_results.items():
                tasks_results[product_name] = {}
                for template_type, result in product_results.items():
                    if result['status'] != 'pending':
                        tasks_results[product_name][template_type] = {
                            'status': result['status'],
                            'filename': result.get('filename', ''),
                            'error': result.get('error', '')
                        }
            
            subject = self.email_reporter.email_config.get('subject', 'Excel生成报告 - {date}')
            subject = subject.replace('{date}', datetime.now().strftime('%Y-%m-%d'))
            
            body = self.email_reporter.build_report(tasks_results, stats)
            self.email_reporter.send(subject, body)
            
        except Exception as e:
            self.logger.error(f"发送邮件报告失败: {e}")