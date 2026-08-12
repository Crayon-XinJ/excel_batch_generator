#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
主窗口 - Excel批量生成工具的核心界面
"""

import os
import sys
import argparse
from datetime import datetime, timedelta

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QDateEdit,
    QComboBox, QListWidget, QListWidgetItem,
    QFileDialog, QProgressBar, QGroupBox,
    QMessageBox, QApplication
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent

from core.config_manager import ConfigManager
from core.date_calculator import DateCalculator
from core.generate_thread import GenerateThread
from core.auto_scheduler import AutoScheduler
from core.email_reporter import EmailReporter
from core.log_manager import get_logger
from models.config_models import Rule, ProductConfig, ProductAutoConfig
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

        self.init_ui()
        self._load_config()
        self._check_auto_mode()

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
        self.rule_list = QListWidget()
        self.rule_list.setMaximumHeight(150)
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

        # 自动化配置按钮
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

        self.sheet_names = self._get_sheet_names(file_path)
        self._load_config()

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
        
        self._update_rule_list()
        self._update_nonwork_display()

    def _save_rules(self):
        product_name = self.product_name_edit.text().strip()
        template_type = self.template_type_combo.currentText()
        if not product_name:
            return
            
        date_ranges = []
        for start, end in self.date_ranges:
            date_ranges.append([start.strftime('%Y-%m-%d'), end.strftime('%Y-%m-%d')])

        config = ProductConfig(
            product_name=product_name,
            template_type=template_type,
            rules=self.rules,
            date_ranges=date_ranges
        )
        self.config_manager.save_product_config(config)

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

    def _clear_date_ranges(self):
        self.date_ranges.clear()
        self.range_list.clear()
        self._save_rules()

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
        self.rule_list.clear()
        for rule in self.rules:
            status = "✓" if rule.enabled else "✗"
            sheet = rule.sheet_name if rule.sheet_name else "所有工作表"
            item_text = f"{status} [{rule.target_type}] {rule.target} (工作表:{sheet}) -> {rule.value_type}"
            if rule.value_type == 'random':
                item_text += f" ({rule.min_val}-{rule.max_val})"
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
            self._update_rule_list()

    def _delete_rule(self):
        current_row = self.rule_list.currentRow()
        if current_row < 0 or current_row >= len(self.rules):
            QMessageBox.warning(self, "警告", "请先选择一条规则")
            return
        rule_id = self.rules[current_row].id
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

        output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if not output_dir:
            return

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
            template_type=self.template_type_combo.currentText()
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
        """检查是否以自动化模式启动"""
        # 检查命令行参数
        if '--auto' in sys.argv:
            self.logger.info("以自动化模式启动")
            # 在UI加载完成后执行自动化
            QApplication.processEvents()
            self._run_auto_mode()
    
    def _open_auto_config(self):
        """打开自动化配置窗口"""
        dialog = AutoConfigWindow(self)
        dialog.exec_()
    
    def _run_auto_mode(self):
        """执行自动化模式"""
        self.logger.info("开始执行自动化任务")
        self.status_label.setText("自动化运行中...")
        
        auto_config = self.config_manager.get_auto_config()
        if not auto_config.get('enabled', False):
            self.logger.info("自动化未启用，跳过")
            self.status_label.setText("自动化未启用")
            return
        
        # 获取产品列表
        products_str = auto_config.get('products', '')
        products = [p.strip() for p in products_str.split(',') if p.strip()] if products_str else None
        
        # 计算今日任务
        tasks = self.auto_scheduler.get_today_tasks(products)
        if not tasks:
            self.logger.info("今日无任务")
            self.status_label.setText("今日无任务")
            return
        
        self.logger.info(f"今日任务: {len(tasks)} 个产品")
        
        # 执行生成
        total_files = 0
        success_files = 0
        failed_files = 0
        
        for product_name, task_info in tasks.items():
            self.logger.info(f"处理产品: {product_name}")
            self.status_label.setText(f"处理: {product_name}")
            
            # 确定输出目录
            output_dir = task_info.get('output_dir') or os.path.join(os.path.dirname(os.path.dirname(__file__)), 'output')
            os.makedirs(output_dir, exist_ok=True)
            
            # 确定模板路径
            # 从产品配置中获取模板路径
            for template_type in ['首件', '过程', '成品']:
                if task_info.get(template_type, False):
                    config = self.config_manager.load_product_config(product_name, template_type)
                    if config and config.template_path and os.path.exists(config.template_path):
                        template_path = config.template_path
                    else:
                        # 尝试从通用模板目录查找
                        template_path = self._find_template(product_name, template_type)
                        if not template_path:
                            self.logger.warning(f"找不到模板: {product_name}_{template_type}")
                            failed_files += 1
                            continue
                    
                    try:
                        # 生成文件
                        output_filename = f"{product_name}_{template_type}检验表_{datetime.now().strftime('%Y%m%d')}.xlsx"
                        output_path = os.path.join(output_dir, output_filename)
                        
                        # 检查是否已存在
                        if os.path.exists(output_path) and not self._should_overwrite():
                            self.logger.info(f"跳过已存在: {output_filename}")
                            continue
                        
                        # 使用ExcelGenerator生成
                        from core.excel_generator import ExcelGenerator
                        gen = ExcelGenerator()
                        gen.set_template(template_path)
                        gen.set_output_dir(output_dir)
                        
                        # 加载规则
                        rules = []
                        if config:
                            rules = config.rules
                        gen.set_rules(rules)
                        gen.set_product_info(product_name, template_type)
                        
                        gen.generate(datetime.now(), output_filename)
                        success_files += 1
                        total_files += 1
                        self.logger.info(f"✓ 已生成: {output_filename}")
                        
                    except Exception as e:
                        self.logger.error(f"生成失败: {e}")
                        failed_files += 1
                        total_files += 1
        
        # 发送邮件报告
        if auto_config.get('enabled', False):
            email_config = self.config_manager.get_email_config()
            if email_config.get('enabled', False):
                self._send_auto_report(tasks, total_files, success_files, failed_files)
        
        self.status_label.setText(f"自动化完成: 成功 {success_files}, 失败 {failed_files}")
        self.logger.info(f"自动化完成: 成功 {success_files}, 失败 {failed_files}")
        
        # 如果是无界面模式，退出
        if '--no-gui' in sys.argv and auto_config.get('exit_after_run', True):
            self.logger.info("无界面模式，退出程序")
            QApplication.quit()
    
    def _find_template(self, product_name: str, template_type: str) -> str:
        """查找模板文件"""
        # 尝试从产品配置中获取模板路径
        config = self.config_manager.load_product_config(product_name, template_type)
        if config and config.template_path and os.path.exists(config.template_path):
            return config.template_path
        
        # 尝试从当前目录查找
        possible_names = [
            f"{product_name}_{template_type}检验表模板.xlsx",
            f"{product_name}_{template_type}检验表.xlsx",
            f"{product_name}_{template_type}模板.xlsx",
        ]
        for name in possible_names:
            path = os.path.join(os.path.dirname(os.path.dirname(__file__)), name)
            if os.path.exists(path):
                return path
        
        # 尝试从模板目录查找
        templates_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'templates')
        for name in possible_names:
            path = os.path.join(templates_dir, name)
            if os.path.exists(path):
                return path
        
        return ''
    
    def _should_overwrite(self) -> bool:
        """是否覆盖已存在的文件（默认不覆盖）"""
        # 可以从配置读取
        return False
    
    def _send_auto_report(self, tasks: dict, total: int, success: int, failed: int):
        """发送自动化报告邮件"""
        try:
            self.email_reporter.email_config = self.config_manager.get_email_config()
            
            # 构建报告
            stats = {
                'total': total,
                'success': success,
                'failed': failed,
                'skipped': 0,
                'duration': '自动化任务',
                'output_dir': ''
            }
            
            # 构建任务结果
            tasks_results = {}
            for product_name, task_info in tasks.items():
                tasks_results[product_name] = {}
                for t in ['首件', '过程', '成品']:
                    if task_info.get(t, False):
                        tasks_results[product_name][t] = {
                            'status': 'success' if task_info.get('dates', {}).get(t) else 'skipped',
                            'filename': f"{product_name}_{t}检验表_{datetime.now().strftime('%Y%m%d')}.xlsx"
                        }
            
            # 确定邮件主题
            subject = self.email_reporter.email_config.get('subject', 'Excel生成报告 - {date}')
            subject = subject.replace('{date}', datetime.now().strftime('%Y-%m-%d'))
            
            body = self.email_reporter.build_report(tasks_results, stats)
            self.email_reporter.send(subject, body)
            
        except Exception as e:
            self.logger.error(f"发送邮件报告失败: {e}")


# ============================================================
# 命令行入口
# ============================================================

def main():
    """程序入口（支持命令行参数）"""
    parser = argparse.ArgumentParser(description='Excel批量生成工具')
    parser.add_argument('--auto', action='store_true', help='以自动化模式运行')
    parser.add_argument('--no-gui', action='store_true', help='无界面模式（需配合--auto使用）')
    parser.add_argument('--product', type=str, help='指定产品名称')
    parser.add_argument('--type', type=str, choices=['首件', '过程', '成品'], help='模板类型')
    parser.add_argument('--date', type=str, help='指定日期 (YYYY-MM-DD)')
    parser.add_argument('--output', type=str, help='输出目录')
    parser.add_argument('--preview', action='store_true', help='预览模式')
    parser.add_argument('--verbose', action='store_true', help='详细日志')
    
    args = parser.parse_args()
    
    # 如果设置了详细日志，调整日志级别
    if args.verbose:
        import logging
        logging.getLogger('ExcelBatchGenerator').setLevel(logging.DEBUG)
    
    # 创建应用
    app = QApplication(sys.argv)
    app.setApplicationName("Excel批量生成工具")
    
    # 创建主窗口
    window = MainWindow()
    
    # 如果是自动化模式且无界面，不显示窗口
    if args.auto and args.no_gui:
        window.show()  # 需要显示以便处理事件，但可以最小化
        window.showMinimized()
        # 在事件循环中执行自动化
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(100, lambda: window._run_auto_mode())
    else:
        window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()