#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
自动化配置窗口 - 配置自动化运行参数
"""

import sys
import os
from datetime import datetime
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QGroupBox, QCheckBox, QLineEdit, QComboBox,
    QPushButton, QLabel, QListWidget, QListWidgetItem,
    QMessageBox, QFileDialog, QSpinBox, QTabWidget,
    QWidget, QScrollArea, QFrame, QDateEdit
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal

from core.config_manager import ConfigManager
from core.log_manager import get_logger
from core.email_reporter import EmailReporter
from models.config_models import ProductConfig


class AutoConfigWindow(QDialog):
    """自动化配置窗口"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("自动化配置 - Excel批量生成工具")
        self.setMinimumSize(800, 750)
        
        self.logger = get_logger('AutoConfigWindow')
        self.config_manager = ConfigManager()
        self.email_reporter = EmailReporter()
        
        # 当前选中的产品
        self.current_product = ""
        
        # 日期范围数据
        self.date_ranges = []  # [(start_date, end_date), ...]
        
        self.init_ui()
        self._load_config()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        tabs = QTabWidget()
        
        # ===== 选项卡1: 基本设置 =====
        basic_tab = QWidget()
        basic_layout = QVBoxLayout(basic_tab)
        
        self.enable_check = QCheckBox("启用自动化")
        self.enable_check.setChecked(False)
        basic_layout.addWidget(self.enable_check)
        
        time_layout = QHBoxLayout()
        time_layout.addWidget(QLabel("运行时间:"))
        self.time_edit = QLineEdit()
        self.time_edit.setPlaceholderText("07:00")
        self.time_edit.setFixedWidth(80)
        time_layout.addWidget(self.time_edit)
        time_layout.addWidget(QLabel("(24小时制)"))
        time_layout.addStretch()
        basic_layout.addLayout(time_layout)
        
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("运行模式:"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(['定时触发', '手动触发', '混合模式'])
        self.mode_combo.setFixedWidth(120)
        mode_layout.addWidget(self.mode_combo)
        mode_layout.addStretch()
        basic_layout.addLayout(mode_layout)
        
        range_mode_layout = QHBoxLayout()
        range_mode_layout.addWidget(QLabel("工期模式:"))
        self.range_mode_combo = QComboBox()
        self.range_mode_combo.addItems(['手动 (manual)', '自动 (auto)', '混合 (hybrid)'])
        self.range_mode_combo.setFixedWidth(150)
        range_mode_layout.addWidget(self.range_mode_combo)
        range_mode_layout.addStretch()
        basic_layout.addLayout(range_mode_layout)
        
        days_layout = QHBoxLayout()
        days_layout.addWidget(QLabel("自动推断天数:"))
        self.days_spin = QSpinBox()
        self.days_spin.setRange(7, 365)
        self.days_spin.setValue(30)
        self.days_spin.setFixedWidth(80)
        days_layout.addWidget(self.days_spin)
        days_layout.addWidget(QLabel("(当无配置时向前/后推算)"))
        days_layout.addStretch()
        basic_layout.addLayout(days_layout)
        
        types_layout = QHBoxLayout()
        types_layout.addWidget(QLabel("生成类型:"))
        self.types_combo = QComboBox()
        self.types_combo.addItems(['全部', '仅首件', '仅过程', '仅成品', '首件+过程', '过程+成品'])
        self.types_combo.setFixedWidth(150)
        types_layout.addWidget(self.types_combo)
        types_layout.addStretch()
        basic_layout.addLayout(types_layout)
        
        self.exit_check = QCheckBox("生成完成后自动退出")
        self.exit_check.setChecked(True)
        basic_layout.addWidget(self.exit_check)
        
        self.check_template_check = QCheckBox("生成前检查模板文件是否存在")
        self.check_template_check.setChecked(True)
        basic_layout.addWidget(self.check_template_check)
        
        basic_layout.addStretch()
        tabs.addTab(basic_tab, "基本设置")
        
        # ===== 选项卡2: 产品管理 =====
        product_tab = QWidget()
        product_layout = QVBoxLayout(product_tab)
        
        product_layout.addWidget(QLabel("参与自动化的产品:"))
        
        self.product_list = QListWidget()
        self.product_list.setMaximumHeight(150)
        self.product_list.itemClicked.connect(self._on_product_selected)
        product_layout.addWidget(self.product_list)
        
        product_btn_layout = QHBoxLayout()
        btn_add_product = QPushButton("添加产品")
        btn_add_product.clicked.connect(self._add_product)
        btn_remove_product = QPushButton("移除产品")
        btn_remove_product.clicked.connect(self._remove_product)
        btn_refresh_products = QPushButton("刷新列表")
        btn_refresh_products.clicked.connect(self._refresh_products)
        product_btn_layout.addWidget(btn_add_product)
        product_btn_layout.addWidget(btn_remove_product)
        product_btn_layout.addWidget(btn_refresh_products)
        product_btn_layout.addStretch()
        product_layout.addLayout(product_btn_layout)
        
        # ===== 产品输出目录配置 =====
        output_group = QGroupBox("产品输出目录配置")
        output_group_layout = QVBoxLayout(output_group)
        
        self.selected_product_label = QLabel("请从上方选择一个产品")
        output_group_layout.addWidget(self.selected_product_label)
        
        # 三种类型的输出目录配置
        self.output_configs = {}
        types = ['首件', '过程', '成品']
        for t in types:
            row_layout = QHBoxLayout()
            row_layout.addWidget(QLabel(f"{t}:"))
            edit = QLineEdit()
            edit.setPlaceholderText("未配置，将使用默认路径")
            edit.setMinimumWidth(300)
            row_layout.addWidget(edit)
            btn_browse = QPushButton("浏览...")
            btn_browse.clicked.connect(lambda checked, typ=t, e=edit: self._browse_output_dir(typ, e))
            row_layout.addWidget(btn_browse)
            row_layout.addStretch()
            output_group_layout.addLayout(row_layout)
            self.output_configs[t] = edit
        
        # ===== 生产工期配置 =====
        date_group = QGroupBox("📅 生产工期配置")
        date_group_layout = QVBoxLayout(date_group)
        
        # 提示文字
        hint_label = QLabel("💡 生产工期将同步应用到「首件」「过程」「成品」三种模板")
        hint_label.setStyleSheet("color: #666; font-size: 11px;")
        date_group_layout.addWidget(hint_label)
        
        date_range_layout = QHBoxLayout()
        date_range_layout.addWidget(QLabel("起始:"))
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.setFixedWidth(120)
        date_range_layout.addWidget(self.start_date_edit)
        
        date_range_layout.addWidget(QLabel("结束:"))
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.setFixedWidth(120)
        date_range_layout.addWidget(self.end_date_edit)
        
        btn_add_range = QPushButton("添加")
        btn_add_range.clicked.connect(self._add_date_range)
        date_range_layout.addWidget(btn_add_range)
        
        btn_clear_ranges = QPushButton("清空所有")
        btn_clear_ranges.clicked.connect(self._clear_date_ranges)
        date_range_layout.addWidget(btn_clear_ranges)
        
        date_range_layout.addStretch()
        date_group_layout.addLayout(date_range_layout)
        
        self.date_range_list = QListWidget()
        self.date_range_list.setMaximumHeight(100)
        self.date_range_list.setSelectionMode(QListWidget.SingleSelection)
        date_group_layout.addWidget(self.date_range_list)
        
        # 删除选中日期段的按钮
        del_btn_layout = QHBoxLayout()
        btn_del_range = QPushButton("删除选中日期段")
        btn_del_range.clicked.connect(self._delete_selected_date_range)
        del_btn_layout.addWidget(btn_del_range)
        del_btn_layout.addStretch()
        date_group_layout.addLayout(del_btn_layout)
        
        output_group_layout.addWidget(date_group)
        
        product_layout.addWidget(output_group)
        product_layout.addStretch()
        tabs.addTab(product_tab, "产品管理")
        
        # ===== 选项卡3: 邮件通知 =====
        email_tab = QWidget()
        email_layout = QVBoxLayout(email_tab)
        
        self.email_enable_check = QCheckBox("启用邮件通知")
        email_layout.addWidget(self.email_enable_check)
        
        form = QFormLayout()
        
        self.smtp_server_edit = QLineEdit()
        self.smtp_server_edit.setPlaceholderText("smtp.qq.com")
        form.addRow("SMTP服务器:", self.smtp_server_edit)
        
        self.smtp_port_edit = QLineEdit()
        self.smtp_port_edit.setPlaceholderText("465")
        form.addRow("端口:", self.smtp_port_edit)
        
        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("your_email@qq.com")
        form.addRow("用户名:", self.username_edit)
        
        self.password_edit = QLineEdit()
        self.password_edit.setPlaceholderText("授权码")
        self.password_edit.setEchoMode(QLineEdit.Password)
        form.addRow("密码/授权码:", self.password_edit)
        
        self.to_edit = QLineEdit()
        self.to_edit.setPlaceholderText("manager@company.com")
        form.addRow("收件人:", self.to_edit)
        
        self.cc_edit = QLineEdit()
        self.cc_edit.setPlaceholderText("backup@company.com (多个用逗号分隔)")
        form.addRow("抄送:", self.cc_edit)
        
        self.subject_edit = QLineEdit()
        self.subject_edit.setPlaceholderText("Excel生成报告 - {date}")
        form.addRow("邮件主题:", self.subject_edit)
        
        self.report_level_combo = QComboBox()
        self.report_level_combo.addItems(['摘要 (summary)', '详细信息 (detailed)', '完整日志 (full)'])
        form.addRow("报告级别:", self.report_level_combo)
        
        self.send_on_error_check = QCheckBox("仅在有错误时发送")
        form.addRow("", self.send_on_error_check)
        
        email_layout.addLayout(form)
        
        email_btn_layout = QHBoxLayout()
        btn_test_email = QPushButton("测试发送")
        btn_test_email.clicked.connect(self._test_email)
        email_btn_layout.addWidget(btn_test_email)
        email_btn_layout.addStretch()
        email_layout.addLayout(email_btn_layout)
        
        email_layout.addStretch()
        tabs.addTab(email_tab, "邮件通知")
        
        # ===== 选项卡4: 定时任务 =====
        task_tab = QWidget()
        task_layout = QVBoxLayout(task_tab)
        
        self.task_status_label = QLabel("状态: 未创建")
        task_layout.addWidget(self.task_status_label)
        
        task_btn_layout = QHBoxLayout()
        btn_create_task = QPushButton("创建定时任务")
        btn_create_task.clicked.connect(self._create_task)
        btn_delete_task = QPushButton("删除定时任务")
        btn_delete_task.clicked.connect(self._delete_task)
        btn_run_now = QPushButton("立即运行")
        btn_run_now.clicked.connect(self._run_now)
        task_btn_layout.addWidget(btn_create_task)
        task_btn_layout.addWidget(btn_delete_task)
        task_btn_layout.addWidget(btn_run_now)
        task_btn_layout.addStretch()
        task_layout.addLayout(task_btn_layout)
        
        task_layout.addStretch()
        tabs.addTab(task_tab, "定时任务")
        
        layout.addWidget(tabs)
        
        # ===== 底部按钮 =====
        btn_layout = QHBoxLayout()
        btn_save = QPushButton("保存配置")
        btn_save.clicked.connect(self._save_config)
        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)
    
    def _load_config(self):
        auto_config = self.config_manager.get_auto_config()
        email_config = self.config_manager.get_email_config()
        
        self.enable_check.setChecked(auto_config.get('enabled', False))
        self.time_edit.setText(auto_config.get('time', '07:00'))
        
        mode_map = {'scheduled': 0, 'once': 1, 'scheduled': 0}
        self.mode_combo.setCurrentIndex(mode_map.get(auto_config.get('run_mode', 'scheduled'), 0))
        
        range_mode_map = {'manual': 0, 'auto': 1, 'hybrid': 2}
        self.range_mode_combo.setCurrentIndex(range_mode_map.get(auto_config.get('range_mode', 'hybrid'), 2))
        self.days_spin.setValue(auto_config.get('auto_default_days', 30))
        
        types_str = auto_config.get('types', '首件,过程,成品')
        types_map = {
            '首件,过程,成品': 0,
            '首件': 1,
            '过程': 2,
            '成品': 3,
            '首件,过程': 4,
            '过程,成品': 5
        }
        self.types_combo.setCurrentIndex(types_map.get(types_str, 0))
        
        self.exit_check.setChecked(auto_config.get('exit_after_run', True))
        self.check_template_check.setChecked(auto_config.get('check_template_exists', True))
        
        products_str = auto_config.get('products', '')
        self.product_list.clear()
        if products_str:
            for p in products_str.split(','):
                self.product_list.addItem(p.strip())
        
        self.email_enable_check.setChecked(email_config.get('enabled', False))
        self.smtp_server_edit.setText(email_config.get('smtp_server', ''))
        self.smtp_port_edit.setText(str(email_config.get('smtp_port', 465)))
        self.username_edit.setText(email_config.get('username', ''))
        self.password_edit.setText(email_config.get('password', ''))
        self.to_edit.setText(email_config.get('to', ''))
        self.cc_edit.setText(email_config.get('cc', ''))
        self.subject_edit.setText(email_config.get('subject', 'Excel生成报告 - {date}'))
        
        report_map = {'summary': 0, 'detailed': 1, 'full': 2}
        self.report_level_combo.setCurrentIndex(report_map.get(email_config.get('report_level', 'detailed'), 1))
        self.send_on_error_check.setChecked(email_config.get('send_only_on_error', False))
        
        # 初始化日期范围列表
        self.date_ranges = []
        self._update_date_range_list()
    
    def _save_config(self):
        types_map = {
            0: '首件,过程,成品',
            1: '首件',
            2: '过程',
            3: '成品',
            4: '首件,过程',
            5: '过程,成品'
        }
        
        mode_map = {0: 'scheduled', 1: 'once', 2: 'scheduled'}
        range_mode_map = {0: 'manual', 1: 'auto', 2: 'hybrid'}
        
        products = []
        for i in range(self.product_list.count()):
            products.append(self.product_list.item(i).text())
        
        # 获取运行时间并规范化
        raw_time = self.time_edit.text().strip()
        normalized_time = self._normalize_time(raw_time)
        self.time_edit.setText(normalized_time)
        
        auto_config = {
            'enabled': self.enable_check.isChecked(),
            'time': normalized_time,
            'products': ','.join(products),
            'types': types_map.get(self.types_combo.currentIndex(), '首件,过程,成品'),
            'range_mode': range_mode_map.get(self.range_mode_combo.currentIndex(), 'hybrid'),
            'auto_default_days': self.days_spin.value(),
            'run_mode': mode_map.get(self.mode_combo.currentIndex(), 'scheduled'),
            'exit_after_run': self.exit_check.isChecked(),
            'check_template_exists': self.check_template_check.isChecked()
        }
        self.config_manager.save_auto_config(auto_config)
        
        report_map = {0: 'summary', 1: 'detailed', 2: 'full'}
        email_config = {
            'enabled': self.email_enable_check.isChecked(),
            'smtp_server': self.smtp_server_edit.text().strip(),
            'smtp_port': int(self.smtp_port_edit.text().strip()) if self.smtp_port_edit.text().strip() else 465,
            'username': self.username_edit.text().strip(),
            'password': self.password_edit.text().strip(),
            'to': self.to_edit.text().strip(),
            'cc': self.cc_edit.text().strip(),
            'subject': self.subject_edit.text().strip(),
            'report_level': report_map.get(self.report_level_combo.currentIndex(), 'detailed'),
            'send_only_on_error': self.send_on_error_check.isChecked()
        }
        self.config_manager.save_email_config(email_config)
        
        # 保存产品输出目录
        self._save_product_output_dirs()
        
        # 保存当前产品的日期范围
        self._save_product_date_ranges()
        
        QMessageBox.information(self, "成功", "配置已保存")
        self.logger.info("自动化配置已保存")
        self.accept()
    
    def _save_product_output_dirs(self):
        """保存当前选中产品的输出目录配置"""
        if not self.current_product:
            return
        
        for template_type, edit in self.output_configs.items():
            output_dir = edit.text().strip()
            if output_dir:
                self.config_manager.save_product_output_dir(self.current_product, template_type, output_dir)
                self.logger.debug(f"保存 {self.current_product} {template_type} 输出目录: {output_dir}")
    
    def _save_product_date_ranges(self):
        """保存当前选中产品的日期范围到所有类型配置中"""
        if not self.current_product:
            return
        
        # 即使日期范围为空也要保存（清空配置）
        date_ranges_str = [
            [start.strftime('%Y-%m-%d'), end.strftime('%Y-%m-%d')]
            for start, end in self.date_ranges
        ]
        
        # 更新到该产品的所有类型配置中
        for template_type in ['首件', '过程', '成品']:
            config = self.config_manager.load_product_config(self.current_product, template_type)
            if config is None:
                config = ProductConfig(
                    product_name=self.current_product,
                    template_type=template_type
                )
            config.date_ranges = date_ranges_str
            self.config_manager.save_product_config(config)
            self.logger.debug(f"保存 {self.current_product} {template_type} 日期范围: {date_ranges_str}")
    
    def _on_product_selected(self, item):
        """产品选中时加载其输出目录配置和日期范围"""
        product_name = item.text()
        self.current_product = product_name
        self.selected_product_label.setText(f"当前产品: {product_name}")
        
        # 加载输出目录
        for template_type, edit in self.output_configs.items():
            output_dir = self.config_manager.get_product_output_dir(product_name, template_type)
            edit.setText(output_dir)
        
        # ============================================================
        # 加载日期范围并检测一致性
        # ============================================================
        # 1. 读取所有类型的日期范围
        all_ranges = {}
        for template_type in ['首件', '过程', '成品']:
            config = self.config_manager.load_product_config(product_name, template_type)
            if config and config.date_ranges:
                all_ranges[template_type] = config.date_ranges
        
        # 2. 确定要显示的日期范围（优先首件）
        display_ranges = []
        if '首件' in all_ranges and all_ranges['首件']:
            display_ranges = all_ranges['首件']
        elif '过程' in all_ranges and all_ranges['过程']:
            display_ranges = all_ranges['过程']
        elif '成品' in all_ranges and all_ranges['成品']:
            display_ranges = all_ranges['成品']
        
        # 3. 加载显示范围到界面
        self.date_ranges = []
        if display_ranges:
            for dr in display_ranges:
                if len(dr) == 2:
                    try:
                        start = datetime.strptime(dr[0], '%Y-%m-%d')
                        end = datetime.strptime(dr[1], '%Y-%m-%d')
                        self.date_ranges.append((start, end))
                    except:
                        pass
        self._update_date_range_list()
        
        # 4. 检测一致性并弹窗提示
        self._check_date_ranges_consistency(product_name, all_ranges)
    
    def _check_date_ranges_consistency(self, product_name: str, all_ranges: dict):
        """
        检查三种类型的日期范围是否一致，不一致时弹窗提示
        """
        # 过滤掉空配置
        non_empty = {k: v for k, v in all_ranges.items() if v}
        
        # 如果只有一种类型有配置，无需提示
        if len(non_empty) <= 1:
            return
        
        # 取第一个作为基准比较
        first_type = list(non_empty.keys())[0]
        first_ranges = non_empty[first_type]
        
        # 检查是否所有非空配置都相同
        is_consistent = True
        for t, r in non_empty.items():
            if r != first_ranges:
                is_consistent = False
                break
        
        if is_consistent:
            return
        
        # ============================================================
        # 不一致：构建提示信息
        # ============================================================
        msg = "检测到首件、过程、成品的生产工期不一致：\n\n"
        for t in ['首件', '过程', '成品']:
            ranges = all_ranges.get(t, [])
            if ranges:
                range_str = ", ".join([f"{r[0]}~{r[1]}" for r in ranges])
                msg += f"  {t}: {range_str}\n"
            else:
                msg += f"  {t}: （未配置）\n"
        
        msg += "\n是否将「首件」的工期应用到所有类型？"
        
        reply = QMessageBox.question(
            self,
            "工期不一致",
            msg,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Yes:
            # 用户选择统一：用首件配置覆盖所有类型
            first_ranges = all_ranges.get('首件', [])
            if first_ranges:
                self.logger.info(f"用户选择统一工期: 首件 {first_ranges}")
                self._sync_date_ranges_to_all(product_name, first_ranges)
                # 刷新界面显示
                self.date_ranges = []
                for dr in first_ranges:
                    if len(dr) == 2:
                        try:
                            start = datetime.strptime(dr[0], '%Y-%m-%d')
                            end = datetime.strptime(dr[1], '%Y-%m-%d')
                            self.date_ranges.append((start, end))
                        except:
                            pass
                self._update_date_range_list()
                QMessageBox.information(self, "已完成", "已将首件工期同步到过程和成品")
            else:
                self.logger.info("首件没有日期配置，无法同步")
                QMessageBox.warning(self, "提示", "首件没有日期配置，无法同步")
        else:
            self.logger.info("用户选择保持现状，暂不同步")
            # 保持现状：界面显示的是首件日期（如果首件有），但过程和成品保留原样
    
    def _sync_date_ranges_to_all(self, product_name: str, date_ranges: list):
        """将指定日期范围同步到所有类型"""
        for template_type in ['首件', '过程', '成品']:
            config = self.config_manager.load_product_config(product_name, template_type)
            if config is None:
                config = ProductConfig(
                    product_name=product_name,
                    template_type=template_type
                )
            config.date_ranges = date_ranges
            self.config_manager.save_product_config(config)
            self.logger.debug(f"同步 {product_name} {template_type} 日期范围: {date_ranges}")
    
    def _update_date_range_list(self):
        """更新日期范围列表显示"""
        self.date_range_list.clear()
        for start, end in self.date_ranges:
            self.date_range_list.addItem(f"{start.strftime('%Y-%m-%d')} ~ {end.strftime('%Y-%m-%d')}")
    
    def _add_date_range(self):
        """添加日期范围"""
        start_qdate = self.start_date_edit.date()
        end_qdate = self.end_date_edit.date()
        start_date = datetime(start_qdate.year(), start_qdate.month(), start_qdate.day())
        end_date = datetime(end_qdate.year(), end_qdate.month(), end_qdate.day())
        
        if start_date > end_date:
            QMessageBox.warning(self, "警告", "起始日期不能晚于结束日期")
            return
        
        # 检查是否已存在相同的日期范围
        for s, e in self.date_ranges:
            if s == start_date and e == end_date:
                QMessageBox.warning(self, "警告", "该日期范围已存在")
                return
        
        self.date_ranges.append((start_date, end_date))
        self._update_date_range_list()
        self.logger.debug(f"添加日期范围: {start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")
    
    def _clear_date_ranges(self):
        """清空所有日期范围"""
        if not self.date_ranges:
            return
        reply = QMessageBox.question(
            self,
            "确认清空",
            "确定要清空所有生产工期吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.date_ranges.clear()
            self._update_date_range_list()
            self.logger.debug("清空所有日期范围")
    
    def _delete_selected_date_range(self):
        """删除选中的日期范围"""
        current_row = self.date_range_list.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "提示", "请先选择一个日期段")
            return
        
        item_text = self.date_range_list.currentItem().text()
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除以下日期段吗？\n\n{item_text}",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.date_ranges.pop(current_row)
            self._update_date_range_list()
            self.logger.debug(f"删除日期段: {item_text}")
    
    def _browse_output_dir(self, template_type, edit_widget):
        """浏览选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(
            self, 
            f"选择 {template_type} 输出目录", 
            edit_widget.text() or ""
        )
        if dir_path:
            edit_widget.setText(dir_path)
            # 自动保存到当前选中的产品
            if self.current_product:
                self.config_manager.save_product_output_dir(self.current_product, template_type, dir_path)
                self.logger.debug(f"保存 {self.current_product} {template_type} 输出目录: {dir_path}")
    
    def _add_product(self):
        products = self.config_manager.get_products_list()
        if not products:
            QMessageBox.warning(self, "警告", "没有已配置的产品")
            return
        
        dialog = QDialog(self)
        dialog.setWindowTitle("选择产品")
        layout = QVBoxLayout(dialog)
        
        combo = QComboBox()
        for p in products:
            combo.addItem(p)
        layout.addWidget(combo)
        
        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("确定")
        btn_ok.clicked.connect(dialog.accept)
        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(dialog.reject)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)
        
        if dialog.exec_() == QDialog.Accepted:
            selected = combo.currentText()
            for i in range(self.product_list.count()):
                if self.product_list.item(i).text() == selected:
                    return
            self.product_list.addItem(selected)
    
    def _remove_product(self):
        current_row = self.product_list.currentRow()
        if current_row >= 0:
            self.product_list.takeItem(current_row)
    
    def _refresh_products(self):
        current_items = []
        for i in range(self.product_list.count()):
            current_items.append(self.product_list.item(i).text())
        
        all_products = self.config_manager.get_products_list()
        
        self.product_list.clear()
        for p in all_products:
            if p in current_items:
                self.product_list.addItem(p)
    
    def _test_email(self):
        email_config = {
            'enabled': True,
            'smtp_server': self.smtp_server_edit.text().strip(),
            'smtp_port': int(self.smtp_port_edit.text().strip()) if self.smtp_port_edit.text().strip() else 465,
            'username': self.username_edit.text().strip(),
            'password': self.password_edit.text().strip(),
            'to': self.to_edit.text().strip(),
            'cc': self.cc_edit.text().strip(),
            'smtp_ssl': True,
            'smtp_tls': False
        }
        
        self.email_reporter.email_config = email_config
        
        success, message = self.email_reporter.test_connection()
        if success:
            test_subject = "Excel批量生成工具 - 测试邮件"
            test_body = "这是一封测试邮件，证明邮件配置正确。"
            send_success = self.email_reporter.send(test_subject, test_body)
            if send_success:
                QMessageBox.information(self, "成功", "测试邮件已发送，请检查收件箱")
            else:
                QMessageBox.warning(self, "警告", "邮件发送失败，请检查收件人地址")
        else:
            QMessageBox.critical(self, "错误", f"连接失败: {message}")

    def _normalize_time(self, time_str: str) -> str:
        """
        将用户输入的时间规范化为 HH:MM 格式
        
        支持的输入格式：
        - 9:28   -> 09:28
        - 09:28  -> 09:28
        - 9：28  -> 09:28（全角冒号）
        - 928    -> 09:28（纯数字）
        - 09:28  -> 09:28
        """
        if not time_str:
            return "07:00"
        
        time_str = time_str.strip()
        time_str = time_str.replace('：', ':')
        
        # 纯数字处理
        if time_str.isdigit():
            if len(time_str) == 4:
                return f"{time_str[:2]}:{time_str[2:]}"
            elif len(time_str) == 3:
                return f"0{time_str[0]}:{time_str[1:]}"
            elif len(time_str) == 2:
                return f"{time_str}:00"
            elif len(time_str) == 1:
                return f"0{time_str}:00"
        
        # 包含冒号处理
        if ':' in time_str:
            parts = time_str.split(':')
            if len(parts) == 2:
                hour_str = parts[0].strip()
                min_str = parts[1].strip()
                
                try:
                    hour = int(hour_str)
                    if hour < 0 or hour > 23:
                        hour = 7
                except ValueError:
                    hour = 7
                
                try:
                    minute = int(min_str)
                    if minute < 0 or minute > 59:
                        minute = 0
                except ValueError:
                    minute = 0
                
                return f"{hour:02d}:{minute:02d}"
        
        return "07:00"
    
    def _create_task(self):
        try:
            import subprocess
            
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = sys.executable
                script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'main.py')
                exe_path = f'"{exe_path}" "{script_path}"'
            
            task_name = "ExcelBatchGenerator_AutoRun"
            time_str = self.time_edit.text().strip()
            if not time_str:
                time_str = "07:00"
            
            cmd = f'schtasks /create /tn "{task_name}" /tr "{exe_path} --auto --no-gui" /sc daily /st {time_str} /f'
            
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                self.task_status_label.setText(f"状态: 已创建 (每天 {time_str} 运行)")
                QMessageBox.information(self, "成功", f"定时任务已创建\n任务名称: {task_name}\n运行时间: 每天 {time_str}")
            else:
                QMessageBox.warning(self, "警告", f"创建任务失败: {result.stderr}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"创建任务失败: {str(e)}")
    
    def _delete_task(self):
        try:
            import subprocess
            task_name = "ExcelBatchGenerator_AutoRun"
            cmd = f'schtasks /delete /tn "{task_name}" /f'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                self.task_status_label.setText("状态: 已删除")
                QMessageBox.information(self, "成功", "定时任务已删除")
            else:
                QMessageBox.warning(self, "警告", f"删除任务失败: {result.stderr}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"删除任务失败: {str(e)}")
    
    def _run_now(self):
        QMessageBox.information(self, "提示", "此功能将在主窗口中执行\n请关闭此窗口后点击主窗口的 '开始生成' 按钮")