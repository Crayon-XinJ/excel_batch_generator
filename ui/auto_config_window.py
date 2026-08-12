#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
自动化配置窗口 - 配置自动化运行参数
"""

import sys
from datetime import datetime
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QGroupBox, QCheckBox, QLineEdit, QComboBox,
    QPushButton, QLabel, QListWidget, QListWidgetItem,
    QMessageBox, QFileDialog, QSpinBox, QTabWidget,
    QWidget
)
from PyQt5.QtCore import Qt

from core.config_manager import ConfigManager
from core.log_manager import get_logger
from core.email_reporter import EmailReporter


class AutoConfigWindow(QDialog):
    """自动化配置窗口"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("自动化配置 - Excel批量生成工具")
        self.setMinimumSize(750, 600)
        
        self.logger = get_logger('AutoConfigWindow')
        self.config_manager = ConfigManager()
        self.email_reporter = EmailReporter()
        
        self.init_ui()
        self._load_config()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # 创建选项卡
        tabs = QTabWidget()
        
        # ===== 选项卡1: 基本设置 =====
        basic_tab = QWidget()
        basic_layout = QVBoxLayout(basic_tab)
        
        # 启用自动化
        self.enable_check = QCheckBox("启用自动化")
        self.enable_check.setChecked(False)
        basic_layout.addWidget(self.enable_check)
        
        # 运行时间
        time_layout = QHBoxLayout()
        time_layout.addWidget(QLabel("运行时间:"))
        self.time_edit = QLineEdit()
        self.time_edit.setPlaceholderText("07:00")
        self.time_edit.setFixedWidth(80)
        time_layout.addWidget(self.time_edit)
        time_layout.addWidget(QLabel("(24小时制)"))
        time_layout.addStretch()
        basic_layout.addLayout(time_layout)
        
        # 运行模式
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("运行模式:"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(['定时触发', '手动触发', '混合模式'])
        self.mode_combo.setFixedWidth(120)
        mode_layout.addWidget(self.mode_combo)
        mode_layout.addStretch()
        basic_layout.addLayout(mode_layout)
        
        # 工期模式
        range_mode_layout = QHBoxLayout()
        range_mode_layout.addWidget(QLabel("工期模式:"))
        self.range_mode_combo = QComboBox()
        self.range_mode_combo.addItems(['手动 (manual)', '自动 (auto)', '混合 (hybrid)'])
        self.range_mode_combo.setFixedWidth(150)
        range_mode_layout.addWidget(self.range_mode_combo)
        range_mode_layout.addStretch()
        basic_layout.addLayout(range_mode_layout)
        
        # 自动推断天数
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
        
        # 生成类型
        types_layout = QHBoxLayout()
        types_layout.addWidget(QLabel("生成类型:"))
        self.types_combo = QComboBox()
        self.types_combo.addItems(['全部', '仅首件', '仅过程', '仅成品', '首件+过程', '过程+成品'])
        self.types_combo.setFixedWidth(150)
        types_layout.addWidget(self.types_combo)
        types_layout.addStretch()
        basic_layout.addLayout(types_layout)
        
        # 退出选项
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
        
        product_layout.addWidget(QLabel("提示: 留空表示所有已配置的产品都参与"))
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
        """加载配置到界面"""
        auto_config = self.config_manager.get_auto_config()
        email_config = self.config_manager.get_email_config()
        
        # 基本设置
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
        
        # 产品列表
        products_str = auto_config.get('products', '')
        self.product_list.clear()
        if products_str:
            for p in products_str.split(','):
                self.product_list.addItem(p.strip())
        
        # 邮件配置
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
    
    def _save_config(self):
        """保存配置"""
        # 保存自动化配置
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
        
        # 获取产品列表
        products = []
        for i in range(self.product_list.count()):
            products.append(self.product_list.item(i).text())
        
        auto_config = {
            'enabled': self.enable_check.isChecked(),
            'time': self.time_edit.text().strip(),
            'products': ','.join(products),
            'types': types_map.get(self.types_combo.currentIndex(), '首件,过程,成品'),
            'range_mode': range_mode_map.get(self.range_mode_combo.currentIndex(), 'hybrid'),
            'auto_default_days': self.days_spin.value(),
            'run_mode': mode_map.get(self.mode_combo.currentIndex(), 'scheduled'),
            'exit_after_run': self.exit_check.isChecked(),
            'check_template_exists': self.check_template_check.isChecked()
        }
        self.config_manager.save_auto_config(auto_config)
        
        # 保存邮件配置
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
        
        QMessageBox.information(self, "成功", "配置已保存")
        self.logger.info("自动化配置已保存")
        self.accept()
    
    def _add_product(self):
        """添加产品到列表"""
        products = self.config_manager.get_products_list()
        if not products:
            QMessageBox.warning(self, "警告", "没有已配置的产品")
            return
        
        # 简单对话框选择产品
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
            # 检查是否已存在
            for i in range(self.product_list.count()):
                if self.product_list.item(i).text() == selected:
                    return
            self.product_list.addItem(selected)
    
    def _remove_product(self):
        """移除选中的产品"""
        current_row = self.product_list.currentRow()
        if current_row >= 0:
            self.product_list.takeItem(current_row)
    
    def _refresh_products(self):
        """刷新产品列表"""
        # 保留当前选中的产品
        current_items = []
        for i in range(self.product_list.count()):
            current_items.append(self.product_list.item(i).text())
        
        all_products = self.config_manager.get_products_list()
        
        self.product_list.clear()
        for p in all_products:
            if p in current_items:
                self.product_list.addItem(p)
    
    def _test_email(self):
        """测试邮件连接"""
        # 临时保存邮件配置用于测试
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
        
        # 临时保存到email_reporter
        self.email_reporter.email_config = email_config
        
        success, message = self.email_reporter.test_connection()
        if success:
            # 发送测试邮件
            test_subject = "Excel批量生成工具 - 测试邮件"
            test_body = "这是一封测试邮件，证明邮件配置正确。"
            send_success = self.email_reporter.send(test_subject, test_body)
            if send_success:
                QMessageBox.information(self, "成功", "测试邮件已发送，请检查收件箱")
            else:
                QMessageBox.warning(self, "警告", "邮件发送失败，请检查收件人地址")
        else:
            QMessageBox.critical(self, "错误", f"连接失败: {message}")
    
    def _create_task(self):
        """创建Windows定时任务"""
        try:
            import subprocess
            import os
            
            # 获取程序路径
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
            
            # 构建schtasks命令
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
        """删除Windows定时任务"""
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
        """立即运行自动化任务"""
        QMessageBox.information(self, "提示", "此功能将在主窗口中执行\n请关闭此窗口后点击主窗口的 '开始生成' 按钮")