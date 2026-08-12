#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
主窗口模块 - Excel批量生成工具的核心界面

功能概述：
    1. 模板文件加载（浏览/拖入）
    2. 产品信息管理（名称、类型）
    3. 生产工期配置（多段日期范围）
    4. 非工作日配置（调用日历对话框）
    5. 内容修改规则管理（增删改查）
    6. 预览与生成控制（预览、开始、中止）
    7. 进度显示（进度条、剩余时间）
    8. 配置自动保存/加载
"""

import os
import sys
from datetime import datetime, timedelta

# ============================================================
# PyQt5 控件导入
# ============================================================
from PyQt5.QtWidgets import (
    QMainWindow,          # 主窗口基类
    QWidget,              # 窗口部件基类
    QVBoxLayout,          # 垂直布局
    QHBoxLayout,          # 水平布局
    QPushButton,          # 按钮
    QLabel,               # 标签
    QLineEdit,            # 单行文本输入
    QDateEdit,            # 日期选择器
    QComboBox,            # 下拉选择框
    QListWidget,          # 列表控件
    QListWidgetItem,      # 列表项
    QFileDialog,          # 文件对话框
    QProgressBar,         # 进度条
    QGroupBox,            # 分组框
    QMessageBox,          # 消息对话框
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QDragEnterEvent, QDropEvent

# ============================================================
# 内部模块导入
# ============================================================
from core.config_manager import ConfigManager
from core.date_calculator import DateCalculator
from core.generate_thread import GenerateThread
from core.log_manager import get_logger
from models.config_models import Rule, ProductConfig
from ui.non_workdays_dialog import NonWorkdaysDialog
from ui.rule_dialog import RuleDialog
from ui.preview_dialog import PreviewDialog


class MainWindow(QMainWindow):
    """
    主窗口类 - 继承自 QMainWindow
    
    属性说明：
        config_manager: 配置管理器（读写JSON文件）
        date_calculator: 日期计算器（计算目标日期）
        date_ranges: 生产工期列表 [(start, end), ...]
        rules: 内容修改规则列表 [Rule, ...]
        product_name: 当前产品名称
        template_type: 当前模板类型（首件/过程/成品）
        template_path: 模板文件路径
        sheet_names: 模板中的工作表名称列表
        thread: 生成线程实例
        logger: 日志记录器
    """
    
    def __init__(self):
        """初始化主窗口"""
        super().__init__()
        
        # ---------- 窗口基本设置 ----------
        self.setWindowTitle("Excel批量生成工具")
        self.setMinimumSize(1000, 750)
        
        # ---------- 日志记录器 ----------
        self.logger = get_logger('MainWindow')
        self.logger.info("程序启动")

        # ---------- 配置管理器 ----------
        # 负责读写JSON配置文件
        self.config_manager = ConfigManager()

        # ---------- 日期计算器 ----------
        # 负责根据规则计算目标日期
        self.date_calculator = DateCalculator()
        
        # 加载已保存的非工作日配置
        non_workdays = self.config_manager.load_non_workdays()
        if non_workdays:
            self.date_calculator.set_non_workdays(non_workdays)

        # ---------- 数据容器 ----------
        self.date_ranges = []          # 生产工期列表
        self.rules = []               # 内容修改规则列表
        self.product_name = ''         # 产品名称
        self.template_type = '过程'     # 模板类型（默认过程）
        self.template_path = ''        # 模板文件路径
        self.sheet_names = []          # 工作表名称列表
        self.thread = None             # 生成线程

        # ---------- 创建界面 ----------
        self.init_ui()
        
        # ---------- 加载配置 ----------
        self._load_config()

    # ============================================================
    # 界面创建
    # ============================================================
    
    def init_ui(self):
        """
        创建用户界面
        
        布局结构：
            ┌─────────────────────────────────────────────────┐
            │  模板文件: [______________] [浏览...]           │ ← 模板加载区
            ├─────────────────────────────────────────────────┤
            │  产品名称: [__________]  模板类型: [首件▼]      │ ← 产品信息区
            ├─────────────────────────────────────────────────┤
            │  生产工期                                       │ ← 日期范围区
            │  起始: [____] 结束: [____] [添加] [清空所有]   │
            │  ├── 2026-06-11 ~ 2026-06-18  [×]             │
            │  └── 2026-07-02 ~ 2026-08-11  [×]             │
            ├─────────────────────────────────────────────────┤
            │  非工作日配置    │  内容修改规则                │ ← 中间区
            │  2026-06-07     │  ✓ [range] C9-L9 -> random  │
            │  2026-06-14     │  ✓ [cell] D3 -> date        │
            │  2026-06-19     │  ✗ [cell] A18 -> night_shift│
            │  [配置非工作日]  │  [+添加] [编辑] [删除]      │
            ├─────────────────────────────────────────────────┤
            │  [预览]  [开始生成]  [中止]  ████░░ 80%  剩余  │ ← 操作区
            └─────────────────────────────────────────────────┘
        """
        # 中央部件（所有控件的容器）
        central = QWidget()
        self.setCentralWidget(central)
        
        # 主布局（垂直方向）
        main_layout = QVBoxLayout(central)

        # ============================================================
        # 1. 模板文件加载区域
        # ============================================================
        template_group = QGroupBox("模板文件")
        template_layout = QHBoxLayout()

        # 模板路径显示框（只读，支持拖入）
        self.template_path_edit = QLineEdit()
        self.template_path_edit.setPlaceholderText("选择或拖入模板文件")
        self.template_path_edit.setReadOnly(True)
        self.template_path_edit.setAcceptDrops(True)
        self.template_path_edit.dragEnterEvent = self._drag_enter
        self.template_path_edit.dropEvent = self._drop

        # 浏览按钮
        btn_browse = QPushButton("浏览...")
        btn_browse.clicked.connect(self._browse_template)

        template_layout.addWidget(self.template_path_edit)
        template_layout.addWidget(btn_browse)
        template_group.setLayout(template_layout)
        main_layout.addWidget(template_group)

        # ============================================================
        # 2. 产品信息区域
        # ============================================================
        info_layout = QHBoxLayout()

        # 产品名称输入框
        info_layout.addWidget(QLabel("产品名称:"))
        self.product_name_edit = QLineEdit()
        self.product_name_edit.setPlaceholderText("自动从文件名提取")
        self.product_name_edit.setFixedWidth(200)
        info_layout.addWidget(self.product_name_edit)

        info_layout.addSpacing(20)

        # 模板类型下拉框
        info_layout.addWidget(QLabel("模板类型:"))
        self.template_type_combo = QComboBox()
        self.template_type_combo.addItems(['首件', '过程', '成品'])
        self.template_type_combo.setFixedWidth(100)
        self.template_type_combo.currentTextChanged.connect(self._on_template_type_changed)
        info_layout.addWidget(self.template_type_combo)

        info_layout.addStretch()
        main_layout.addLayout(info_layout)

        # ============================================================
        # 3. 生产工期管理区域
        # ============================================================
        range_group = QGroupBox("生产工期")
        range_layout = QVBoxLayout()

        # 添加日期范围的控件行
        add_range_layout = QHBoxLayout()
        add_range_layout.addWidget(QLabel("起始:"))

        # 起始日期选择器（带日历弹出）
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.setFixedWidth(120)
        add_range_layout.addWidget(self.start_date_edit)

        add_range_layout.addWidget(QLabel("结束:"))

        # 结束日期选择器（带日历弹出）
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.setFixedWidth(120)
        add_range_layout.addWidget(self.end_date_edit)

        # 添加按钮
        btn_add_range = QPushButton("添加")
        btn_add_range.clicked.connect(self._add_date_range)
        add_range_layout.addWidget(btn_add_range)

        # 清空按钮
        btn_clear_ranges = QPushButton("清空所有")
        btn_clear_ranges.clicked.connect(self._clear_date_ranges)
        add_range_layout.addWidget(btn_clear_ranges)

        add_range_layout.addStretch()
        range_layout.addLayout(add_range_layout)

        # 日期范围列表（显示已添加的工期）
        self.range_list = QListWidget()
        self.range_list.setMaximumHeight(120)
        range_layout.addWidget(self.range_list)

        range_group.setLayout(range_layout)
        main_layout.addWidget(range_group)

        # ============================================================
        # 4. 中间区域：非工作日 + 规则（横向分栏）
        # ============================================================
        mid_layout = QHBoxLayout()

        # ----- 4.1 非工作日配置 -----
        nonwork_group = QGroupBox("非工作日配置")
        nonwork_layout = QVBoxLayout()

        # 非工作日列表（纵向显示，每个日期一行）
        self.nonwork_list = QListWidget()
        self.nonwork_list.setMaximumHeight(100)
        self.nonwork_list.setStyleSheet("QListWidget::item { padding: 2px; }")
        nonwork_layout.addWidget(self.nonwork_list)

        # 配置非工作日按钮
        btn_nonwork = QPushButton("配置非工作日")
        btn_nonwork.clicked.connect(self._configure_non_workdays)
        nonwork_layout.addWidget(btn_nonwork)

        # 导入/导出非工作日按钮
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

        # ----- 4.2 内容修改规则 -----
        rule_group = QGroupBox("内容修改规则")
        rule_layout = QVBoxLayout()

        # 规则列表
        self.rule_list = QListWidget()
        self.rule_list.setMaximumHeight(150)
        rule_layout.addWidget(self.rule_list)

        # 规则操作按钮
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

        # ============================================================
        # 5. 底部操作区域
        # ============================================================
        action_layout = QHBoxLayout()

        # 预览按钮
        btn_preview = QPushButton("预览")
        btn_preview.clicked.connect(self._preview)
        action_layout.addWidget(btn_preview)

        # 开始生成按钮（绿色高亮）
        btn_generate = QPushButton("开始生成")
        btn_generate.clicked.connect(self._start_generate)
        btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        action_layout.addWidget(btn_generate)

        # 中止按钮（初始禁用）
        btn_cancel = QPushButton("中止")
        btn_cancel.clicked.connect(self._cancel_generate)
        btn_cancel.setEnabled(False)
        action_layout.addWidget(btn_cancel)

        action_layout.addStretch()

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedWidth(200)
        action_layout.addWidget(self.progress_bar)

        # 剩余时间标签
        self.time_label = QLabel("剩余时间: --")
        action_layout.addWidget(self.time_label)

        # 状态标签
        self.status_label = QLabel("就绪")
        action_layout.addWidget(self.status_label)

        main_layout.addLayout(action_layout)

        # 初始化非工作日显示
        self._update_nonwork_display()

    # ============================================================
    # 非工作日显示更新
    # ============================================================
    
    def _update_nonwork_display(self):
        """
        更新非工作日列表显示
        
        从配置文件加载非工作日列表，在 QListWidget 中纵向显示
        """
        dates = self.config_manager.load_non_workdays()
        self.nonwork_list.clear()
        if dates:
            for d in sorted(dates):
                self.nonwork_list.addItem(d)
        else:
            self.nonwork_list.addItem("未配置非工作日")

    # ============================================================
    # 模板加载相关（拖入/浏览）
    # ============================================================
    
    def _drag_enter(self, event: QDragEnterEvent):
        """拖入事件：接受文件拖入"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def _drop(self, event: QDropEvent):
        """拖放事件：处理文件拖入"""
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith(('.xlsx', '.xls')):
                self._load_template(file_path)

    def _browse_template(self):
        """浏览按钮：打开文件选择对话框"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择模板文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self._load_template(file_path)

    def _load_template(self, file_path: str):
        """
        加载模板文件
        
        执行流程：
            1. 记录模板路径
            2. 从文件名提取产品名称和模板类型
            3. 读取模板中的工作表名称
            4. 加载对应产品的配置
        """
        """加载模板文件"""
        self.template_path = file_path
        self.template_path_edit.setText(file_path)

        # 提取产品名称
        filename = os.path.basename(file_path)
        name = os.path.splitext(filename)[0]

        # 定义要移除的后缀列表（支持“模板”和“模版”两种写法）
        suffixes_to_remove = [
            '首件检验表模板',
            '首件检验表模版',
            '过程检验表模板',
            '过程检验表模版',
            '成品检验表模板',
            '成品检验表模版',
            '首件检验表',
            '过程检验表',
            '成品检验表',
        ]

        matched_suffix = None
        for suffix in suffixes_to_remove:
            if name.endswith(suffix):
                matched_suffix = suffix
                self.template_type = suffix.replace('检验表模板', '').replace('检验表模版', '').replace('检验表', '')
                break

        if matched_suffix:
            # 去掉匹配的后缀，剩余部分作为产品名
            self.product_name = name[:-len(matched_suffix)].rstrip('_')
            # 确保产品名不为空
            if not self.product_name:
                self.product_name = name
        else:
            # 如果无法匹配，使用整个文件名作为产品名
            self.product_name = name
            # 尝试从文件名中猜测模板类型
            if '首件' in name:
                self.template_type = '首件'
            elif '过程' in name:
                self.template_type = '过程'
            elif '成品' in name:
                self.template_type = '成品'
            else:
                self.template_type = '过程'  # 默认

        self.product_name_edit.setText(self.product_name)
        self.template_type_combo.setCurrentText(self.template_type)

        # ----- 读取工作表名称 -----
        # 用于规则配置时选择目标工作表
        self.sheet_names = self._get_sheet_names(file_path)
        self.logger.debug(f"  工作表列表: {self.sheet_names}")

        # ----- 加载产品配置 -----
        self._load_config()

    def _get_sheet_names(self, file_path: str) -> list:
        """
        使用win32com读取Excel模板的工作表名称
        
        注意：此方法会启动Excel应用程序，然后关闭
        """
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
        """
        加载产品配置（规则 + 日期范围）
        
        配置按 产品名 + 模板类型 分别存储
        切换模板类型时会自动重新加载
        """
        product_name = self.product_name_edit.text().strip()
        template_type = self.template_type_combo.currentText()
        if not product_name:
            return

        self.product_name = product_name
        self.template_type = template_type

        self.logger.info(f"加载配置: {product_name} - {template_type}")

        # ----- 加载规则 -----
        config = self.config_manager.load_product_config(product_name, template_type)
        if config:
            self.rules = config.rules
            
            # ----- 加载日期范围 -----
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
        
        self.logger.debug(f"  规则数: {len(self.rules)}")
        self.logger.debug(f"  日期范围: {len(self.date_ranges)} 段")
        
        self._update_rule_list()
        self._update_nonwork_display()

    def _save_rules(self):
        """
        保存产品配置（规则 + 日期范围）
        
        每次修改规则或日期范围时自动调用
        """
        product_name = self.product_name_edit.text().strip()
        template_type = self.template_type_combo.currentText()
        if not product_name:
            return
            
        self.logger.info(f"保存配置: {product_name} - {template_type}")
        self.logger.debug(f"  规则数: {len(self.rules)}")
        self.logger.debug(f"  日期范围: {len(self.date_ranges)} 段")

        # 转换日期范围为字符串列表
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
        """模板类型切换时重新加载配置"""
        self.template_type = template_type
        self._load_config()

    # ============================================================
    # 日期范围管理
    # ============================================================
    
    def _add_date_range(self):
        """添加生产工期"""
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
        
        self.logger.debug(f"添加日期范围: {start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")
        
        # 自动保存
        self._save_rules()

    def _clear_date_ranges(self):
        """清空所有生产工期"""
        self.date_ranges.clear()
        self.range_list.clear()
        self.logger.debug("清空所有日期范围")
        self._save_rules()

    # ============================================================
    # 非工作日管理
    # ============================================================
    
    def _configure_non_workdays(self):
        """打开非工作日配置对话框"""
        current_dates = self.config_manager.load_non_workdays()
        dialog = NonWorkdaysDialog(current_dates, self)
        if dialog.exec_():
            dates = dialog.get_selected_dates()
            self.config_manager.save_non_workdays(dates)
            self.date_calculator.set_non_workdays(dates)
            self._update_nonwork_display()
            self.logger.info(f"更新非工作日: {len(dates)} 天")

    def _import_non_workdays(self):
        """导入非工作日配置（JSON文件）"""
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
        """导出非工作日配置为JSON文件"""
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
        """更新规则列表显示"""
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
        """添加新规则"""
        if not self.sheet_names:
            QMessageBox.warning(self, "警告", "请先加载模板文件以获取工作表列表")
            return
        dialog = RuleDialog(sheet_names=self.sheet_names, parent=self)
        if dialog.exec_():
            rule = dialog.get_rule()
            self.rules.append(rule)
            self._save_rules()
            self._update_rule_list()
            self.logger.info(f"添加规则: {rule.id} ({rule.target} -> {rule.value_type})")

    def _edit_rule(self):
        """编辑选中规则"""
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
            self.logger.info(f"编辑规则: {new_rule.id}")

    def _delete_rule(self):
        """删除选中规则"""
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
        """导入规则配置（JSON文件）"""
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
        """导出规则配置为JSON文件"""
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
        """
        计算目标日期列表
        
        根据当前模板类型（首件/过程/成品）和非工作日配置，
        计算需要生成的所有日期。
        """
        template_type = self.template_type_combo.currentText()
        if not self.date_ranges:
            QMessageBox.warning(self, "警告", "请先添加生产工期")
            return []

        dates = self.date_calculator.preview_dates(template_type, self.date_ranges)
        self.logger.debug(f"计算目标日期: {len(dates)} 天")
        return dates

    def _preview(self):
        """打开预览对话框"""
        dates = self._get_target_dates()
        if not dates:
            QMessageBox.information(self, "提示", "没有需要生成的日期")
            return
        dialog = PreviewDialog(dates, self)
        dialog.exec_()

    def _start_generate(self):
        """开始批量生成"""
        # ----- 验证 -----
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

        # 选择输出目录
        output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if not output_dir:
            return

        # ----- 禁用控件 -----
        self._set_controls_enabled(False)
        self.progress_bar.setValue(0)
        self.time_label.setText("剩余时间: 计算中...")
        self.status_label.setText("生成中...")

        self.logger.info(f"开始生成，类型: {self.template_type}，共 {len(dates)} 个文件")
        self.logger.debug(f"  日期列表: {[d.strftime('%Y-%m-%d') for d in dates]}")
        self.logger.debug(f"  规则数: {len(self.rules)}")

        # ----- 创建并启动生成线程 -----
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
        """中止生成"""
        if self.thread and self.thread.isRunning():
            self.thread.cancel()
            self.status_label.setText("正在中止...")
            self.logger.info("用户请求中止生成")

    # ============================================================
    # 生成线程信号处理
    # ============================================================
    
    def _on_progress_updated(self, current: int, total: int, message: str):
        """进度更新"""
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.status_label.setText(message)

    def _on_time_updated(self, remaining: int):
        """剩余时间更新"""
        if remaining > 0:
            minutes = remaining // 60
            seconds = remaining % 60
            self.time_label.setText(f"剩余时间: {minutes}分{seconds}秒")
        else:
            self.time_label.setText("剩余时间: 即将完成")

    def _on_generate_finished(self, success_count: int, total: int):
        """生成完成"""
        self._set_controls_enabled(True)
        self.progress_bar.setValue(total)
        self.status_label.setText(f"完成！成功 {success_count}/{total}")
        self.time_label.setText("剩余时间: --")
        
        self.logger.info(f"生成完成，成功 {success_count}/{total}")
        QMessageBox.information(self, "完成", f"成功生成 {success_count} 个文件")
        self.thread = None

    def _on_generate_error(self, error: str):
        """生成出错"""
        self._set_controls_enabled(True)
        self.status_label.setText(f"错误: {error}")
        self.time_label.setText("剩余时间: --")
        self.logger.error(f"生成错误: {error}")
        QMessageBox.critical(self, "错误", error)
        self.thread = None

    # ============================================================
    # 控件状态控制
    # ============================================================
    
    def _set_controls_enabled(self, enabled: bool):
        """
        设置控件启用/禁用状态
        
        生成过程中禁用大部分控件，防止用户误操作
        """
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