#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
规则配置对话框 - 添加/编辑内容修改规则

功能：
    1. 选择目标类型（行范围/列范围/单个单元格/离散单元格）
    2. 输入目标位置（如 C9-L9）
    3. 选择目标工作表（从模板中读取）
    4. 选择值类型（随机数/日期/文本日期/夜班检验员）
    5. 设置随机数参数（最小/最大值、小数位数）
    6. 启用/禁用规则
"""

import uuid
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QComboBox, QLineEdit, QSpinBox, QDoubleSpinBox,
    QPushButton, QGroupBox, QCheckBox, QMessageBox
)
from PyQt5.QtCore import Qt

from models.config_models import Rule


class RuleDialog(QDialog):
    """
    规则配置对话框
    
    用户通过此对话框创建或编辑一条内容修改规则。
    规则包含：目标位置、目标工作表、值类型、随机数参数等。
    """
    
    def __init__(self, rule: Rule = None, sheet_names: list = None, parent=None):
        """
        初始化对话框
        
        Args:
            rule: 已有规则（编辑模式），None 表示添加模式
            sheet_names: 模板中的工作表名称列表
            parent: 父窗口
        """
        super().__init__(parent)
        
        self.rule = rule
        self.is_edit = rule is not None      # True=编辑模式, False=添加模式
        self.sheet_names = sheet_names or []  # 工作表列表
        
        self.setWindowTitle("编辑规则" if self.is_edit else "添加规则")
        self.setMinimumWidth(480)
        
        self.init_ui()
        
        # 如果是编辑模式，加载已有规则数据
        if self.is_edit:
            self._load_rule()

    def init_ui(self):
        """创建界面"""
        layout = QVBoxLayout(self)
        form = QFormLayout()

        # ----- 目标类型（中文显示） -----
        self.target_type_combo = QComboBox()
        self.target_type_combo.addItems([
            '行范围 (range)',
            '列范围 (column)',
            '单个单元格 (cell)',
            '离散单元格 (cells)'
        ])
        self.target_type_combo.setToolTip(
            '行范围: C9-L9\n列范围: D6-D15\n单个单元格: D3\n离散单元格: F9,K9'
        )
        form.addRow("目标类型:", self.target_type_combo)

        # ----- 目标位置 -----
        self.target_edit = QLineEdit()
        self.target_edit.setPlaceholderText('例如: C9-L9 或 F9,K9')
        form.addRow("目标位置:", self.target_edit)

        # ----- 工作表选择 -----
        self.sheet_combo = QComboBox()
        self.sheet_combo.addItem("所有工作表", "")  # 空字符串表示所有工作表
        if self.sheet_names:
            for name in self.sheet_names:
                self.sheet_combo.addItem(name, name)
        form.addRow("工作表:", self.sheet_combo)

        # ----- 值类型（中文显示） -----
        self.value_type_combo = QComboBox()
        self.value_type_combo.addItems([
            '随机数 (random)',
            '日期替换 (date)',
            '文本中的日期 (text_with_date)',
            '夜班检验员 (night_shift)'
        ])
        self.value_type_combo.setToolTip(
            '随机数: 生成指定范围内的随机数\n'
            '日期替换: 将单元格中的日期替换为生成日期\n'
            '文本中的日期: 替换文本中的所有日期\n'
            '夜班检验员: 根据日期替换夜班检验员名字（仅过程表）'
        )
        self.value_type_combo.currentTextChanged.connect(self._on_value_type_changed)
        form.addRow("值类型:", self.value_type_combo)

        # ----- 随机数参数（仅当值类型为随机数时启用） -----
        param_group = QGroupBox("随机数参数")
        param_layout = QFormLayout(param_group)
        
        self.min_spin = QDoubleSpinBox()
        self.min_spin.setRange(-99999, 99999)
        self.min_spin.setDecimals(3)
        param_layout.addRow("最小值:", self.min_spin)
        
        self.max_spin = QDoubleSpinBox()
        self.max_spin.setRange(-99999, 99999)
        self.max_spin.setDecimals(3)
        param_layout.addRow("最大值:", self.max_spin)
        
        self.decimals_spin = QSpinBox()
        self.decimals_spin.setRange(0, 6)
        self.decimals_spin.setValue(2)
        param_layout.addRow("小数位数:", self.decimals_spin)
        
        form.addRow(param_group)

        # ----- 启用开关 -----
        self.enabled_check = QCheckBox("启用此规则")
        self.enabled_check.setChecked(True)
        form.addRow("", self.enabled_check)

        layout.addLayout(form)

        # ----- 按钮 -----
        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("确定")
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def _on_value_type_changed(self, text: str):
        """
        值类型变化时，控制随机数参数控件的启用状态
        
        只有选择"随机数"时，参数控件才可编辑
        """
        is_random = text.startswith('随机数')
        self.min_spin.setEnabled(is_random)
        self.max_spin.setEnabled(is_random)
        self.decimals_spin.setEnabled(is_random)

    def _load_rule(self):
        """加载已有规则数据到界面"""
        # 目标类型映射（英文 → 中文显示）
        type_map = {
            'range': '行范围 (range)',
            'column': '列范围 (column)',
            'cell': '单个单元格 (cell)',
            'cells': '离散单元格 (cells)'
        }
        self.target_type_combo.setCurrentText(type_map.get(self.rule.target_type, '行范围 (range)'))
        self.target_edit.setText(self.rule.target)
        
        # 工作表
        if self.rule.sheet_name:
            idx = self.sheet_combo.findData(self.rule.sheet_name)
            if idx >= 0:
                self.sheet_combo.setCurrentIndex(idx)
        else:
            self.sheet_combo.setCurrentIndex(0)
        
        # 值类型映射
        value_map = {
            'random': '随机数 (random)',
            'date': '日期替换 (date)',
            'text_with_date': '文本中的日期 (text_with_date)',
            'night_shift': '夜班检验员 (night_shift)'
        }
        self.value_type_combo.setCurrentText(value_map.get(self.rule.value_type, '随机数 (random)'))
        
        # 随机数参数
        if self.rule.min_val is not None:
            self.min_spin.setValue(self.rule.min_val)
        if self.rule.max_val is not None:
            self.max_spin.setValue(self.rule.max_val)
        self.decimals_spin.setValue(self.rule.decimals)
        self.enabled_check.setChecked(self.rule.enabled)
        
        self._on_value_type_changed(self.value_type_combo.currentText())

    def get_rule(self) -> Rule:
        """
        从界面获取规则对象
        
        Returns:
            Rule: 配置好的规则对象
        """
        # 解析目标类型（从中文显示中提取英文标识）
        type_text = self.target_type_combo.currentText()
        if 'range' in type_text:
            target_type = 'range'
        elif 'column' in type_text:
            target_type = 'column'
        elif 'cell' in type_text and 'cells' not in type_text:
            target_type = 'cell'
        else:
            target_type = 'cells'

        # 解析值类型
        value_text = self.value_type_combo.currentText()
        if 'random' in value_text:
            value_type = 'random'
        elif 'date' in value_text and 'text' not in value_text:
            value_type = 'date'
        elif 'text' in value_text:
            value_type = 'text_with_date'
        else:
            value_type = 'night_shift'

        return Rule(
            id=self.rule.id if self.is_edit else str(uuid.uuid4())[:8],
            target_type=target_type,
            target=self.target_edit.text().strip(),
            sheet_name=self.sheet_combo.currentData() or '',
            value_type=value_type,
            min_val=self.min_spin.value() if value_type == 'random' else None,
            max_val=self.max_spin.value() if value_type == 'random' else None,
            decimals=self.decimals_spin.value(),
            enabled=self.enabled_check.isChecked()
        )

    def accept(self):
        """验证并确认"""
        if not self.target_edit.text().strip():
            QMessageBox.warning(self, "警告", "请填写目标位置")
            return
        super().accept()