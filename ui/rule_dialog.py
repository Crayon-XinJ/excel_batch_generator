#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
规则配置对话框 - 添加/编辑内容修改规则
"""

import uuid
import re
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QComboBox, QLineEdit, QSpinBox, QDoubleSpinBox,
    QPushButton, QGroupBox, QCheckBox, QMessageBox
)
from PyQt5.QtCore import Qt

from models.config_models import Rule


class RuleDialog(QDialog):
    def __init__(self, rule: Rule = None, sheet_names: list = None, parent=None):
        super().__init__(parent)
        self.rule = rule
        self.is_edit = rule is not None
        self.sheet_names = sheet_names or []
        self.setWindowTitle("编辑规则" if self.is_edit else "添加规则")
        self.setMinimumWidth(480)
        self.init_ui()
        if self.is_edit:
            self._load_rule()
        else:
            # 新增模式：绑定自动识别
            self.target_edit.textChanged.connect(self._auto_detect_target_type)

    def init_ui(self):
        layout = QVBoxLayout(self)
        form = QFormLayout()

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

        self.target_edit = QLineEdit()
        self.target_edit.setPlaceholderText('例如: C9-L9 或 F9,K9')
        # 编辑模式下也绑定自动识别（用户修改时触发）
        if self.is_edit:
            self.target_edit.textChanged.connect(self._auto_detect_target_type)
        form.addRow("目标位置:", self.target_edit)

        self.sheet_combo = QComboBox()
        self.sheet_combo.addItem("所有工作表", "")
        if self.sheet_names:
            for name in self.sheet_names:
                self.sheet_combo.addItem(name, name)
        form.addRow("工作表:", self.sheet_combo)

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
            '夜班检验员: 根据日期替换夜班检验员名字'
        )
        self.value_type_combo.currentTextChanged.connect(self._on_value_type_changed)
        form.addRow("值类型:", self.value_type_combo)

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

        self.enabled_check = QCheckBox("启用此规则")
        self.enabled_check.setChecked(True)
        form.addRow("", self.enabled_check)

        layout.addLayout(form)

        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("确定")
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    # ============================================================
    # 自动识别目标类型
    # ============================================================
    
    def _auto_detect_target_type(self):
        """根据输入的目标位置自动识别类型"""
        target = self.target_edit.text().strip()
        if not target:
            return
        
        # 离散单元格检测（包含逗号）
        if ',' in target:
            self._set_target_type('cells')
            return
        
        # 范围检测（包含 '-'）
        if '-' in target:
            left, right = target.split('-')
            left_match = re.match(r'([A-Z]+)(\d+)', left.strip(), re.IGNORECASE)
            right_match = re.match(r'([A-Z]+)(\d+)', right.strip(), re.IGNORECASE)
            
            if left_match and right_match:
                left_row = int(left_match.group(2))
                right_row = int(right_match.group(2))
                left_col = left_match.group(1).upper()
                right_col = right_match.group(1).upper()
                
                if left_row == right_row:
                    # 同一行 → 行范围
                    self._set_target_type('range')
                    return
                elif left_col == right_col:
                    # 同一列 → 列范围
                    self._set_target_type('column')
                    return
                else:
                    # 矩形范围 → 优先识别为行范围（用户可手动修改）
                    self._set_target_type('range')
                    return
        
        # 单个单元格检测（纯字母+数字）
        match = re.match(r'^([A-Z]+)(\d+)$', target, re.IGNORECASE)
        if match:
            self._set_target_type('cell')
            return
        
        # 无法识别，不做修改（保留用户当前选择）
    
    def _set_target_type(self, target_type: str):
        """设置目标类型下拉框的值（不触发信号，避免循环）"""
        type_map = {
            'range': '行范围 (range)',
            'column': '列范围 (column)',
            'cell': '单个单元格 (cell)',
            'cells': '离散单元格 (cells)'
        }
        text = type_map.get(target_type, '行范围 (range)')
        if self.target_type_combo.currentText() != text:
            # 使用 blockSignals 避免信号循环
            self.target_type_combo.blockSignals(True)
            self.target_type_combo.setCurrentText(text)
            self.target_type_combo.blockSignals(False)

    def _on_value_type_changed(self, text: str):
        is_random = text.startswith('随机数')
        self.min_spin.setEnabled(is_random)
        self.max_spin.setEnabled(is_random)
        self.decimals_spin.setEnabled(is_random)

    def _load_rule(self):
        type_map = {
            'range': '行范围 (range)',
            'column': '列范围 (column)',
            'cell': '单个单元格 (cell)',
            'cells': '离散单元格 (cells)'
        }
        self.target_type_combo.setCurrentText(type_map.get(self.rule.target_type, '行范围 (range)'))
        self.target_edit.setText(self.rule.target)
        if self.rule.sheet_name:
            idx = self.sheet_combo.findData(self.rule.sheet_name)
            if idx >= 0:
                self.sheet_combo.setCurrentIndex(idx)
        else:
            self.sheet_combo.setCurrentIndex(0)
        value_map = {
            'random': '随机数 (random)',
            'date': '日期替换 (date)',
            'text_with_date': '文本中的日期 (text_with_date)',
            'night_shift': '夜班检验员 (night_shift)'
        }
        self.value_type_combo.setCurrentText(value_map.get(self.rule.value_type, '随机数 (random)'))
        if self.rule.min_val is not None:
            self.min_spin.setValue(self.rule.min_val)
        if self.rule.max_val is not None:
            self.max_spin.setValue(self.rule.max_val)
        self.decimals_spin.setValue(self.rule.decimals)
        self.enabled_check.setChecked(self.rule.enabled)
        self._on_value_type_changed(self.value_type_combo.currentText())

    def get_rule(self) -> Rule:
        type_text = self.target_type_combo.currentText()
        if 'range' in type_text:
            target_type = 'range'
        elif 'column' in type_text:
            target_type = 'column'
        elif 'cell' in type_text and 'cells' not in type_text:
            target_type = 'cell'
        else:
            target_type = 'cells'

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
        if not self.target_edit.text().strip():
            QMessageBox.warning(self, "警告", "请填写目标位置")
            return
        super().accept()