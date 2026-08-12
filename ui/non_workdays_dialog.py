#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
非工作日配置对话框 - 日历选择器

功能：
    1. 显示日历，点击日期切换非工作日标记
    2. 高亮显示所有已选中的日期（浅蓝色背景）
    3. 支持导入/导出 JSON 配置
    4. 支持一键清空所有选择
    5. 隐藏周数列（更清爽）
"""

import json
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, 
    QCalendarWidget, QPushButton, QLabel,
    QMessageBox, QFileDialog
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QTextCharFormat, QColor


class NonWorkdaysDialog(QDialog):
    """
    非工作日配置对话框
    
    使用 QCalendarWidget 提供日历界面，用户点击日期切换选中状态。
    所有选中的日期会以浅蓝色背景高亮显示。
    """
    
    def __init__(self, initial_dates: list = None, parent=None):
        """
        初始化对话框
        
        Args:
            initial_dates: 已选中的日期列表（字符串格式 YYYY-MM-DD）
            parent: 父窗口
        """
        super().__init__(parent)
        
        # 存储选中日期的集合（QDate 对象）
        self.selected_dates = set()
        
        self.setWindowTitle("配置非工作日")
        self.setMinimumSize(500, 450)
        
        self.init_ui()
        
        # 如果有初始日期，加载到日历中
        if initial_dates:
            self._load_dates(initial_dates)

    def init_ui(self):
        """创建界面"""
        layout = QVBoxLayout(self)

        # ----- 日历控件 -----
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)                      # 显示网格
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)  # 隐藏周数列
        self.calendar.clicked.connect(self._on_date_clicked)   # 点击切换
        layout.addWidget(self.calendar)

        # ----- 提示信息 -----
        self.info_label = QLabel("点击日期切换非工作日标记")
        self.info_label.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(self.info_label)

        # ----- 统计信息 -----
        self.count_label = QLabel("已选择: 0 天")
        layout.addWidget(self.count_label)

        # ----- 按钮区域 -----
        btn_layout = QHBoxLayout()
        
        btn_clear = QPushButton("清空所有")
        btn_clear.clicked.connect(self._clear_all)
        
        btn_import = QPushButton("导入")
        btn_import.clicked.connect(self._import_dates)
        
        btn_export = QPushButton("导出")
        btn_export.clicked.connect(self._export_dates)
        
        btn_ok = QPushButton("确定")
        btn_ok.clicked.connect(self.accept)
        
        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addWidget(btn_clear)
        btn_layout.addWidget(btn_import)
        btn_layout.addWidget(btn_export)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    # ============================================================
    # 核心交互方法
    # ============================================================
    
    def _on_date_clicked(self, date: QDate):
        """
        点击日期时的处理
        
        如果日期已在选中集合中，则移除（取消选中）
        否则添加到选中集合（选中）
        """
        if date in self.selected_dates:
            self.selected_dates.remove(date)
        else:
            self.selected_dates.add(date)
        self._update_display()

    def _update_display(self):
        """
        更新界面显示
        
        1. 更新统计标签
        2. 清除所有日期的格式
        3. 为所有选中的日期设置浅蓝色背景
        """
        self.count_label.setText(f"已选择: {len(self.selected_dates)} 天")
        
        # 清除所有日期的格式（传入无效日期和空格式）
        self.calendar.setDateTextFormat(QDate(), QTextCharFormat())
        
        # 为所有选中的日期设置高亮
        highlight_fmt = QTextCharFormat()
        highlight_fmt.setBackground(QColor(135, 206, 250))  # 浅蓝色
        for date in self.selected_dates:
            self.calendar.setDateTextFormat(date, highlight_fmt)
        
        self.calendar.update()

    def _clear_all(self):
        """清空所有选中的日期"""
        self.selected_dates.clear()
        self._update_display()

    # ============================================================
    # 数据导入/导出
    # ============================================================
    
    def _load_dates(self, date_strings: list):
        """
        从字符串列表加载日期
        
        Args:
            date_strings: ["2026-06-07", "2026-06-14", ...]
        """
        for ds in date_strings:
            try:
                date = QDate.fromString(ds, "yyyy-MM-dd")
                if date.isValid():
                    self.selected_dates.add(date)
            except:
                pass
        self._update_display()

    def _import_dates(self):
        """从JSON文件导入非工作日列表"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "导入非工作日", "", "JSON文件 (*.json)"
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    dates = data.get('non_workdays', [])
                    self._load_dates(dates)
                    QMessageBox.information(self, "成功", f"已导入 {len(dates)} 天")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"导入失败: {str(e)}")

    def _export_dates(self):
        """导出非工作日列表为JSON文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出非工作日", "", "JSON文件 (*.json)"
        )
        if file_path:
            try:
                date_strings = self.get_selected_dates()
                data = {'non_workdays': date_strings}
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                QMessageBox.information(self, "成功", f"已导出到: {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"导出失败: {str(e)}")

    # ============================================================
    # 获取结果
    # ============================================================
    
    def get_selected_dates(self) -> list:
        """
        获取选中的日期列表
        
        Returns:
            list: 日期字符串列表 ["2026-06-07", "2026-06-14", ...]
        """
        result = []
        for date in sorted(self.selected_dates):
            result.append(date.toString("yyyy-MM-dd"))
        return result