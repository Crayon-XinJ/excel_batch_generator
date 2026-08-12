#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
非工作日配置对话框
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
    def __init__(self, initial_dates: list = None, parent=None):
        super().__init__(parent)
        self.selected_dates = set()
        self.setWindowTitle("配置非工作日")
        self.setMinimumSize(500, 450)
        self.init_ui()
        if initial_dates:
            self._load_dates(initial_dates)

    def init_ui(self):
        layout = QVBoxLayout(self)

        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.calendar.clicked.connect(self._on_date_clicked)
        layout.addWidget(self.calendar)

        self.info_label = QLabel("点击日期切换非工作日标记")
        self.info_label.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(self.info_label)

        self.count_label = QLabel("已选择: 0 天")
        layout.addWidget(self.count_label)

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

    def _on_date_clicked(self, date: QDate):
        if date in self.selected_dates:
            self.selected_dates.remove(date)
        else:
            self.selected_dates.add(date)
        self._update_display()

    def _update_display(self):
        self.count_label.setText(f"已选择: {len(self.selected_dates)} 天")
        self.calendar.setDateTextFormat(QDate(), QTextCharFormat())
        highlight_fmt = QTextCharFormat()
        highlight_fmt.setBackground(QColor(135, 206, 250))
        for date in self.selected_dates:
            self.calendar.setDateTextFormat(date, highlight_fmt)
        self.calendar.update()

    def _clear_all(self):
        self.selected_dates.clear()
        self._update_display()

    def _load_dates(self, date_strings: list):
        for ds in date_strings:
            try:
                date = QDate.fromString(ds, "yyyy-MM-dd")
                if date.isValid():
                    self.selected_dates.add(date)
            except:
                pass
        self._update_display()

    def _import_dates(self):
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

    def get_selected_dates(self) -> list:
        result = []
        for date in sorted(self.selected_dates):
            result.append(date.toString("yyyy-MM-dd"))
        return result