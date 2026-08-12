#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
预览对话框 - 显示将要生成的日期列表
"""

from datetime import datetime
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget,
    QPushButton, QLabel, QListWidgetItem, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt

from utils.helpers import get_weekday_name


class PreviewDialog(QDialog):
    def __init__(self, dates: list, parent=None):
        super().__init__(parent)
        self.dates = dates
        self.setWindowTitle("预览 - 将要生成的日期")
        self.setMinimumSize(400, 500)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        label = QLabel(f"共 {len(self.dates)} 个日期")
        label.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(label)

        self.list_widget = QListWidget()
        for date in self.dates:
            weekday = get_weekday_name(date)
            item_text = f"{date.strftime('%Y-%m-%d')} ({weekday})"
            item = QListWidgetItem(item_text)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)

        btn_layout = QHBoxLayout()
        btn_export = QPushButton("导出为文本")
        btn_export.clicked.connect(self._export_text)
        btn_close = QPushButton("关闭")
        btn_close.clicked.connect(self.accept)
        btn_close.setDefault(True)

        btn_layout.addWidget(btn_export)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_close)
        layout.addLayout(btn_layout)

    def _export_text(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存日期列表", "", "文本文件 (*.txt)"
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(f"共 {len(self.dates)} 个日期\n")
                    f.write("=" * 40 + "\n")
                    for date in self.dates:
                        weekday = get_weekday_name(date)
                        f.write(f"{date.strftime('%Y-%m-%d')} ({weekday})\n")
                QMessageBox.information(self, "成功", f"已导出到: {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"导出失败: {str(e)}")