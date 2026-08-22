#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
自定义标题栏组件
支持拖动窗口、双击切换最大化、窗口控制按钮
"""

from PyQt5.QtWidgets import QWidget, QHBoxLayout, QPushButton, QLabel, QDialog
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QMouseEvent


class TitleBar(QWidget):
    """自定义标题栏"""
    
    def __init__(self, parent, title: str = "Excel批量生成工具", icon: str = "⚙"):
        super().__init__(parent)
        self.parent = parent
        self.drag_pos = None
        self._is_dialog = isinstance(parent, QDialog)
        
        self.setFixedHeight(38)
        self.setObjectName("TitleBar")
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 0, 10, 0)
        layout.setSpacing(0)
        
        # ========== 左侧：图标 + 标题 ==========
        left_layout = QHBoxLayout()
        left_layout.setSpacing(8)
        
        # 图标
        self.icon_label = QLabel(icon)
        self.icon_label.setObjectName("TitleIcon")
        self.icon_label.setFixedSize(24, 24)
        self.icon_label.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(self.icon_label)
        
        # 标题文字
        self.title_label = QLabel(title)
        self.title_label.setObjectName("TitleText")
        left_layout.addWidget(self.title_label)
        left_layout.addStretch()
        
        layout.addLayout(left_layout)
        
        # ========== 右侧：窗口控制按钮 ==========
        right_layout = QHBoxLayout()
        right_layout.setSpacing(0)
        
        self.btn_minimize = QPushButton("─")
        self.btn_minimize.setFixedSize(46, 38)
        self.btn_minimize.setObjectName("TitleBtnMin")
        self.btn_minimize.clicked.connect(self.parent.showMinimized)
        right_layout.addWidget(self.btn_minimize)
        
        self.btn_maximize = QPushButton("□")
        self.btn_maximize.setFixedSize(46, 38)
        self.btn_maximize.setObjectName("TitleBtnMax")
        self.btn_maximize.clicked.connect(self._toggle_maximize)
        right_layout.addWidget(self.btn_maximize)
        
        self.btn_close = QPushButton("✕")
        self.btn_close.setFixedSize(46, 38)
        self.btn_close.setObjectName("TitleBtnClose")
        if self._is_dialog:
            self.btn_close.clicked.connect(self.parent.reject)
        else:
            self.btn_close.clicked.connect(self.parent.close)
        right_layout.addWidget(self.btn_close)
        
        layout.addLayout(right_layout)
    
    def _toggle_maximize(self):
        if self._is_dialog:
            return  # 对话框不支持最大化切换
        
        if self.parent.isMaximized():
            self.parent.showNormal()
            self.btn_maximize.setText("□")
        else:
            self.parent.showMaximized()
            self.btn_maximize.setText("❐")
    
    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            self.drag_pos = event.globalPos()
    
    def mouseMoveEvent(self, event: QMouseEvent):
        if self.drag_pos and not self.parent.isMaximized():
            delta = event.globalPos() - self.drag_pos
            self.parent.move(self.parent.pos() + delta)
            self.drag_pos = event.globalPos()
    
    def mouseReleaseEvent(self, event):
        self.drag_pos = None
    
    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton and not self._is_dialog:
            self._toggle_maximize()
    
    def set_title(self, title: str):
        """更新标题栏文字"""
        self.title_label.setText(title)
    
    def set_icon(self, icon: str):
        """更新标题栏图标"""
        self.icon_label.setText(icon)