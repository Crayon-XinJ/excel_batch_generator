#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
暗黑风格样式表 - 统一管理所有窗口的暗黑样式
"""

DARK_STYLE = """
/* ============================================================
   暗黑风格样式表 - 全局
   ============================================================ */

/* ----- 主窗口容器 ----- */
QMainWindow, QDialog {
    background-color: #1e1e1e;
}

/* ----- 自定义标题栏 ----- */
#TitleBar {
    background-color: #252526;
}

#TitleBar QLabel {
    color: #cccccc;
}

#TitleIcon {
    color: #0078d4;
    font-size: 16px;
}

#TitleText {
    font-size: 12px;
    font-weight: bold;
}

/* ----- 标题栏按钮 ----- */
#TitleBtnMin, #TitleBtnMax, #TitleBtnClose {
    background-color: transparent;
    border: none;
    color: #cccccc;
    font-size: 14px;
    font-weight: normal;
    font-family: "Segoe UI", "Microsoft YaHei";
}
#TitleBtnMin:hover {
    background-color: #3a3a3a;
}
#TitleBtnMax:hover {
    background-color: #3a3a3a;
}
#TitleBtnClose:hover {
    background-color: #e81123;
    color: white;
}

/* ----- 分组框 ----- */
QGroupBox {
    color: #cccccc;
    border: 1px solid #3a3a3a;
    border-radius: 6px;
    margin-top: 10px;
    padding-top: 10px;
    background-color: #252526;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 8px;
    color: #cccccc;
    font-weight: bold;
}

/* ----- 标签 ----- */
QLabel {
    color: #cccccc;
}

/* ----- 按钮 ----- */
QPushButton {
    padding: 6px 16px;
    border-radius: 4px;
    font-weight: bold;
    font-size: 12px;
    border: none;
    background-color: #3a3a3a;
    color: #cccccc;
}
QPushButton:hover {
    background-color: #4a4a4a;
}
QPushButton:pressed {
    background-color: #2a2a2a;
}
QPushButton:disabled {
    color: #666666;
    background-color: #2a2a2a;
}

/* ----- 主要按钮（蓝色） ----- */
QPushButton[objectName="PrimaryBtn"] {
    background-color: #0078d4;
    color: white;
}
QPushButton[objectName="PrimaryBtn"]:hover {
    background-color: #1a8be0;
}
QPushButton[objectName="PrimaryBtn"]:disabled {
    background-color: #2a4a6a;
    color: #666666;
}

/* ----- 成功按钮（绿色） ----- */
QPushButton[objectName="SuccessBtn"] {
    background-color: #2ea043;
    color: white;
}
QPushButton[objectName="SuccessBtn"]:hover {
    background-color: #3fb950;
}
QPushButton[objectName="SuccessBtn"]:disabled {
    background-color: #2a4a3a;
    color: #666666;
}

/* ----- 危险按钮（红色） ----- */
QPushButton[objectName="DangerBtn"] {
    background-color: #d32f2f;
    color: white;
}
QPushButton[objectName="DangerBtn"]:hover {
    background-color: #e53935;
}
QPushButton[objectName="DangerBtn"]:disabled {
    background-color: #4a2a2a;
    color: #666666;

/* ----- 输入框 ----- */
QLineEdit, QTextEdit, QComboBox, QDateEdit, QSpinBox, QDoubleSpinBox {
    background-color: #2d2d2d;
    color: #cccccc;
    border: 1px solid #3a3a3a;
    border-radius: 4px;
    padding: 4px 8px;
}
QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QDateEdit:focus {
    border: 1px solid #0078d4;
}
QLineEdit:disabled, QTextEdit:disabled {
    color: #666666;
    background-color: #222222;
}

/* ----- 下拉框 ----- */
QComboBox::drop-down {
    border: none;
    width: 20px;
}
QComboBox::down-arrow {
    image: none;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 6px solid #888888;
    margin-right: 8px;
}
QComboBox QAbstractItemView {
    background-color: #2d2d2d;
    color: #cccccc;
    border: 1px solid #3a3a3a;
    selection-background-color: #0078d4;
}

/* ----- 列表控件 ----- */
QListWidget, QTableWidget {
    background-color: #1e1e1e;
    color: #cccccc;
    border: 1px solid #3a3a3a;
    border-radius: 4px;
    gridline-color: #2d2d2d;
    outline: none;
}
QListWidget::item, QTableWidget::item {
    padding: 4px 8px;
}
QListWidget::item:selected, QTableWidget::item:selected {
    background-color: #0078d4;
    color: white;
}
QListWidget::item:hover, QTableWidget::item:hover {
    background-color: #2d2d2d;
}
QTableWidget::item:selected:hover {
    background-color: #1a8be0;
}
QHeaderView::section {
    background-color: #252526;
    color: #cccccc;
    padding: 4px 8px;
    border: none;
    border-right: 1px solid #3a3a3a;
    border-bottom: 1px solid #3a3a3a;
    font-weight: bold;
}

/* ----- 进度条 ----- */
QProgressBar {
    background-color: #2d2d2d;
    border: 1px solid #3a3a3a;
    border-radius: 4px;
    height: 20px;
    text-align: center;
    color: #cccccc;
}
QProgressBar::chunk {
    background-color: #0078d4;
    border-radius: 3px;
}

/* ----- 滚动条 ----- */
QScrollBar:vertical {
    background-color: #1e1e1e;
    width: 12px;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background-color: #3a3a3a;
    border-radius: 4px;
    min-height: 20px;
}
QScrollBar::handle:vertical:hover {
    background-color: #4a4a4a;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QScrollBar:horizontal {
    background-color: #1e1e1e;
    height: 12px;
    border-radius: 4px;
}
QScrollBar::handle:horizontal {
    background-color: #3a3a3a;
    border-radius: 4px;
    min-width: 20px;
}
QScrollBar::handle:horizontal:hover {
    background-color: #4a4a4a;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
}

/* ----- 选项卡 ----- */
QTabWidget::pane {
    border: 1px solid #3a3a3a;
    border-radius: 4px;
    background-color: #1e1e1e;
}
QTabBar::tab {
    background-color: #252526;
    color: #888888;
    padding: 6px 16px;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
}
QTabBar::tab:selected {
    background-color: #1e1e1e;
    color: #cccccc;
    border-bottom: 2px solid #0078d4;
}
QTabBar::tab:hover {
    color: #ffffff;
}

/* ----- 日历控件 ----- */
QCalendarWidget {
    background-color: #1e1e1e;
    color: #cccccc;
}
QCalendarWidget QTableView {
    background-color: #2d2d2d;
    selection-background-color: #0078d4;
}
QCalendarWidget QHeaderView::section {
    background-color: #252526;
    color: #cccccc;
}

/* ----- 消息框 ----- */
QMessageBox {
    background-color: #1e1e1e;
    color: #cccccc;
}
QMessageBox QLabel {
    color: #cccccc;
}
QMessageBox QPushButton {
    min-width: 80px;
}

/* ----- 菜单 ----- */
QMenuBar {
    background-color: #252526;
    color: #cccccc;
}
QMenuBar::item:selected {
    background-color: #3a3a3a;
}
QMenu {
    background-color: #252526;
    color: #cccccc;
    border: 1px solid #3a3a3a;
}
QMenu::item:selected {
    background-color: #0078d4;
}

/* ----- 状态栏 ----- */
QStatusBar {
    background-color: #252526;
    color: #888888;
    border-top: 1px solid #3a3a3a;
}
"""