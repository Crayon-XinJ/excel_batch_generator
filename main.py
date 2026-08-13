#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel批量生成工具 - 程序入口
支持命令行参数：
  --auto       : 以自动化模式运行
  --no-gui     : 无界面模式（需配合 --auto 使用）
  --verbose    : 详细日志
"""

import sys
import os
import warnings
import argparse

# 屏蔽Qt框架的libpng警告（不影响程序功能）
warnings.filterwarnings("ignore", category=UserWarning, module="PyQt5")
os.environ['QT_LOGGING_RULES'] = '*.debug=false;*.warning=false'

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer
from ui.main_window import MainWindow


def main():
    """程序主入口函数"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='Excel批量生成工具')
    parser.add_argument('--auto', action='store_true', help='以自动化模式运行')
    parser.add_argument('--no-gui', action='store_true', help='无界面模式（需配合 --auto 使用）')
    parser.add_argument('--verbose', action='store_true', help='详细日志输出')
    args = parser.parse_args()
    
    # 如果指定了详细日志，设置日志级别为 DEBUG
    if args.verbose:
        import logging
        logging.getLogger('ExcelBatchGenerator').setLevel(logging.DEBUG)
    
    # 创建应用程序
    app = QApplication(sys.argv)
    app.setApplicationName("Excel批量生成工具")
    
    # 创建主窗口
    window = MainWindow()
    
    # 如果是无界面自动化模式，不显示窗口
    if args.auto and args.no_gui:
        # 窗口已创建但未显示，直接执行自动化任务
        # 使用单次定时器让事件循环先运行一次，确保初始化完成
        QTimer.singleShot(100, lambda: window._run_auto_mode())
        # 进入事件循环，自动化完成后会调用 QApplication.quit() 退出
        sys.exit(app.exec_())
    else:
        # 正常显示窗口
        window.show()
        sys.exit(app.exec_())


if __name__ == '__main__':
    main()