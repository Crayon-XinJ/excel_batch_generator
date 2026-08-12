#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel批量生成工具 - 程序入口
功能：初始化PyQt5应用，创建主窗口，启动事件循环
"""

import sys
import os
import warnings

# 屏蔽Qt框架的libpng警告（不影响程序功能）
warnings.filterwarnings("ignore", category=UserWarning, module="PyQt5")
os.environ['QT_LOGGING_RULES'] = '*.debug=false;*.warning=false'

from PyQt5.QtWidgets import QApplication
from ui.main_window import MainWindow


def main():
    """程序主入口函数"""
    app = QApplication(sys.argv)
    app.setApplicationName("Excel批量生成工具")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()