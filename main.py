#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel批量生成工具 - 程序入口
功能：初始化PyQt5应用，创建主窗口，启动事件循环
"""

import sys
import os
import warnings

# ============================================================
# 屏蔽Qt框架的libpng警告（不影响程序功能）
# 原因：Qt加载PNG图标时，如果图标包含不标准的颜色配置文件（iCCP），
#       会输出警告信息，但图标仍能正常显示。
#       这些警告对程序运行无影响，屏蔽可让控制台输出更清爽。
# ============================================================
warnings.filterwarnings("ignore", category=UserWarning, module="PyQt5")
os.environ['QT_LOGGING_RULES'] = '*.debug=false;*.warning=false'

# 导入PyQt5核心模块
from PyQt5.QtWidgets import QApplication

# 导入主窗口类
from ui.main_window import MainWindow


def main():
    """
    程序主入口函数
    
    执行流程：
        1. 创建 QApplication 实例（每个Qt程序有且只有一个）
        2. 设置应用程序名称（用于任务栏显示）
        3. 创建主窗口实例
        4. 显示主窗口
        5. 进入事件循环（等待用户操作）
    """
    # 创建应用程序实例
    # sys.argv 是命令行参数列表，Qt需要解析其中的参数
    app = QApplication(sys.argv)
    
    # 设置应用程序显示名称
    app.setApplicationName("Excel批量生成工具")
    
    # 创建主窗口
    window = MainWindow()
    
    # 显示主窗口
    window.show()
    
    # 进入事件循环，程序在此等待用户操作
    # 当所有窗口关闭时，exec_() 返回0，程序退出
    sys.exit(app.exec_())


# ============================================================
# 标准Python入口判断
# 当直接运行此文件时（而非作为模块导入），执行main()
# ============================================================
if __name__ == '__main__':
    main()