#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
生成线程 - 在独立线程中执行批量生成任务

设计原因：
    1. 批量生成可能耗时较长，如果在主线程执行会阻塞界面
    2. 使用 QThread 在后台执行，通过信号与主线程通信
    3. 支持取消操作（安全停止）
    4. 支持断点续做（跳过已存在文件）
"""

import os
import time
from datetime import datetime
from typing import List

from PyQt5.QtCore import QThread, pyqtSignal

from core.excel_generator import ExcelGenerator
from core.log_manager import get_logger
from models.config_models import Rule


class GenerateThread(QThread):
    """
    生成线程
    
    信号说明：
        progress_updated: 进度更新 (当前进度, 总数, 状态消息)
        finished_signal: 完成信号 (成功数, 总数)
        error_signal: 错误信号 (错误消息)
        time_updated: 剩余时间更新 (剩余秒数)
    """
    
    progress_updated = pyqtSignal(int, int, str)
    finished_signal = pyqtSignal(int, int)
    error_signal = pyqtSignal(str)
    time_updated = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self.logger = get_logger('GenerateThread')
        
        # 生成参数
        self.dates: List[datetime] = []
        self.template_path: str = ''
        self.output_dir: str = ''
        self.rules: List[Rule] = []
        self.product_name: str = ''
        self.template_type: str = ''
        
        # 取消标志
        self._is_cancelled = False
        
        # 开始时间（用于计算剩余时间）
        self._start_time = 0

    def setup(self, dates: List[datetime], template_path: str, output_dir: str,
              rules: List[Rule], product_name: str, template_type: str):
        """设置生成参数"""
        self.dates = dates
        self.template_path = template_path
        self.output_dir = output_dir
        self.rules = rules
        self.product_name = product_name
        self.template_type = template_type
        self._is_cancelled = False

    def cancel(self):
        """取消生成（设置标志，线程会在下一个检查点停止）"""
        self.logger.info("用户请求中止生成")
        self._is_cancelled = True

    def run(self):
        """执行生成（在独立线程中运行）"""
        # ----- 参数验证 -----
        if not self.dates:
            self.error_signal.emit("没有需要生成的日期")
            self.logger.error("生成失败：日期列表为空")
            return
        
        if not os.path.exists(self.template_path):
            self.error_signal.emit(f"模板文件不存在: {self.template_path}")
            self.logger.error(f"生成失败：模板文件不存在 - {self.template_path}")
            return

        os.makedirs(self.output_dir, exist_ok=True)
        
        self.logger.info(f"开始生成，类型: {self.template_type}，共 {len(self.dates)} 个文件")
        self.logger.debug(f"  日期列表: {[d.strftime('%Y-%m-%d') for d in self.dates]}")
        self.logger.debug(f"  规则数量: {len(self.rules)}")

        # ----- 创建生成器 -----
        generator = ExcelGenerator()
        generator.set_template(self.template_path)
        generator.set_output_dir(self.output_dir)
        generator.set_rules(self.rules)
        generator.set_product_info(self.product_name, self.template_type)

        total = len(self.dates)
        success_count = 0
        self._start_time = time.time()

        # ----- 逐日生成 -----
        for i, date in enumerate(self.dates):
            if self._is_cancelled:
                self.logger.info(f"生成已取消，已处理 {i}/{total} 个文件")
                break

            filename = f"{self.product_name}_{self.template_type}检验表_{date.strftime('%Y%m%d')}.xlsx"
            output_path = os.path.join(self.output_dir, filename)

            # 断点续做：跳过已存在的文件
            if os.path.exists(output_path):
                self.logger.warning(f"文件已存在，跳过: {filename}")
                self.progress_updated.emit(i + 1, total, f"跳过已存在: {filename}")
                success_count += 1
                continue

            try:
                generator.generate(date, filename)
                success_count += 1
                self.logger.info(f"✓ 已生成: {filename}")
                self.progress_updated.emit(i + 1, total, f"已生成: {filename}")

                # 计算剩余时间
                elapsed = time.time() - self._start_time
                avg_time = elapsed / (i + 1)
                remaining_count = total - (i + 1)
                remaining_seconds = int(remaining_count * avg_time)
                self.time_updated.emit(remaining_seconds)
                self.logger.debug(f"  进度: {i+1}/{total}, 剩余时间: {remaining_seconds}秒")

            except Exception as e:
                self.logger.error(f"✗ 生成失败: {filename} - {e}")
                self.logger.exception(f"详细错误: {e}")
                self.progress_updated.emit(i + 1, total, f"生成失败: {str(e)}")

        # ----- 完成 -----
        self.logger.info(f"生成{'完成' if not self._is_cancelled else '已取消'}，成功 {success_count}/{total}")
        self.finished_signal.emit(success_count, total)