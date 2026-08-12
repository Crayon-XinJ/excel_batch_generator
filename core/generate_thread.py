#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
生成线程 - 在独立线程中执行批量生成任务
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
    """生成线程"""
    
    progress_updated = pyqtSignal(int, int, str)
    finished_signal = pyqtSignal(int, int)
    error_signal = pyqtSignal(str)
    time_updated = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self.logger = get_logger('GenerateThread')
        
        self.dates: List[datetime] = []
        self.template_path: str = ''
        self.output_dir: str = ''
        self.rules: List[Rule] = []
        self.product_name: str = ''
        self.template_type: str = ''
        
        self._is_cancelled = False
        self._start_time = 0

    def setup(self, dates: List[datetime], template_path: str, output_dir: str,
              rules: List[Rule], product_name: str, template_type: str):
        self.dates = dates
        self.template_path = template_path
        self.output_dir = output_dir
        self.rules = rules
        self.product_name = product_name
        self.template_type = template_type
        self._is_cancelled = False

    def cancel(self):
        self.logger.info("用户请求中止生成")
        self._is_cancelled = True

    def run(self):
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

        generator = ExcelGenerator()
        generator.set_template(self.template_path)
        generator.set_output_dir(self.output_dir)
        generator.set_rules(self.rules)
        generator.set_product_info(self.product_name, self.template_type)

        total = len(self.dates)
        success_count = 0
        self._start_time = time.time()

        for i, date in enumerate(self.dates):
            if self._is_cancelled:
                self.logger.info(f"生成已取消，已处理 {i}/{total} 个文件")
                break

            filename = f"{self.product_name}_{self.template_type}检验表_{date.strftime('%Y%m%d')}.xlsx"
            output_path = os.path.join(self.output_dir, filename)

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

        self.logger.info(f"生成{'完成' if not self._is_cancelled else '已取消'}，成功 {success_count}/{total}")
        self.finished_signal.emit(success_count, total)