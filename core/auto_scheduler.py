#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
自动化调度器 - 负责判断每天需要生成什么
"""

from datetime import datetime, timedelta
from typing import List, Dict, Tuple, Optional
from core.log_manager import get_logger
from core.date_calculator import DateCalculator
from core.config_manager import ConfigManager


class AutoScheduler:
    """自动化调度器"""
    
    def __init__(self):
        self.logger = get_logger('AutoScheduler')
        self.config_manager = ConfigManager()
        self.date_calculator = DateCalculator()
        
        # 加载非工作日
        non_workdays = self.config_manager.load_non_workdays()
        if non_workdays:
            self.date_calculator.set_non_workdays(non_workdays)
    
    def get_today_tasks(self, products: List[str] = None, date: datetime = None) -> Dict:
        """
        计算今天每个产品需要生成的任务
        
        Returns:
            {
                "product_name": {
                    "首件": True/False,
                    "过程": True/False,
                    "成品": True/False,
                    "dates": {
                        "首件": "2026-08-11",
                        ...
                    },
                    "date_ranges": [(start, end), ...],
                    "output_dir": "/path/to/output"
                }
            }
        """
        if date is None:
            date = datetime.now()
        
        # 获取自动化配置
        auto_config = self.config_manager.get_auto_config()
        
        # 如果没有指定产品列表，从配置读取
        if products is None:
            products_str = auto_config.get('products', '')
            if products_str:
                products = [p.strip() for p in products_str.split(',') if p.strip()]
            else:
                # 获取所有已配置的产品
                products = self.config_manager.get_products_list()
        
        if not products:
            self.logger.warning("没有配置任何产品")
            return {}
        
        # 获取全局过滤类型
        global_types = [t.strip() for t in auto_config.get('types', '首件,过程,成品').split(',') if t.strip()]
        
        tasks = {}
        
        for product_name in products:
            self.logger.debug(f"计算产品任务: {product_name}")
            
            # 获取产品配置（读取首件配置获取基本信息）
            product_config = self.config_manager.load_product_config(product_name, '首件')
            if product_config is None:
                self.logger.warning(f"产品 {product_name} 没有配置，跳过")
                continue
            
            # 获取产品的自动化配置
            auto = product_config.auto_config
            if auto and not auto.enabled:
                self.logger.debug(f"产品 {product_name} 已禁用自动化")
                continue
            
            # 确定日期范围
            date_ranges = self._get_date_ranges(product_name, auto, date)
            if not date_ranges:
                # 如果没配置日期范围，使用自动推断
                days = auto_config.get('auto_default_days', 30)
                start = date - timedelta(days=days)
                end = date + timedelta(days=days)
                date_ranges = [(start, end)]
                self.logger.debug(f"使用自动推断日期范围: {start.strftime('%Y-%m-%d')} ~ {end.strftime('%Y-%m-%d')}")
            
            # 计算该产品的所有工作日段
            segments = self.date_calculator._split_into_work_segments(date_ranges)
            
            # 判断今天属于哪个段
            product_tasks = {
                "首件": False,
                "过程": False,
                "成品": False,
                "dates": {},
                "date_ranges": date_ranges,
                "output_dir": auto.output_dir if auto and auto.output_dir else ''
            }
            
            for seg_start, seg_end in segments:
                if seg_start <= date <= seg_end:
                    if date == seg_start:
                        product_tasks["首件"] = True
                        product_tasks["dates"]["首件"] = date.strftime('%Y-%m-%d')
                    if date == seg_end:
                        product_tasks["成品"] = True
                        product_tasks["dates"]["成品"] = date.strftime('%Y-%m-%d')
                    if self.date_calculator.is_workday(date):
                        # 检查类型过滤
                        product_types = auto.types if auto and auto.types else global_types
                        if "过程" in product_types:
                            product_tasks["过程"] = True
                            product_tasks["dates"]["过程"] = date.strftime('%Y-%m-%d')
                    break
            
            # 应用全局类型过滤
            if global_types:
                for key in list(product_tasks.keys()):
                    if key in ["首件", "过程", "成品"] and key not in global_types:
                        product_tasks[key] = False
            
            # 如果没有任何任务，跳过该产品
            if not product_tasks["首件"] and not product_tasks["过程"] and not product_tasks["成品"]:
                self.logger.debug(f"产品 {product_name} 今日无任务")
                continue
            
            tasks[product_name] = product_tasks
        
        return tasks
    
    def _get_date_ranges(self, product_name: str, auto, today: datetime) -> List[Tuple[datetime, datetime]]:
        """获取产品的日期范围"""
        # 优先使用产品配置中的 date_ranges
        product_config = self.config_manager.load_product_config(product_name, '首件')
        if product_config and product_config.date_ranges:
            ranges = []
            for r in product_config.date_ranges:
                try:
                    start = datetime.strptime(r[0], '%Y-%m-%d')
                    end = datetime.strptime(r[1], '%Y-%m-%d')
                    ranges.append((start, end))
                except:
                    pass
            if ranges:
                return ranges
        
        # 如果有 auto_config 且配置了 date_ranges
        if auto and auto.date_ranges:
            ranges = []
            for r in auto.date_ranges:
                try:
                    start = datetime.strptime(r[0], '%Y-%m-%d')
                    end = datetime.strptime(r[1], '%Y-%m-%d')
                    ranges.append((start, end))
                except:
                    pass
            if ranges:
                return ranges
        
        # hybrid 模式：尝试全局日期范围
        if auto and auto.range_mode == 'hybrid':
            # 检查是否有全局配置
            global_ranges = self.config_manager.get_global_date_ranges()
            if global_ranges:
                ranges = []
                for r in global_ranges:
                    try:
                        start = datetime.strptime(r[0], '%Y-%m-%d')
                        end = datetime.strptime(r[1], '%Y-%m-%d')
                        ranges.append((start, end))
                    except:
                        pass
                if ranges:
                    return ranges
        
        # auto 模式：返回空，由调用方处理自动推断
        if auto and auto.range_mode == 'auto':
            return []
        
        # 没有配置，返回空
        return []