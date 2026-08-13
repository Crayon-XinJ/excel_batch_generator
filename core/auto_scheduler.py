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
            self.logger.debug(f"加载非工作日: {len(non_workdays)} 天")
        else:
            self.logger.debug("未配置非工作日")
    
    def get_today_tasks(self, products: List[str] = None, date: datetime = None) -> Dict:
        """
        计算今天每个产品需要生成的任务
        
        Args:
            products: 产品名称列表，如果为 None 则从配置读取所有产品
            date: 指定日期，如果为 None 则使用当前日期
        
        Returns:
            {
                "product_name": {
                    "首件": True/False,
                    "过程": True/False,
                    "成品": True/False,
                    "dates": {
                        "首件": "2026-08-12",
                        ...
                    },
                    "date_ranges": [(start, end), ...],
                    "output_dir": "/path/to/output"
                }
            }
        """
        # ----- 1. 确定日期（时间部分清零） -----
        if date is None:
            date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        else:
            date = date.replace(hour=0, minute=0, second=0, microsecond=0)
        self.logger.debug(f"计算任务日期（归一化后）: {date.strftime('%Y-%m-%d')}")
        
        # ----- 2. 获取自动化配置 -----
        auto_config = self.config_manager.get_auto_config()
        self.logger.debug(f"自动化配置: enabled={auto_config.get('enabled')}, range_mode={auto_config.get('range_mode')}")
        
        # ----- 3. 确定产品列表 -----
        if products is None:
            products_str = auto_config.get('products', '')
            if products_str:
                products = [p.strip() for p in products_str.split(',') if p.strip()]
            else:
                # 如果配置中没有指定产品，则获取所有已配置的产品
                products = self.config_manager.get_products_list()
                self.logger.info(f"从配置管理器获取所有产品: {products}")
        
        if not products:
            self.logger.warning("没有配置任何产品")
            return {}
        
        self.logger.debug(f"参与产品: {products}")
        
        # ----- 4. 获取全局过滤类型 -----
        global_types = [t.strip() for t in auto_config.get('types', '首件,过程,成品').split(',') if t.strip()]
        self.logger.debug(f"全局类型过滤: {global_types}")
        
        # ----- 5. 遍历产品计算任务 -----
        tasks = {}
        
        for product_name in products:
            self.logger.debug(f"计算产品任务: {product_name}")
            
            # 尝试加载该产品的配置（优先用首件获取基本信息）
            product_config = self.config_manager.load_product_config(product_name, '首件')
            if product_config is None:
                self.logger.warning(f"产品 {product_name} 没有首件配置，跳过")
                continue
            
            # 获取产品的自动化配置（可能为 None）
            auto = product_config.auto_config if hasattr(product_config, 'auto_config') else None
            if auto and not auto.enabled:
                self.logger.debug(f"产品 {product_name} 已禁用自动化")
                continue
            
            # ----- 5.1 获取日期范围 -----
            date_ranges = self._get_date_ranges(product_name, auto, date)
            
            if not date_ranges:
                # 如果没有任何配置，使用自动推断
                days = auto_config.get('auto_default_days', 30)
                start = date - timedelta(days=days)
                end = date + timedelta(days=days)
                date_ranges = [(start, end)]
                self.logger.debug(f"使用自动推断日期范围: {start.strftime('%Y-%m-%d')} ~ {end.strftime('%Y-%m-%d')}")
            
            self.logger.debug(f"产品 {product_name} 日期范围: {[(s.strftime('%Y-%m-%d'), e.strftime('%Y-%m-%d')) for s, e in date_ranges]}")
            
            # ----- 5.2 计算工作日段（基于非工作日切分） -----
            segments = self.date_calculator._split_into_work_segments(date_ranges)
            self.logger.debug(f"工作日段: {[(s.strftime('%Y-%m-%d'), e.strftime('%Y-%m-%d')) for s, e in segments]}")
            
            # ----- 5.3 判断今天属于哪个段 -----
            product_tasks = {
                "首件": False,
                "过程": False,
                "成品": False,
                "dates": {},
                "date_ranges": date_ranges,
                "output_dir": auto.output_dir if auto and hasattr(auto, 'output_dir') and auto.output_dir else ''
            }
            
            # 检查今天是否在任何一个段内
            found_segment = False
            for seg_start, seg_end in segments:
                if seg_start <= date <= seg_end:
                    found_segment = True
                    self.logger.debug(f"今日 {date.strftime('%Y-%m-%d')} 在段 {seg_start.strftime('%Y-%m-%d')}~{seg_end.strftime('%Y-%m-%d')} 中")
                    
                    # 检查是否为段的第一天（首件）
                    if date == seg_start:
                        product_tasks["首件"] = True
                        product_tasks["dates"]["首件"] = date.strftime('%Y-%m-%d')
                        self.logger.debug(f"  ✅ 首件")
                    
                    # 检查是否为段的最后一天（成品）
                    if date == seg_end:
                        product_tasks["成品"] = True
                        product_tasks["dates"]["成品"] = date.strftime('%Y-%m-%d')
                        self.logger.debug(f"  ✅ 成品")
                    
                    # 检查是否为工作日（过程）
                    if self.date_calculator.is_workday(date):
                        # 检查类型过滤
                        product_types = auto.types if auto and hasattr(auto, 'types') and auto.types else global_types
                        if "过程" in product_types:
                            product_tasks["过程"] = True
                            product_tasks["dates"]["过程"] = date.strftime('%Y-%m-%d')
                            self.logger.debug(f"  ✅ 过程")
                    break
            
            if not found_segment:
                self.logger.debug(f"今日 {date.strftime('%Y-%m-%d')} 不在任何日期段内")
            
            # ----- 5.4 应用全局类型过滤 -----
            if global_types:
                for key in list(product_tasks.keys()):
                    if key in ["首件", "过程", "成品"] and key not in global_types:
                        product_tasks[key] = False
                        self.logger.debug(f"  类型 {key} 被全局过滤禁用")
            
            # ----- 5.5 如果没有任何任务，跳过该产品 -----
            if not product_tasks["首件"] and not product_tasks["过程"] and not product_tasks["成品"]:
                self.logger.debug(f"产品 {product_name} 今日无任务")
                continue
            
            tasks[product_name] = product_tasks
            self.logger.debug(f"产品 {product_name} 最终任务: 首件={product_tasks['首件']}, 过程={product_tasks['过程']}, 成品={product_tasks['成品']}")
        
        self.logger.info(f"今日任务: {len(tasks)} 个产品")
        return tasks
    
    def _get_date_ranges(self, product_name: str, auto, today: datetime) -> List[Tuple[datetime, datetime]]:
        """
        获取产品的日期范围
        
        优先级：
        1. 产品配置中直接配置的 date_ranges（首件/过程/成品）
        2. auto_config 中的 date_ranges
        3. 全局 date_ranges（暂未实现）
        4. 空列表（由调用方处理自动推断）
        """
        # ----- 1. 从产品配置中读取（遍历三种类型，只要有配置就用） -----
        for template_type in ['首件', '过程', '成品']:
            config = self.config_manager.load_product_config(product_name, template_type)
            if config and config.date_ranges:
                ranges = []
                for r in config.date_ranges:
                    try:
                        if len(r) == 2:
                            start = datetime.strptime(r[0], '%Y-%m-%d')
                            end = datetime.strptime(r[1], '%Y-%m-%d')
                            ranges.append((start, end))
                    except Exception as e:
                        self.logger.warning(f"解析日期范围失败: {r} - {e}")
                if ranges:
                    self.logger.debug(f"从 {template_type} 配置读取日期范围: {ranges}")
                    return ranges
        
        # ----- 2. 从 auto_config 中读取 -----
        if auto and hasattr(auto, 'date_ranges') and auto.date_ranges:
            ranges = []
            for r in auto.date_ranges:
                try:
                    if len(r) == 2:
                        start = datetime.strptime(r[0], '%Y-%m-%d')
                        end = datetime.strptime(r[1], '%Y-%m-%d')
                        ranges.append((start, end))
                except Exception as e:
                    self.logger.warning(f"解析 auto_config 日期范围失败: {r} - {e}")
            if ranges:
                self.logger.debug(f"从 auto_config 读取日期范围: {ranges}")
                return ranges
        
        # ----- 3. 混合模式尝试全局（暂未实现） -----
        if auto and hasattr(auto, 'range_mode') and auto.range_mode == 'hybrid':
            # 检查全局日期范围（当前未实现，返回空）
            pass
        
        # ----- 4. 如果 auto 或 hybrid 模式且无配置，返回空 -----
        if auto and hasattr(auto, 'range_mode') and auto.range_mode in ['auto', 'hybrid']:
            self.logger.debug(f"产品 {product_name} 模式为 {auto.range_mode}，无手动配置，将使用自动推断")
            return []
        
        # ----- 5. 默认返回空 -----
        self.logger.debug(f"产品 {product_name} 无任何日期范围配置")
        return []