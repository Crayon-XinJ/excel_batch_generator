#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
日期计算模块 - 根据模板类型计算目标日期

核心逻辑：
    1. 首件：每个连续工作日段的第一天
    2. 过程：所有工作日（排除非工作日）
    3. 成品：每个连续工作日段的最后一天

关键概念：
    - 工作日段：被非工作日分割成的连续工作日区间
    - 非工作日：用户通过日历配置的日期

示例：
    非工作日: 6.7, 6.14, 6.19, 6.20, 6.28
    工期: 6.11 ~ 7.2
    
    工作日段:
        [6.11, 6.13] → 首件:6.11, 成品:6.13
        [6.15, 6.18] → 首件:6.15, 成品:6.18
        [6.21, 6.27] → 首件:6.21, 成品:6.27
        [6.29, 7.2]  → 首件:6.29, 成品:7.2
"""
"""
日期计算模块 - 根据模板类型计算目标日期

核心逻辑：
    1. 首件：每个连续工作日段的第一天
    2. 过程：所有工作日（排除非工作日）
    3. 成品：每个连续工作日段的最后一天
"""

from datetime import datetime, timedelta
from typing import List, Tuple, Set


class DateCalculator:
    """日期计算器"""
    
    def __init__(self, non_workdays: List[str] = None):
        self.non_workdays = set()
        if non_workdays:
            for d in non_workdays:
                try:
                    self.non_workdays.add(datetime.strptime(d, '%Y-%m-%d').date())
                except:
                    pass
    
    def set_non_workdays(self, dates: List[str]) -> None:
        self.non_workdays.clear()
        for d in dates:
            try:
                self.non_workdays.add(datetime.strptime(d, '%Y-%m-%d').date())
            except:
                pass
    
    def is_workday(self, date: datetime) -> bool:
        """判断是否为工作日（非非工作日）"""
        return date.date() not in self.non_workdays
    
    def _split_into_work_segments(self, date_ranges: List[Tuple[datetime, datetime]]) -> List[Tuple[datetime, datetime]]:
        """将日期范围按非工作日分割成连续的工作日段"""
        all_dates = set()
        for start, end in date_ranges:
            current = start
            while current <= end:
                all_dates.add(current)
                current += timedelta(days=1)
        
        sorted_dates = sorted(all_dates)
        
        segments = []
        segment_start = None
        segment_end = None
        
        for i, date in enumerate(sorted_dates):
            if not self.is_workday(date):
                if segment_start is not None:
                    segments.append((segment_start, segment_end))
                    segment_start = None
                    segment_end = None
                continue
            
            if segment_start is None:
                segment_start = date
                segment_end = date
            else:
                prev_date = sorted_dates[i-1] if i > 0 else None
                if prev_date is not None and (date - prev_date).days == 1 and self.is_workday(prev_date):
                    segment_end = date
                else:
                    segments.append((segment_start, segment_end))
                    segment_start = date
                    segment_end = date
        
        if segment_start is not None:
            segments.append((segment_start, segment_end))
        
        return segments
    
    def calc_first_article_dates(self, date_ranges: List[Tuple[datetime, datetime]]) -> List[datetime]:
        """首件：每个工作日段的第一天"""
        if not date_ranges:
            return []
        segments = self._split_into_work_segments(date_ranges)
        result = [seg[0] for seg in segments]
        return sorted(result)
    
    def calc_process_dates(self, date_ranges: List[Tuple[datetime, datetime]]) -> List[datetime]:
        """过程：所有工作日"""
        if not date_ranges:
            return []
        result = []
        for start, end in date_ranges:
            current = start
            while current <= end:
                if self.is_workday(current):
                    result.append(current)
                current += timedelta(days=1)
        return sorted(result)
    
    def calc_final_article_dates(self, date_ranges: List[Tuple[datetime, datetime]]) -> List[datetime]:
        """成品：每个工作日段的最后一天"""
        if not date_ranges:
            return []
        segments = self._split_into_work_segments(date_ranges)
        result = [seg[1] for seg in segments]
        return sorted(result)
    
    def preview_dates(self, template_type: str, date_ranges: List[Tuple[datetime, datetime]]) -> List[datetime]:
        """根据模板类型计算目标日期"""
        if template_type == '首件':
            return self.calc_first_article_dates(date_ranges)
        elif template_type == '过程':
            return self.calc_process_dates(date_ranges)
        elif template_type == '成品':
            return self.calc_final_article_dates(date_ranges)
        else:
            return []