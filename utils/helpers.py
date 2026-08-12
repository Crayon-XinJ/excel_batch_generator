#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
工具函数模块 - 日期解析、格式化、文本处理等

功能：
    1. 多种日期格式的解析（YYYY-MM-DD, YYYY/MM/DD, YYYYMMDD, 中文格式等）
    2. 根据原始格式重新格式化日期（保持分隔符和补零方式）
    3. 从文本中提取所有日期
    4. 获取中文星期名称
    5. 日期范围判断
"""

import re
from datetime import datetime


def parse_date_string(date_str: str) -> datetime:
    """
    解析多种格式的日期字符串
    
    支持的格式：
        - 2026-06-11
        - 2026/6/11
        - 2026.6.11
        - 2026年6月11日
        - 20260611 (8位纯数字)
    """
    date_str = date_str.strip()
    patterns = [
        r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})',
        r'(\d{4})年(\d{1,2})月(\d{1,2})日',
        r'(\d{4})(\d{2})(\d{2})',
    ]
    for pattern in patterns:
        match = re.search(pattern, date_str)
        if match:
            year, month, day = map(int, match.groups())
            return datetime(year, month, day)
    raise ValueError(f"无法解析日期: {date_str}")


def format_date_for_template(date: datetime, original_date_str: str) -> str:
    """
    根据原始日期字符串的格式，生成新的日期字符串
    强制月日补零为两位（8位格式）
    """
    if '-' in original_date_str:
        return f"{date.year}-{date.month:02d}-{date.day:02d}"
    elif '/' in original_date_str:
        return f"{date.year}/{date.month:02d}/{date.day:02d}"
    elif '.' in original_date_str:
        return f"{date.year}.{date.month:02d}.{date.day:02d}"
    elif '年' in original_date_str and '月' in original_date_str and '日' in original_date_str:
        return f"{date.year}年{date.month:02d}月{date.day:02d}日"
    else:
        return f"{date.year}{date.month:02d}{date.day:02d}"


def extract_first_date_from_text(text: str) -> tuple:
    """
    从文本中提取第一个日期
    
    Returns:
        tuple: (date_obj, date_str, prefix, suffix)
        如果没有找到日期，返回 (None, None, None, None)
    """
    patterns = [
        (r'(\d{4})-(\d{1,2})-(\d{1,2})', '-'),
        (r'(\d{4})/(\d{1,2})/(\d{1,2})', '/'),
        (r'(\d{4})\.(\d{1,2})\.(\d{1,2})', '.'),
        (r'(\d{4})年(\d{1,2})月(\d{1,2})日', '年'),
        (r'(\d{4})(\d{2})(\d{2})', ''),
    ]
    for pattern, _ in patterns:
        match = re.search(pattern, text)
        if match:
            year, month, day = map(int, match.groups())
            date_obj = datetime(year, month, day)
            date_str = match.group(0)
            prefix = text[:match.start()]
            suffix = text[match.end():]
            return date_obj, date_str, prefix, suffix
    return None, None, None, None


def extract_all_dates_from_text(text: str) -> list:
    """
    从文本中提取所有日期
    
    Returns:
        list: [(date_obj, date_str, start_pos, end_pos), ...]
        按日期在文本中出现的顺序排列
    """
    patterns = [
        r'(\d{4})-(\d{1,2})-(\d{1,2})',
        r'(\d{4})/(\d{1,2})/(\d{1,2})',
        r'(\d{4})\.(\d{1,2})\.(\d{1,2})',
        r'(\d{4})年(\d{1,2})月(\d{1,2})日',
        r'(\d{4})(\d{2})(\d{2})',
    ]
    results = []
    for pattern in patterns:
        for match in re.finditer(pattern, text):
            year, month, day = map(int, match.groups())
            date_obj = datetime(year, month, day)
            results.append((date_obj, match.group(0), match.start(), match.end()))
    results.sort(key=lambda x: x[2])
    return results


def get_weekday_name(date: datetime) -> str:
    """获取中文星期名称"""
    weekdays = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    return weekdays[date.weekday()]


def is_within_range(date: datetime, ranges: list) -> bool:
    """判断日期是否在任一日期范围内"""
    for start, end in ranges:
        if start <= date <= end:
            return True
    return False


def get_products_from_config(config_dir: str) -> list:
    """获取所有已保存的产品名称"""
    import os
    products_path = os.path.join(config_dir, 'products')
    if not os.path.exists(products_path):
        return []
    return [d for d in os.listdir(products_path) 
            if os.path.isdir(os.path.join(products_path, d))]