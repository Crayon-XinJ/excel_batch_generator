#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
配置数据模型 - 使用 dataclass 定义配置结构

为什么使用 dataclass？
    - 自动生成 __init__、__repr__ 等方法
    - 类型提示清晰
    - 易于序列化/反序列化（to_dict / from_dict）
"""

from dataclasses import dataclass, field
from typing import List, Optional, Literal
import json


@dataclass
class Rule:
    """
    内容修改规则
    
    属性说明：
        id: 规则唯一标识（UUID前8位）
        target_type: 目标类型 (range/column/cell/cells)
        target: 目标位置 (如 "C9-L9", "D6-D15", "F9,K9")
        value_type: 值类型 (random/date/text_with_date/night_shift)
        sheet_name: 工作表名称（空表示所有工作表）
        min_val: 随机数最小值
        max_val: 随机数最大值
        decimals: 随机数小数位数
        enabled: 是否启用
    """
    id: str
    target_type: Literal['range', 'column', 'cell', 'cells']
    target: str
    value_type: Literal['random', 'date', 'text_with_date', 'night_shift']
    sheet_name: str = ''
    min_val: Optional[float] = None
    max_val: Optional[float] = None
    decimals: int = 2
    enabled: bool = True

    def to_dict(self) -> dict:
        """序列化为字典（用于JSON存储）"""
        return {
            'id': self.id,
            'target_type': self.target_type,
            'target': self.target,
            'value_type': self.value_type,
            'sheet_name': self.sheet_name,
            'min_val': self.min_val,
            'max_val': self.max_val,
            'decimals': self.decimals,
            'enabled': self.enabled
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'Rule':
        """从字典反序列化"""
        return cls(
            id=data['id'],
            target_type=data['target_type'],
            target=data['target'],
            value_type=data['value_type'],
            sheet_name=data.get('sheet_name', ''),
            min_val=data.get('min_val'),
            max_val=data.get('max_val'),
            decimals=data.get('decimals', 2),
            enabled=data.get('enabled', True)
        )


@dataclass
class ProductConfig:
    """
    产品配置
    
    属性说明：
        product_name: 产品名称
        template_type: 模板类型（首件/过程/成品）
        rules: 规则列表
        date_ranges: 生产工期范围 [[start, end], ...]
    """
    product_name: str
    template_type: Literal['首件', '过程', '成品']
    rules: List[Rule] = field(default_factory=list)
    date_ranges: List[List[str]] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            'product_name': self.product_name,
            'template_type': self.template_type,
            'rules': [r.to_dict() for r in self.rules],
            'date_ranges': self.date_ranges
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'ProductConfig':
        return cls(
            product_name=data['product_name'],
            template_type=data['template_type'],
            rules=[Rule.from_dict(r) for r in data.get('rules', [])],
            date_ranges=data.get('date_ranges', [])
        )


@dataclass
class NonWorkdaysConfig:
    """非工作日配置"""
    non_workdays: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {'non_workdays': self.non_workdays}

    @classmethod
    def from_dict(cls, data: dict) -> 'NonWorkdaysConfig':
        return cls(non_workdays=data.get('non_workdays', []))