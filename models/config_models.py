#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
配置数据模型 - 使用 dataclass 定义配置结构
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
class ProductAutoConfig:
    """
    产品自动化配置（可选）
    """
    enabled: bool = True
    date_ranges: List[List[str]] = field(default_factory=list)
    types: List[str] = field(default_factory=list)
    output_dir: str = ''  # 产品级别的输出目录（覆盖全局）
    range_mode: str = 'hybrid'

    def to_dict(self) -> dict:
        return {
            'enabled': self.enabled,
            'date_ranges': self.date_ranges,
            'types': self.types,
            'output_dir': self.output_dir,
            'range_mode': self.range_mode
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'ProductAutoConfig':
        return cls(
            enabled=data.get('enabled', True),
            date_ranges=data.get('date_ranges', []),
            types=data.get('types', []),
            output_dir=data.get('output_dir', ''),
            range_mode=data.get('range_mode', 'hybrid')
        )


@dataclass
class ProductConfig:
    """
    产品配置
    
    属性说明：
        product_name: 产品名称
        template_type: 模板类型（首件/过程/成品）
        template_path: 模板文件路径
        output_dir: 输出目录路径（每个产品/类型独立）
        rules: 规则列表
        date_ranges: 生产工期范围
        auto_config: 自动化配置（可选）
    """
    product_name: str
    template_type: Literal['首件', '过程', '成品']
    template_path: str = ''
    output_dir: str = ''  # ✅ 新增：输出目录
    rules: List[Rule] = field(default_factory=list)
    date_ranges: List[List[str]] = field(default_factory=list)
    auto_config: Optional[ProductAutoConfig] = None

    def to_dict(self) -> dict:
        result = {
            'product_name': self.product_name,
            'template_type': self.template_type,
            'rules': [r.to_dict() for r in self.rules],
            'date_ranges': self.date_ranges
        }
        if self.template_path:
            result['template_path'] = self.template_path
        if self.output_dir:
            result['output_dir'] = self.output_dir  # ✅ 新增
        if self.auto_config:
            result['auto_config'] = self.auto_config.to_dict()
        return result

    @classmethod
    def from_dict(cls, data: dict) -> 'ProductConfig':
        auto_config = None
        if 'auto_config' in data:
            auto_config = ProductAutoConfig.from_dict(data['auto_config'])
        
        return cls(
            product_name=data['product_name'],
            template_type=data['template_type'],
            template_path=data.get('template_path', ''),
            output_dir=data.get('output_dir', ''),  # ✅ 新增
            rules=[Rule.from_dict(r) for r in data.get('rules', [])],
            date_ranges=data.get('date_ranges', []),
            auto_config=auto_config
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