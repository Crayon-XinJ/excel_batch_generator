#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
配置管理器 - 程序的统一配置入口

职责：
    1. 管理 config.ini 文件的读取、写入和默认创建
    2. 提供程序根目录（EXE所在目录或源码根目录）的定位
    3. 管理非工作日配置（JSON格式）
    4. 管理产品规则配置（按产品名+模板类型分组存储）
    5. 支持用户自定义 config 和 logs 目录路径

设计模式：
    使用单例模式，确保整个程序只有一个配置管理器实例，
    所有模块通过同一个实例获取配置，保证一致性。

使用示例：
    from core.config_manager import ConfigManager
    cm = ConfigManager()
    non_workdays = cm.load_non_workdays()
    config = cm.load_product_config('产品名', '首件')
"""

import os
import sys
import json
import configparser
from typing import List, Optional

from models.config_models import ProductConfig, NonWorkdaysConfig, Rule


class ConfigManager:
    """
    配置管理器（单例模式）
    
    主要功能：
        1. 自动定位程序根目录（EXE环境或开发环境）
        2. 自动创建默认 config.ini
        3. 提供配置目录和日志目录的路径
        4. 管理非工作日配置的读写
        5. 管理产品规则配置的读写
    """
    
    _instance = None          # 单例实例
    _initialized = False      # 是否已初始化

    def __new__(cls):
        """单例模式：确保只有一个实例"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        """
        初始化配置管理器
        
        执行流程：
            1. 定位程序根目录（EXE所在目录 或 源码根目录）
            2. 加载或创建 config.ini
            3. 读取用户自定义的路径配置
            4. 确保必要的目录存在
        """
        if self._initialized:
            return
        self._initialized = True

        # ============================================================
        # 1. 定位程序根目录
        # ============================================================
        # 判断是否为打包后的EXE环境：
        #   - sys.frozen 为 True 表示打包环境
        #   - sys.executable 是 EXE 文件的完整路径
        #   - 开发环境使用 __file__ 定位到 core 目录的上级
        # ============================================================
        if getattr(sys, 'frozen', False):
            # 打包后：根目录为 EXE 文件所在的目录
            self.base_dir = os.path.dirname(sys.executable)
        else:
            # 开发环境：根目录为 core 目录的上级目录
            # __file__ = /path/to/core/config_manager.py
            # os.path.dirname两次 = /path/to/项目根目录
            self.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        # ============================================================
        # 2. 加载或创建 config.ini
        # ============================================================
        # config.ini 位于程序根目录，用户可通过编辑此文件自定义配置
        # 如果文件不存在，程序会自动创建默认配置
        # ============================================================
        self.ini_path = os.path.join(self.base_dir, 'config.ini')
        self._load_or_create_ini()

        # ============================================================
        # 3. 读取路径配置
        # ============================================================
        # 从 config.ini 的 [PATHS] 节读取：
        #   - config_dir: 配置文件存放目录（默认 config）
        #   - log_dir: 日志文件存放目录（默认 logs）
        # 支持相对路径（相对于程序根目录）和绝对路径
        # ============================================================
        self.config_dir = self._get_path('PATHS', 'config_dir', 'config')
        self.log_dir = self._get_path('PATHS', 'log_dir', 'logs')

        # 确保目录存在（如果不存在则创建）
        os.makedirs(self.config_dir, exist_ok=True)
        os.makedirs(self.log_dir, exist_ok=True)

        # ============================================================
        # 4. 设置配置文件的子路径
        # ============================================================
        # non_workdays.json: 存储非工作日列表
        # products/: 存储各产品的规则配置，按产品名分目录
        # ============================================================
        self.non_workdays_path = os.path.join(self.config_dir, 'non_workdays.json')
        self.products_path = os.path.join(self.config_dir, 'products')
        os.makedirs(self.products_path, exist_ok=True)

    # ============================================================
    # config.ini 管理
    # ============================================================

    def _load_or_create_ini(self):
        """
        加载或创建 config.ini 配置文件
        
        执行逻辑：
            1. 创建 ConfigParser 实例
            2. 如果 config.ini 不存在，调用 _create_default_ini() 创建
            3. 读取 config.ini 内容
        """
        self.config = configparser.ConfigParser()
        if not os.path.exists(self.ini_path):
            self._create_default_ini()
        self.config.read(self.ini_path, encoding='utf-8')

    def _create_default_ini(self):
        """
        创建默认的 config.ini 配置文件
        
        默认配置内容：
            [PATHS]
            config_dir = config      # 配置文件存放目录
            log_dir = logs           # 日志文件存放目录
            
            [LOG]
            level = INFO             # 日志级别
            max_files = 30           # 保留最近日志文件数
            console_output = true    # 是否输出到控制台
            file_output = true       # 是否写入日志文件
        
        注意：configparser 不支持写入注释，所以创建后手动添加注释
        """
        # 设置配置值
        self.config['PATHS'] = {
            'config_dir': 'config',
            'log_dir': 'logs'
        }
        self.config['LOG'] = {
            'level': 'INFO',
            'max_files': '30',
            'console_output': 'true',
            'file_output': 'true'
        }

        # 写入配置文件（先写入基础内容）
        with open(self.ini_path, 'w', encoding='utf-8') as f:
            self.config.write(f)

        # 重新读取并手动添加注释到文件开头
        with open(self.ini_path, 'r', encoding='utf-8') as f:
            content = f.read()

        with open(self.ini_path, 'w', encoding='utf-8') as f:
            f.write("; ============================================================\n")
            f.write("; Excel批量生成工具配置文件\n")
            f.write("; ============================================================\n")
            f.write("; 路径说明：\n")
            f.write(";   - 相对路径：相对于程序所在目录\n")
            f.write(";   - 绝对路径：使用完整路径（如 D:/my_config）\n")
            f.write("; ============================================================\n\n")
            f.write(content)

    def _get_path(self, section: str, key: str, default: str) -> str:
        """
        从 config.ini 获取路径配置
        
        如果配置的是相对路径（不以盘符或/开头），则拼接 base_dir；
        如果是绝对路径，则直接返回。
        
        Args:
            section: INI 文件中的节名，如 'PATHS'
            key: 键名，如 'config_dir'
            default: 默认值（当配置项不存在时使用）
        
        Returns:
            str: 完整的路径（绝对路径）
        """
        val = self.config.get(section, key, fallback=default)
        if os.path.isabs(val):
            return val
        else:
            return os.path.join(self.base_dir, val)

    def get_log_config(self) -> dict:
        """
        获取日志配置（供 LogManager 使用）
        
        Returns:
            dict: 包含以下键值：
                - level: 日志级别 (DEBUG/INFO/WARNING/ERROR)
                - max_files: 保留日志文件数
                - console_output: 是否输出到控制台 (bool)
                - file_output: 是否写入日志文件 (bool)
        """
        return {
            'level': self.config.get('LOG', 'level', fallback='INFO'),
            'max_files': self.config.getint('LOG', 'max_files', fallback=30),
            'console_output': self.config.getboolean('LOG', 'console_output', fallback=True),
            'file_output': self.config.getboolean('LOG', 'file_output', fallback=True)
        }

    # ============================================================
    # 非工作日配置管理
    # ============================================================

    def load_non_workdays(self) -> List[str]:
        """
        加载非工作日列表
        
        从 config/non_workdays.json 读取。
        如果文件不存在，返回空列表。
        
        Returns:
            List[str]: 非工作日日期列表，格式 ['2026-06-07', '2026-06-14', ...]
        """
        if not os.path.exists(self.non_workdays_path):
            return []
        try:
            with open(self.non_workdays_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('non_workdays', [])
        except (json.JSONDecodeError, KeyError) as e:
            # 配置文件损坏时返回空列表，不影响程序运行
            print(f"警告：读取非工作日配置失败: {e}")
            return []

    def save_non_workdays(self, dates: List[str]) -> None:
        """
        保存非工作日列表
        
        Args:
            dates: 非工作日日期列表，格式 ['2026-06-07', '2026-06-14', ...]
        """
        config = NonWorkdaysConfig(non_workdays=dates)
        with open(self.non_workdays_path, 'w', encoding='utf-8') as f:
            json.dump(config.to_dict(), f, ensure_ascii=False, indent=2)

    def export_non_workdays(self, file_path: str) -> None:
        """
        导出非工作日配置到指定文件（用于用户导出备份）
        
        Args:
            file_path: 目标文件路径
        """
        dates = self.load_non_workdays()
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump({'non_workdays': dates}, f, ensure_ascii=False, indent=2)

    def import_non_workdays(self, file_path: str) -> List[str]:
        """
        从指定文件导入非工作日配置（用于用户导入备份）
        
        Args:
            file_path: 源文件路径
        
        Returns:
            List[str]: 导入的日期列表
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data.get('non_workdays', [])

    # ============================================================
    # 产品规则配置管理
    # ============================================================

    def _get_product_config_path(self, product_name: str, template_type: str) -> str:
        """
        获取产品配置文件的完整路径
        
        路径结构：config/products/{产品名}/{模板类型}.json
        例如：config/products/23036_N24四方管/首件.json
        
        Args:
            product_name: 产品名称（如 '23036_N24四方管'）
            template_type: 模板类型（'首件' / '过程' / '成品'）
        
        Returns:
            str: 配置文件的完整路径
        """
        product_dir = os.path.join(self.products_path, product_name)
        os.makedirs(product_dir, exist_ok=True)
        return os.path.join(product_dir, f'{template_type}.json')

    def load_product_config(self, product_name: str, template_type: str) -> Optional[ProductConfig]:
        """
        加载产品配置
        
        Args:
            product_name: 产品名称
            template_type: 模板类型
        
        Returns:
            ProductConfig: 配置对象，如果文件不存在则返回 None
        """
        config_path = self._get_product_config_path(product_name, template_type)
        if not os.path.exists(config_path):
            return None
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return ProductConfig.from_dict(data)
        except (json.JSONDecodeError, KeyError) as e:
            print(f"警告：加载产品配置失败 ({product_name}/{template_type}): {e}")
            return None

    def save_product_config(self, config: ProductConfig) -> None:
        """
        保存产品配置
        
        Args:
            config: ProductConfig 对象
        """
        config_path = self._get_product_config_path(config.product_name, config.template_type)
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config.to_dict(), f, ensure_ascii=False, indent=2)

    def delete_product_config(self, product_name: str, template_type: str) -> bool:
        """
        删除产品配置
        
        Args:
            product_name: 产品名称
            template_type: 模板类型
        
        Returns:
            bool: 删除成功返回 True，否则返回 False
        """
        config_path = self._get_product_config_path(product_name, template_type)
        if os.path.exists(config_path):
            os.remove(config_path)
            return True
        return False

    def export_product_config(self, product_name: str, template_type: str, file_path: str) -> None:
        """
        导出产品配置到指定文件（用于用户导出备份）
        
        Args:
            product_name: 产品名称
            template_type: 模板类型
            file_path: 目标文件路径
        """
        config = self.load_product_config(product_name, template_type)
        if config:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config.to_dict(), f, ensure_ascii=False, indent=2)

    def import_product_config(self, file_path: str) -> ProductConfig:
        """
        从指定文件导入产品配置（用于用户导入备份）
        
        Args:
            file_path: 源文件路径
        
        Returns:
            ProductConfig: 导入的配置对象
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return ProductConfig.from_dict(data)

    # ============================================================
    # 产品列表查询
    # ============================================================

    def get_products_list(self) -> List[str]:
        """
        获取所有已保存的产品名称
        
        Returns:
            List[str]: 产品名称列表
        """
        if not os.path.exists(self.products_path):
            return []
        return [d for d in os.listdir(self.products_path) 
                if os.path.isdir(os.path.join(self.products_path, d))]

    def get_product_template_types(self, product_name: str) -> List[str]:
        """
        获取指定产品的所有模板类型
        
        Args:
            product_name: 产品名称
        
        Returns:
            List[str]: 模板类型列表 ['首件', '过程', '成品', ...]
        """
        product_dir = os.path.join(self.products_path, product_name)
        if not os.path.exists(product_dir):
            return []
        types = []
        for f in os.listdir(product_dir):
            if f.endswith('.json'):
                types.append(f.replace('.json', ''))
        return types

    # ============================================================
    # 规则管理（便捷方法，供主窗口直接调用）
    # ============================================================

    def add_rule(self, product_name: str, template_type: str, rule: Rule) -> None:
        """添加一条规则到指定产品的配置中"""
        config = self.load_product_config(product_name, template_type)
        if config is None:
            config = ProductConfig(product_name=product_name, template_type=template_type)
        config.rules.append(rule)
        self.save_product_config(config)

    def update_rule(self, product_name: str, template_type: str, rule_id: str, new_rule: Rule) -> bool:
        """更新指定产品的配置中的一条规则"""
        config = self.load_product_config(product_name, template_type)
        if config is None:
            return False
        for i, r in enumerate(config.rules):
            if r.id == rule_id:
                config.rules[i] = new_rule
                self.save_product_config(config)
                return True
        return False

    def delete_rule(self, product_name: str, template_type: str, rule_id: str) -> bool:
        """删除指定产品的配置中的一条规则"""
        config = self.load_product_config(product_name, template_type)
        if config is None:
            return False
        for i, r in enumerate(config.rules):
            if r.id == rule_id:
                del config.rules[i]
                self.save_product_config(config)
                return True
        return False

    def get_rules(self, product_name: str, template_type: str) -> List[Rule]:
        """获取指定产品的所有规则"""
        config = self.load_product_config(product_name, template_type)
        if config is None:
            return []
        return config.rules