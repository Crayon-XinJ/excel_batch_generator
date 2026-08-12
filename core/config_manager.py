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
from pathlib import Path

from models.config_models import ProductConfig, NonWorkdaysConfig, Rule, ProductAutoConfig


class ConfigManager:
    """配置管理器（单例模式）"""
    
    _instance = None
    _initialized = False

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self._initialized = True

        # 定位程序根目录
        if getattr(sys, 'frozen', False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        # 加载 config.ini
        self.ini_path = os.path.join(self.base_dir, 'config.ini')
        self._load_or_create_ini()

        # 读取路径配置
        self.config_dir = self._get_path('PATHS', 'config_dir', 'config')
        self.log_dir = self._get_path('PATHS', 'log_dir', 'logs')
        self.templates_dir = self._get_path('PATHS', 'templates_dir', 'config/templates')

        os.makedirs(self.config_dir, exist_ok=True)
        os.makedirs(self.log_dir, exist_ok=True)
        os.makedirs(self.templates_dir, exist_ok=True)

        # 设置子路径
        self.non_workdays_path = os.path.join(self.config_dir, 'non_workdays.json')
        self.products_path = os.path.join(self.config_dir, 'products')
        os.makedirs(self.products_path, exist_ok=True)

    def _load_or_create_ini(self):
        """加载或创建 config.ini"""
        self.config = configparser.ConfigParser()
        if not os.path.exists(self.ini_path):
            self._create_default_ini()
        self.config.read(self.ini_path, encoding='utf-8')

    def _create_default_ini(self):
        """创建默认配置文件"""
        self.config['PATHS'] = {
            'config_dir': 'config',
            'log_dir': 'logs',
            'templates_dir': 'config/templates'
        }
        self.config['LOG'] = {
            'level': 'INFO',
            'max_files': '30',
            'console_output': 'true',
            'file_output': 'true'
        }
        self.config['AUTO'] = {
            'enabled': 'false',
            'time': '07:00',
            'products': '',
            'types': '首件,过程,成品',
            'range_mode': 'hybrid',
            'auto_default_days': '30',
            'run_mode': 'scheduled',
            'exit_after_run': 'true',
            'check_template_exists': 'true'
        }
        self.config['EMAIL'] = {
            'enabled': 'false',
            'smtp_server': 'smtp.qq.com',
            'smtp_port': '465',
            'smtp_ssl': 'true',
            'smtp_tls': 'false',
            'username': '',
            'password': '',
            'to': '',
            'cc': '',
            'subject': 'Excel生成报告 - {date}',
            'report_level': 'detailed',
            'send_only_on_error': 'false',
            'timeout': '30'
        }

        with open(self.ini_path, 'w', encoding='utf-8') as f:
            self.config.write(f)

        # 添加注释
        with open(self.ini_path, 'r', encoding='utf-8') as f:
            content = f.read()
        with open(self.ini_path, 'w', encoding='utf-8') as f:
            f.write("; ============================================================\n")
            f.write("; Excel批量生成工具 - 配置文件\n")
            f.write("; ============================================================\n")
            f.write("; 路径说明：相对路径相对于程序所在目录\n")
            f.write("; ============================================================\n\n")
            f.write(content)

    def _get_path(self, section: str, key: str, default: str) -> str:
        """获取路径配置（支持相对路径转绝对路径）"""
        val = self.config.get(section, key, fallback=default)
        if os.path.isabs(val):
            return val
        else:
            return os.path.join(self.base_dir, val)

    def get_log_config(self) -> dict:
        """获取日志配置"""
        return {
            'level': self.config.get('LOG', 'level', fallback='INFO'),
            'max_files': self.config.getint('LOG', 'max_files', fallback=30),
            'console_output': self.config.getboolean('LOG', 'console_output', fallback=True),
            'file_output': self.config.getboolean('LOG', 'file_output', fallback=True),
            'log_dir': self._get_path('PATHS', 'log_dir', 'logs')
        }

    def get_auto_config(self) -> dict:
        """获取自动化配置"""
        return {
            'enabled': self.config.getboolean('AUTO', 'enabled', fallback=False),
            'time': self.config.get('AUTO', 'time', fallback='07:00'),
            'products': self.config.get('AUTO', 'products', fallback=''),
            'types': self.config.get('AUTO', 'types', fallback='首件,过程,成品'),
            'range_mode': self.config.get('AUTO', 'range_mode', fallback='hybrid'),
            'auto_default_days': self.config.getint('AUTO', 'auto_default_days', fallback=30),
            'run_mode': self.config.get('AUTO', 'run_mode', fallback='scheduled'),
            'exit_after_run': self.config.getboolean('AUTO', 'exit_after_run', fallback=True),
            'check_template_exists': self.config.getboolean('AUTO', 'check_template_exists', fallback=True)
        }

    def get_email_config(self) -> dict:
        """获取邮件配置"""
        return {
            'enabled': self.config.getboolean('EMAIL', 'enabled', fallback=False),
            'smtp_server': self.config.get('EMAIL', 'smtp_server', fallback='smtp.qq.com'),
            'smtp_port': self.config.getint('EMAIL', 'smtp_port', fallback=465),
            'smtp_ssl': self.config.getboolean('EMAIL', 'smtp_ssl', fallback=True),
            'smtp_tls': self.config.getboolean('EMAIL', 'smtp_tls', fallback=False),
            'username': self.config.get('EMAIL', 'username', fallback=''),
            'password': self.config.get('EMAIL', 'password', fallback=''),
            'to': self.config.get('EMAIL', 'to', fallback=''),
            'cc': self.config.get('EMAIL', 'cc', fallback=''),
            'subject': self.config.get('EMAIL', 'subject', fallback='Excel生成报告 - {date}'),
            'report_level': self.config.get('EMAIL', 'report_level', fallback='detailed'),
            'send_only_on_error': self.config.getboolean('EMAIL', 'send_only_on_error', fallback=False),
            'timeout': self.config.getint('EMAIL', 'timeout', fallback=30)
        }

    def save_auto_config(self, config: dict) -> None:
        """保存自动化配置"""
        self.config.set('AUTO', 'enabled', str(config.get('enabled', False)))
        self.config.set('AUTO', 'time', config.get('time', '07:00'))
        self.config.set('AUTO', 'products', config.get('products', ''))
        self.config.set('AUTO', 'types', config.get('types', '首件,过程,成品'))
        self.config.set('AUTO', 'range_mode', config.get('range_mode', 'hybrid'))
        self.config.set('AUTO', 'auto_default_days', str(config.get('auto_default_days', 30)))
        self.config.set('AUTO', 'run_mode', config.get('run_mode', 'scheduled'))
        self.config.set('AUTO', 'exit_after_run', str(config.get('exit_after_run', True)))
        self.config.set('AUTO', 'check_template_exists', str(config.get('check_template_exists', True)))
        
        with open(self.ini_path, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def save_email_config(self, config: dict) -> None:
        """保存邮件配置"""
        for key, value in config.items():
            self.config.set('EMAIL', key, str(value))
        
        with open(self.ini_path, 'w', encoding='utf-8') as f:
            self.config.write(f)

    # ============================================================
    # 非工作日配置
    # ============================================================

    def load_non_workdays(self) -> List[str]:
        if not os.path.exists(self.non_workdays_path):
            return []
        try:
            with open(self.non_workdays_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('non_workdays', [])
        except:
            return []

    def save_non_workdays(self, dates: List[str]) -> None:
        config = NonWorkdaysConfig(non_workdays=dates)
        with open(self.non_workdays_path, 'w', encoding='utf-8') as f:
            json.dump(config.to_dict(), f, ensure_ascii=False, indent=2)

    def export_non_workdays(self, file_path: str) -> None:
        dates = self.load_non_workdays()
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump({'non_workdays': dates}, f, ensure_ascii=False, indent=2)

    def import_non_workdays(self, file_path: str) -> List[str]:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data.get('non_workdays', [])

    # ============================================================
    # 产品配置
    # ============================================================

    def _get_product_config_path(self, product_name: str, template_type: str) -> str:
        product_dir = os.path.join(self.products_path, product_name)
        os.makedirs(product_dir, exist_ok=True)
        return os.path.join(product_dir, f'{template_type}.json')

    def load_product_config(self, product_name: str, template_type: str) -> Optional[ProductConfig]:
        config_path = self._get_product_config_path(product_name, template_type)
        if not os.path.exists(config_path):
            return None
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return ProductConfig.from_dict(data)
        except:
            return None

    def save_product_config(self, config: ProductConfig) -> None:
        config_path = self._get_product_config_path(config.product_name, config.template_type)
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config.to_dict(), f, ensure_ascii=False, indent=2)

    def delete_product_config(self, product_name: str, template_type: str) -> bool:
        config_path = self._get_product_config_path(product_name, template_type)
        if os.path.exists(config_path):
            os.remove(config_path)
            return True
        return False

    def export_product_config(self, product_name: str, template_type: str, file_path: str) -> None:
        config = self.load_product_config(product_name, template_type)
        if config:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config.to_dict(), f, ensure_ascii=False, indent=2)

    def import_product_config(self, file_path: str) -> ProductConfig:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return ProductConfig.from_dict(data)

    def get_products_list(self) -> List[str]:
        if not os.path.exists(self.products_path):
            return []
        return [d for d in os.listdir(self.products_path) 
                if os.path.isdir(os.path.join(self.products_path, d))]

    def get_product_template_types(self, product_name: str) -> List[str]:
        product_dir = os.path.join(self.products_path, product_name)
        if not os.path.exists(product_dir):
            return []
        types = []
        for f in os.listdir(product_dir):
            if f.endswith('.json'):
                types.append(f.replace('.json', ''))
        return types

    # ============================================================
    # 规则管理
    # ============================================================

    def add_rule(self, product_name: str, template_type: str, rule: Rule) -> None:
        config = self.load_product_config(product_name, template_type)
        if config is None:
            config = ProductConfig(product_name=product_name, template_type=template_type)
        config.rules.append(rule)
        self.save_product_config(config)

    def update_rule(self, product_name: str, template_type: str, rule_id: str, new_rule: Rule) -> bool:
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
        config = self.load_product_config(product_name, template_type)
        if config is None:
            return []
        return config.rules

    # ============================================================
    # 全局日期范围
    # ============================================================

    def get_global_date_ranges(self) -> List[List[str]]:
        """从config.ini读取全局日期范围"""
        # 暂不支持在ini中配置，由产品配置各自管理
        return []

    def get_auto_default_days(self) -> int:
        return self.config.getint('AUTO', 'auto_default_days', fallback=30)