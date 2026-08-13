#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
配置管理器 - 程序的统一配置入口
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

        if getattr(sys, 'frozen', False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        self.ini_path = os.path.join(self.base_dir, 'config.ini')
        self._load_or_create_ini()

        self.config_dir = self._get_path('PATHS', 'config_dir', 'config')
        self.log_dir = self._get_path('PATHS', 'log_dir', 'logs')
        self.templates_dir = self._get_path('PATHS', 'templates_dir', 'config/templates')

        os.makedirs(self.config_dir, exist_ok=True)
        os.makedirs(self.log_dir, exist_ok=True)
        os.makedirs(self.templates_dir, exist_ok=True)

        self.non_workdays_path = os.path.join(self.config_dir, 'non_workdays.json')
        self.products_path = os.path.join(self.config_dir, 'products')
        os.makedirs(self.products_path, exist_ok=True)

    def _load_or_create_ini(self):
        self.config = configparser.ConfigParser()
        if not os.path.exists(self.ini_path):
            self._create_default_ini()
        self.config.read(self.ini_path, encoding='utf-8')

    def _create_default_ini(self):
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
            'check_template_exists': 'true',
            'global_output_dir': ''  # ✅ 全局输出目录（可选）
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
        val = self.config.get(section, key, fallback=default)
        if os.path.isabs(val):
            return val
        else:
            return os.path.join(self.base_dir, val)

    def get_log_config(self) -> dict:
        return {
            'level': self.config.get('LOG', 'level', fallback='INFO'),
            'max_files': self.config.getint('LOG', 'max_files', fallback=30),
            'console_output': self.config.getboolean('LOG', 'console_output', fallback=True),
            'file_output': self.config.getboolean('LOG', 'file_output', fallback=True),
            'log_dir': self._get_path('PATHS', 'log_dir', 'logs')
        }

    def get_auto_config(self) -> dict:
        return {
            'enabled': self.config.getboolean('AUTO', 'enabled', fallback=False),
            'time': self.config.get('AUTO', 'time', fallback='07:00'),
            'products': self.config.get('AUTO', 'products', fallback=''),
            'types': self.config.get('AUTO', 'types', fallback='首件,过程,成品'),
            'range_mode': self.config.get('AUTO', 'range_mode', fallback='hybrid'),
            'auto_default_days': self.config.getint('AUTO', 'auto_default_days', fallback=30),
            'run_mode': self.config.get('AUTO', 'run_mode', fallback='scheduled'),
            'exit_after_run': self.config.getboolean('AUTO', 'exit_after_run', fallback=True),
            'check_template_exists': self.config.getboolean('AUTO', 'check_template_exists', fallback=True),
            'global_output_dir': self.config.get('AUTO', 'global_output_dir', fallback='')
        }

    def save_auto_config(self, config: dict) -> None:
        self.config.set('AUTO', 'enabled', str(config.get('enabled', False)))
        self.config.set('AUTO', 'time', config.get('time', '07:00'))
        self.config.set('AUTO', 'products', config.get('products', ''))
        self.config.set('AUTO', 'types', config.get('types', '首件,过程,成品'))
        self.config.set('AUTO', 'range_mode', config.get('range_mode', 'hybrid'))
        self.config.set('AUTO', 'auto_default_days', str(config.get('auto_default_days', 30)))
        self.config.set('AUTO', 'run_mode', config.get('run_mode', 'scheduled'))
        self.config.set('AUTO', 'exit_after_run', str(config.get('exit_after_run', True)))
        self.config.set('AUTO', 'check_template_exists', str(config.get('check_template_exists', True)))
        self.config.set('AUTO', 'global_output_dir', config.get('global_output_dir', ''))
        
        with open(self.ini_path, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def get_email_config(self) -> dict:
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

    def save_email_config(self, config: dict) -> None:
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
    # 获取产品输出目录
    # ============================================================

    def get_product_output_dir(self, product_name: str, template_type: str) -> str:
        """获取产品的输出目录"""
        config = self.load_product_config(product_name, template_type)
        if config and config.output_dir:
            return config.output_dir
        return ''

    def save_product_output_dir(self, product_name: str, template_type: str, output_dir: str) -> None:
        """保存产品的输出目录"""
        config = self.load_product_config(product_name, template_type)
        if config is None:
            config = ProductConfig(product_name=product_name, template_type=template_type)
        config.output_dir = output_dir
        self.save_product_config(config)