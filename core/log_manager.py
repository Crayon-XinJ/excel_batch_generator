#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
日志管理模块 - 程序的统一日志系统
"""

import os
import sys
import json
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from pathlib import Path


class LogManager:
    """日志管理器（单例模式）"""
    
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
        self.config = self._load_config()
        self._setup_logger()

    def _load_config(self):
        """加载日志配置"""
        default_config = {
            'level': 'INFO',
            'max_files': 30,
            'console_output': True,
            'file_output': True,
            'log_dir': 'logs'
        }
        
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        config_path = Path(base_dir) / 'config.ini'
        if config_path.exists():
            try:
                import configparser
                cp = configparser.ConfigParser()
                cp.read(config_path, encoding='utf-8')
                if 'LOG' in cp:
                    default_config['level'] = cp.get('LOG', 'level', fallback='INFO')
                    default_config['max_files'] = cp.getint('LOG', 'max_files', fallback=30)
                    default_config['console_output'] = cp.getboolean('LOG', 'console_output', fallback=True)
                    default_config['file_output'] = cp.getboolean('LOG', 'file_output', fallback=True)
                    default_config['log_dir'] = cp.get('PATHS', 'log_dir', fallback='logs')
            except Exception as e:
                print(f"警告：加载日志配置失败，使用默认配置: {e}")
        
        return default_config

    def _setup_logger(self):
        """设置日志系统"""
        self.logger = logging.getLogger('ExcelBatchGenerator')
        level_name = self.config['level'].upper()
        self.logger.setLevel(getattr(logging, level_name, logging.INFO))
        self.logger.handlers.clear()

        if self.config['file_output']:
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            log_dir = Path(base_dir) / self.config['log_dir']
            log_dir.mkdir(exist_ok=True)
            
            log_file = log_dir / f"generate_{datetime.now().strftime('%Y-%m-%d')}.log"
            file_handler = RotatingFileHandler(
                log_file, 
                maxBytes=10 * 1024 * 1024,
                backupCount=5,
                encoding='utf-8'
            )
            file_handler.setFormatter(self._get_formatter())
            self.logger.addHandler(file_handler)
            self._clean_old_logs(log_dir)

        if self.config['console_output']:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(self._get_formatter())
            self.logger.addHandler(console_handler)

    def _get_formatter(self):
        return logging.Formatter(
            '[%(asctime)s.%(msecs)03d] [%(levelname)s] [%(name)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

    def _clean_old_logs(self, log_dir):
        try:
            log_files = sorted(log_dir.glob('generate_*.log'))
            max_files = self.config.get('max_files', 30)
            if len(log_files) > max_files:
                for f in log_files[:-max_files]:
                    f.unlink()
        except Exception:
            pass

    def get_logger(self, name=None):
        if name:
            return self.logger.getChild(name)
        return self.logger

    def debug(self, msg, *args, **kwargs):
        self.logger.debug(msg, *args, **kwargs)

    def info(self, msg, *args, **kwargs):
        self.logger.info(msg, *args, **kwargs)

    def warning(self, msg, *args, **kwargs):
        self.logger.warning(msg, *args, **kwargs)

    def error(self, msg, *args, **kwargs):
        self.logger.error(msg, *args, **kwargs)

    def exception(self, msg, *args, **kwargs):
        self.logger.exception(msg, *args, **kwargs)


log_manager = LogManager()

def get_logger(name=None):
    return log_manager.get_logger(name)

def debug(msg, *args, **kwargs):
    log_manager.debug(msg, *args, **kwargs)

def info(msg, *args, **kwargs):
    log_manager.info(msg, *args, **kwargs)

def warning(msg, *args, **kwargs):
    log_manager.warning(msg, *args, **kwargs)

def error(msg, *args, **kwargs):
    log_manager.error(msg, *args, **kwargs)

def exception(msg, *args, **kwargs):
    log_manager.exception(msg, *args, **kwargs)