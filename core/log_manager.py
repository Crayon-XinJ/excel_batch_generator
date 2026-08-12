#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
日志管理模块 - 程序的统一日志系统

职责：
    1. 提供统一的多级别日志记录（DEBUG / INFO / WARNING / ERROR）
    2. 同时支持文件输出和控制台输出
    3. 日志按天滚动，自动清理过期日志
    4. 从 ConfigManager 获取日志配置（日志级别、保留天数等）
    5. 支持模块级日志器（每个模块可创建独立的子日志器）

设计模式：
    使用单例模式，确保整个程序只有一个日志管理器实例，
    所有模块共享同一个日志配置。

使用示例：
    from core.log_manager import get_logger
    logger = get_logger('ExcelGenerator')
    logger.info('开始生成文件')
    logger.debug('应用规则: C9-L9 -> random(92.31, 92.69)')
    logger.error('生成失败: 模板文件不存在')
    logger.exception('详细异常信息')  # 自动包含堆栈跟踪
"""

import os
import sys
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from pathlib import Path

from core.config_manager import ConfigManager


class LogManager:
    """
    日志管理器（单例模式）
    
    主要功能：
        1. 从 ConfigManager 读取日志配置
        2. 创建并配置 Python logging 实例
        3. 管理日志文件的滚动和清理
        4. 提供便捷的日志记录方法
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
        初始化日志管理器
        
        执行流程：
            1. 获取 ConfigManager 实例（单例）
            2. 读取日志配置
            3. 创建 Python logging 实例
            4. 配置文件输出和控制台输出
            5. 清理过期日志文件
        """
        if self._initialized:
            return
        self._initialized = True

        # ============================================================
        # 1. 获取配置管理器
        # ============================================================
        # ConfigManager 负责读取 config.ini 中的配置
        # 日志配置位于 [LOG] 节：
        #   - level: 日志级别
        #   - max_files: 保留日志文件数
        #   - console_output: 控制台输出开关
        #   - file_output: 文件输出开关
        # ============================================================
        self.config_manager = ConfigManager()
        self._setup_logger()

    def _setup_logger(self):
        """
        配置日志系统
        
        执行步骤：
            1. 获取日志配置
            2. 创建 logger 实例
            3. 设置日志级别
            4. 添加文件输出（如果启用）
            5. 添加控制台输出（如果启用）
            6. 清理过期日志
        """
        # ----- 获取日志配置 -----
        log_config = self.config_manager.get_log_config()

        # ----- 创建 logger -----
        # 使用固定名称 'ExcelBatchGenerator'，方便识别
        self.logger = logging.getLogger('ExcelBatchGenerator')
        self.logger.setLevel(getattr(logging, log_config['level'].upper(), logging.INFO))
        
        # 清除已有的 handlers（防止重复添加，避免日志重复输出）
        self.logger.handlers.clear()

        # ----- 文件输出 -----
        if log_config['file_output']:
            # 从 ConfigManager 获取日志目录路径
            log_dir = Path(self.config_manager.log_dir)
            log_dir.mkdir(exist_ok=True)

            # 日志文件名：按天区分
            # 格式：generate_2026-08-10.log
            log_file = log_dir / f"generate_{datetime.now().strftime('%Y-%m-%d')}.log"
            
            # RotatingFileHandler：
            #   - maxBytes=10MB：单个日志文件超过10MB时自动轮转
            #   - backupCount=5：保留5个备份文件
            #   为什么用大小轮转而不是时间轮转？
            #   因为DEBUG模式下日志量可能很大，按大小更可控
            file_handler = RotatingFileHandler(
                log_file,
                maxBytes=10 * 1024 * 1024,  # 10MB
                backupCount=5,
                encoding='utf-8'
            )
            file_handler.setFormatter(self._get_formatter())
            self.logger.addHandler(file_handler)
            
            # 清理过期日志
            self._clean_old_logs(log_dir)

        # ----- 控制台输出 -----
        if log_config['console_output']:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(self._get_formatter())
            self.logger.addHandler(console_handler)

    def _get_formatter(self):
        """
        返回日志格式器
        
        格式说明：
            [时间.毫秒] [日志级别] [模块名] 消息内容
        
        示例：
            [2026-08-10 14:32:15.123] [INFO] [MainWindow] 程序启动
            [2026-08-10 14:32:18.456] [DEBUG] [ExcelGenerator] 应用规则: rule_001
        
        为什么包含毫秒？
            DEBUG 模式下精确到毫秒有助于性能分析和问题定位。
        """
        return logging.Formatter(
            '[%(asctime)s.%(msecs)03d] [%(levelname)s] [%(name)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

    def _clean_old_logs(self, log_dir: Path):
        """
        清理过期的日志文件
        
        保留最近 N 天的日志文件（由 config.ini 中的 max_files 控制），
        超过 N 天的自动删除。
        
        Args:
            log_dir: 日志目录路径
        """
        try:
            # 获取所有日志文件（按修改时间排序）
            log_files = sorted(log_dir.glob('generate_*.log'))
            max_files = self.config_manager.get_log_config().get('max_files', 30)
            
            # 如果文件数超过保留数量，删除最旧的
            if len(log_files) > max_files:
                for f in log_files[:-max_files]:
                    f.unlink()  # 删除文件
        except Exception:
            # 清理失败不影响程序运行
            pass

    def get_logger(self, name: str = None):
        """
        获取 logger 实例
        
        如果指定 name，则创建子 logger，日志中会显示模块名称。
        
        Args:
            name: 模块名称，如 'MainWindow'、'ExcelGenerator' 等
                  如果不传，返回根 logger
        
        Returns:
            logging.Logger: logger 实例
        
        使用示例：
            logger = log_manager.get_logger('ExcelGenerator')
            logger.info('这条日志会显示 [ExcelGenerator]')
        """
        if name:
            return self.logger.getChild(name)
        return self.logger

    # ============================================================
    # 便捷日志方法（直接调用，无需先 get_logger）
    # ============================================================
    # 这些方法让调用更简洁：
    #   log_manager.info('消息')
    #   而不是 log_manager.logger.info('消息')
    # 
    # 同时支持在调用时直接传参，与 logging 模块方法签名一致
    # ============================================================

    def debug(self, msg, *args, **kwargs):
        """记录 DEBUG 级别日志"""
        self.logger.debug(msg, *args, **kwargs)

    def info(self, msg, *args, **kwargs):
        """记录 INFO 级别日志"""
        self.logger.info(msg, *args, **kwargs)

    def warning(self, msg, *args, **kwargs):
        """记录 WARNING 级别日志"""
        self.logger.warning(msg, *args, **kwargs)

    def error(self, msg, *args, **kwargs):
        """记录 ERROR 级别日志"""
        self.logger.error(msg, *args, **kwargs)

    def exception(self, msg, *args, **kwargs):
        """
        记录异常信息（自动包含堆栈跟踪）
        
        此方法会在日志中自动添加异常的完整堆栈信息，
        专门用于 except 块中记录异常。
        
        使用示例：
            try:
                do_something()
            except Exception as e:
                log_manager.exception('操作失败')
        """
        self.logger.exception(msg, *args, **kwargs)


# ============================================================
# 全局实例和便捷函数
# ============================================================
# 程序启动时自动创建单例实例，所有模块共享
# 提供模块级函数，让调用更简洁：
#   from core.log_manager import get_logger
#   logger = get_logger('MyModule')
#   logger.info('消息')
# ============================================================

# 创建全局单例实例
log_manager = LogManager()


def get_logger(name: str = None):
    """
    获取 logger 实例（便捷函数）
    
    这是最推荐的使用方式：
        from core.log_manager import get_logger
        logger = get_logger('ExcelGenerator')
        logger.info('开始生成')
    """
    return log_manager.get_logger(name)


def debug(msg, *args, **kwargs):
    """记录 DEBUG 级别日志（使用根 logger）"""
    log_manager.debug(msg, *args, **kwargs)


def info(msg, *args, **kwargs):
    """记录 INFO 级别日志（使用根 logger）"""
    log_manager.info(msg, *args, **kwargs)


def warning(msg, *args, **kwargs):
    """记录 WARNING 级别日志（使用根 logger）"""
    log_manager.warning(msg, *args, **kwargs)


def error(msg, *args, **kwargs):
    """记录 ERROR 级别日志（使用根 logger）"""
    log_manager.error(msg, *args, **kwargs)


def exception(msg, *args, **kwargs):
    """记录异常信息（使用根 logger）"""
    log_manager.exception(msg, *args, **kwargs)