#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
邮件报告模块 - 生成和发送邮件报告
"""

import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from typing import Dict, List, Optional

from core.log_manager import get_logger
from core.config_manager import ConfigManager


class EmailReporter:
    """邮件报告生成器"""
    
    def __init__(self):
        self.logger = get_logger('EmailReporter')
        self.config_manager = ConfigManager()
        self.email_config = self.config_manager.get_email_config()
    
    def build_report(self, tasks_results: Dict, stats: Dict, date: datetime = None) -> str:
        """构建邮件报告内容"""
        if date is None:
            date = datetime.now()
        
        lines = []
        lines.append("=" * 60)
        lines.append("  Excel批量生成工具 - 每日生成报告")
        lines.append("=" * 60)
        lines.append(f"生成时间: {date.strftime('%Y-%m-%d %H:%M:%S')}")
        if 'duration' in stats:
            lines.append(f"耗时: {stats['duration']}")
        lines.append("-" * 60)
        
        # 统计汇总
        lines.append("📊 统计汇总")
        lines.append(f"  - 总文件数: {stats.get('total', 0)}")
        lines.append(f"  - ✅ 成功: {stats.get('success', 0)}")
        lines.append(f"  - ⏭️ 跳过: {stats.get('skipped', 0)}")
        lines.append(f"  - ❌ 失败: {stats.get('failed', 0)}")
        lines.append("-" * 60)
        
        # 生成详情
        lines.append("📁 生成详情")
        for product, results in tasks_results.items():
            lines.append(f"  {product}:")
            for task_type, info in results.items():
                if task_type in ['首件', '过程', '成品'] and info:
                    if isinstance(info, dict):
                        status = info.get('status', '')
                        filename = info.get('filename', '')
                        if status == 'success':
                            lines.append(f"    ✅ {task_type}: {filename}")
                        elif status == 'skipped':
                            lines.append(f"    ⏭️ {task_type}: 已存在")
                        elif status == 'failed':
                            lines.append(f"    ❌ {task_type}: {info.get('error', '未知错误')}")
                    else:
                        lines.append(f"    ✅ {task_type}: 已生成")
        
        lines.append("-" * 60)
        if 'output_dir' in stats:
            lines.append(f"🔗 文件位置: {stats['output_dir']}")
        lines.append("=" * 60)
        
        return "\n".join(lines)
    
    def send(self, subject: str, body: str) -> bool:
        """发送邮件"""
        if not self.email_config.get('enabled', False):
            self.logger.debug("邮件通知未启用")
            return False
        
        config = self.email_config
        
        if not config.get('username') or not config.get('password'):
            self.logger.warning("邮件配置不完整（缺少用户名或密码），跳过发送")
            return False
        
        if not config.get('to'):
            self.logger.warning("邮件配置不完整（缺少收件人），跳过发送")
            return False
        
        try:
            # 构建邮件
            msg = MIMEMultipart()
            msg['From'] = config.get('username', '')
            msg['To'] = config.get('to', '')
            if config.get('cc'):
                msg['Cc'] = config.get('cc', '')
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # 获取收件人列表
            recipients = [config.get('to', '')]
            if config.get('cc'):
                recipients.extend([c.strip() for c in config.get('cc', '').split(',') if c.strip()])
            
            # 创建SSL连接并发送
            if config.get('smtp_ssl', True):
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(config.get('smtp_server', ''), config.get('smtp_port', 465), context=context) as server:
                    server.login(config.get('username', ''), config.get('password', ''))
                    server.sendmail(config.get('username', ''), recipients, msg.as_string())
            else:
                with smtplib.SMTP(config.get('smtp_server', ''), config.get('smtp_port', 25)) as server:
                    if config.get('smtp_tls', False):
                        server.starttls()
                    server.login(config.get('username', ''), config.get('password', ''))
                    server.sendmail(config.get('username', ''), recipients, msg.as_string())
            
            self.logger.info(f"邮件已发送: {subject}")
            return True
            
        except Exception as e:
            self.logger.error(f"邮件发送失败: {e}")
            return False
    
    def test_connection(self) -> tuple:
        """测试邮件连接，返回 (success, message)"""
        config = self.email_config
        
        if not config.get('username') or not config.get('password'):
            return False, "邮件配置不完整（缺少用户名或密码）"
        
        if not config.get('smtp_server'):
            return False, "邮件配置不完整（缺少SMTP服务器）"
        
        try:
            if config.get('smtp_ssl', True):
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(config.get('smtp_server', ''), config.get('smtp_port', 465), context=context) as server:
                    server.login(config.get('username', ''), config.get('password', ''))
            else:
                with smtplib.SMTP(config.get('smtp_server', ''), config.get('smtp_port', 25)) as server:
                    if config.get('smtp_tls', False):
                        server.starttls()
                    server.login(config.get('username', ''), config.get('password', ''))
            return True, "连接成功"
        except Exception as e:
            return False, f"连接失败: {str(e)}"