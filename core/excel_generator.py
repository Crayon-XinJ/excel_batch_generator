#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel生成模块 - 使用win32com操作Excel

核心职责：
    1. 打开模板文件（调用本地Excel应用程序）
    2. 应用用户配置的规则（修改单元格内容）
    3. 保存为新文件
    4. 设置文件时间属性（创建/修改/访问时间）

为什么用 win32com 而不是 openpyxl？
    - win32com 调用真实的 Excel 应用程序，完美保留所有格式
      （合并单元格、上下标、斜线、字体等）
    - openpyxl 在处理复杂格式时存在局限

注意事项：
    - 需要本地安装 Microsoft Excel
    - 生成过程中 Excel 在后台运行（不可见）
"""

import re
import os
import sys
import random
from datetime import datetime, timedelta
from typing import List, Optional

import win32com.client as win32
from win32com.client import constants
import pythoncom
import win32file
import pywintypes

from models.config_models import Rule
from core.log_manager import get_logger
from utils.helpers import (
    format_date_for_template, 
    extract_first_date_from_text, 
    extract_all_dates_from_text
)


class ExcelGenerator:
    """
    Excel 生成器
    
    每个生成任务创建一个独立的生成器实例，避免状态污染。
    """
    
    # ============================================================
    # 夜班检验员配置
    # 集中管理，方便修改
    # ============================================================
    NIGHT_SHIFT_INSPECTORS = {
        'first_half': '颜大丰',   # 1-15日
        'second_half': '张志宇',  # 16-31日
    }

    def __init__(self):
        """初始化生成器"""
        self.logger = get_logger('ExcelGenerator')
        self.template_path = None
        self.output_dir = None
        self.rules: List[Rule] = []
        self.product_name = ''
        self.template_type = ''

    # ============================================================
    # 配置方法
    # ============================================================
    
    def set_template(self, path: str) -> None:
        """设置模板文件路径"""
        self.template_path = path

    def set_output_dir(self, path: str) -> None:
        """设置输出目录"""
        self.output_dir = path
        os.makedirs(path, exist_ok=True)

    def set_rules(self, rules: List[Rule]) -> None:
        """设置规则列表（只保留已启用的规则）"""
        self.rules = [r for r in rules if r.enabled]

    def set_product_info(self, product_name: str, template_type: str) -> None:
        """设置产品信息"""
        self.product_name = product_name
        self.template_type = template_type

    # ============================================================
    # 夜班检验员
    # ============================================================
    
    def _get_night_shift_inspector(self, date: datetime) -> str:
        """根据日期获取夜班检验员"""
        if date.day <= 15:
            return self.NIGHT_SHIFT_INSPECTORS['first_half']
        else:
            return self.NIGHT_SHIFT_INSPECTORS['second_half']

    # ============================================================
    # 单元格范围解析
    # ============================================================
    
    def _parse_cell_range(self, target: str) -> List[tuple]:
        """
        解析单元格范围
        
        支持格式：
            - 行范围: "C9-L9"   → (9, 3, 9, 12)  行9，C列到L列
            - 列范围: "D6-D15"  → (6, 4, 15, 4)  列D，行6到行15
            - 离散单元格: "F9,K9" → [(9,6,9,6), (9,11,9,11)]
            - 单个单元格: "D3"   → (3, 4, 3, 4)
        
        返回格式：(起始行, 起始列, 结束行, 结束列)
        行列均为1-based索引（与Excel一致）
        
        列字母转换规则：
            A=1, B=2, ..., Z=26, AA=27, AB=28, ...
        """
        def col_letter_to_index(letters: str) -> int:
            """将列字母（如 A, Z, AA）转换为1-based列索引"""
            letters = letters.upper()
            index = 0
            for ch in letters:
                index = index * 26 + (ord(ch) - ord('A') + 1)
            return index
        
        result = []
        
        # ----- 离散单元格 (F9,K9) -----
        if ',' in target:
            parts = [p.strip() for p in target.split(',')]
            for part in parts:
                match = re.match(r'([A-Z]+)(\d+)', part, re.IGNORECASE)
                if match:
                    col_letters, row_str = match.groups()
                    row = int(row_str)
                    col = col_letter_to_index(col_letters)
                    result.append((row, col, row, col))
            return result
        
        # ----- 范围 (C9-L9 或 D6-D15) -----
        if '-' in target:
            left, right = target.split('-')
            left_match = re.match(r'([A-Z]+)(\d+)', left, re.IGNORECASE)
            right_match = re.match(r'([A-Z]+)(\d+)', right, re.IGNORECASE)
            if left_match and right_match:
                left_col_letters, left_row_str = left_match.groups()
                right_col_letters, right_row_str = right_match.groups()
                left_row = int(left_row_str)
                right_row = int(right_row_str)
                left_col = col_letter_to_index(left_col_letters)
                right_col = col_letter_to_index(right_col_letters)
                
                if left_row == right_row:
                    # 行范围 (C9-L9)
                    result.append((left_row, left_col, right_row, right_col))
                elif left_col == right_col:
                    # 列范围 (D6-D15)
                    result.append((left_row, left_col, right_row, right_col))
                else:
                    # 矩形区域，逐单元格
                    for r in range(left_row, right_row + 1):
                        for c in range(left_col, right_col + 1):
                            result.append((r, c, r, c))
                return result
        
        # ----- 单个单元格 (D3) -----
        match = re.match(r'([A-Z]+)(\d+)', target, re.IGNORECASE)
        if match:
            col_letters, row_str = match.groups()
            row = int(row_str)
            col = col_letter_to_index(col_letters)
            result.append((row, col, row, col))
        
        return result

    # ============================================================
    # 规则应用
    # ============================================================
    
    def _apply_random_rule(self, ws, rule: Rule, date: datetime) -> None:
        """应用随机数规则"""
        if rule.min_val is None or rule.max_val is None:
            return
        cells = self._parse_cell_range(rule.target)
        for row1, col1, row2, col2 in cells:
            for r in range(row1, row2 + 1):
                for c in range(col1, col2 + 1):
                    val = round(random.uniform(rule.min_val, rule.max_val), rule.decimals)
                    ws.Cells(r, c).Value = val

    def _apply_date_rule(self, ws, rule: Rule, date: datetime) -> None:
        """
        应用日期规则
        
        使用 Excel 日期序列号（浮点数）赋值，避免时区转换问题。
        Excel 日期序列号基准：1899-12-30
        """
        import datetime as dt
        EXCEL_DATE_BASE = dt.datetime(1899, 12, 30)
        
        cells = self._parse_cell_range(rule.target)
        for row1, col1, row2, col2 in cells:
            for r in range(row1, row2 + 1):
                for c in range(col1, col2 + 1):
                    cell = ws.Cells(r, c)
                    old_val = cell.Value
                    
                    # 确定显示格式
                    _, date_str, _, _ = extract_first_date_from_text(str(old_val) if old_val else '')
                    if date_str:
                        if '-' in date_str:
                            fmt = 'yyyy-mm-dd'
                        elif '/' in date_str:
                            fmt = 'yyyy/mm/dd'
                        elif '.' in date_str:
                            fmt = 'yyyy.mm.dd'
                        elif '年' in date_str:
                            fmt = 'yyyy年mm月dd日'
                        else:
                            fmt = 'yyyy-mm-dd'
                    else:
                        fmt = 'yyyy-mm-dd'
                    
                    # 计算日期序列号并赋值
                    excel_date_number = (date - EXCEL_DATE_BASE).days
                    cell.Value = float(excel_date_number)
                    cell.NumberFormat = fmt
                    
                    self.logger.debug(f"  {cell.Address}: '{old_val}' → '{date.strftime(fmt)}'")

    def _apply_text_with_date_rule(self, ws, rule: Rule, date: datetime) -> None:
        """
        应用文本+日期规则
        
        替换文本中的所有日期，保留前后缀和格式。
        从后往前替换，避免位置偏移。
        """
        cells = self._parse_cell_range(rule.target)
        for row1, col1, row2, col2 in cells:
            for r in range(row1, row2 + 1):
                for c in range(col1, col2 + 1):
                    cell = ws.Cells(r, c)
                    old_text = str(cell.Value) if cell.Value else ''
                    if not old_text:
                        continue
                    
                    matches = extract_all_dates_from_text(old_text)
                    if not matches:
                        continue
                    
                    # 从后往前替换
                    new_text = old_text
                    for date_obj, date_str, start, end in reversed(matches):
                        new_date_str = format_date_for_template(date, date_str)
                        new_text = new_text[:start] + new_date_str + new_text[end:]
                    
                    cell.NumberFormat = '@'  # 文本格式
                    cell.Value = new_text
                    self.logger.debug(f"  {cell.Address}: 替换 {len(matches)} 处日期")

    def _apply_night_shift_rule(self, ws, rule: Rule, date: datetime) -> None:
        """
        应用夜班检验员规则（过程检验表专用）
        
        1. 根据日期判断前半月/后半月
        2. 替换文本中的夜班检验员名字
        3. 同时替换日期部分
        """
        night_inspector = self._get_night_shift_inspector(date)
        
        cells = self._parse_cell_range(rule.target)
        for row1, col1, row2, col2 in cells:
            for r in range(row1, row2 + 1):
                for c in range(col1, col2 + 1):
                    cell = ws.Cells(r, c)
                    old_text = str(cell.Value) if cell.Value else ''
                    if not old_text:
                        continue
                    
                    new_text = old_text
                    
                    # 替换日期
                    date_obj, date_str, prefix, suffix = extract_first_date_from_text(old_text)
                    if date_obj is not None:
                        new_date_str = format_date_for_template(date, date_str)
                        new_text = f"{prefix}{new_date_str}{suffix}"
                    
                    # 替换夜班检验员
                    match = re.search(r'夜班检验员[：:]\s*([^\s]+)', new_text)
                    if match:
                        old_name = match.group(1)
                        new_text = new_text.replace(old_name, night_inspector)
                    
                    cell.NumberFormat = '@'
                    cell.Value = new_text

    def _apply_rules(self, ws, date: datetime) -> None:
        """应用所有规则到指定工作表"""
        ws_name = ws.Name
        for rule in self.rules:
            if rule.sheet_name and rule.sheet_name != ws_name:
                continue
            try:
                if rule.value_type == 'random':
                    self._apply_random_rule(ws, rule, date)
                elif rule.value_type == 'date':
                    self._apply_date_rule(ws, rule, date)
                elif rule.value_type == 'text_with_date':
                    self._apply_text_with_date_rule(ws, rule, date)
                elif rule.value_type == 'night_shift':
                    self._apply_night_shift_rule(ws, rule, date)
            except Exception as e:
                self.logger.warning(f"应用规则 {rule.id} 失败: {e}")

    # ============================================================
    # 文件时间设置
    # ============================================================
    
    def _set_file_time(self, file_path: str, date: datetime) -> None:
        """
        设置文件的创建时间、修改时间、访问时间
        
        规则：
            - 创建时间：日期当天 8:30~9:00 随机
            - 修改时间：创建时间 + 30~60 分钟随机
            - 访问时间：修改时间 + 0~2 分钟随机
        """
        try:
            base_date = date.replace(hour=0, minute=0, second=0, microsecond=0)
            
            # 创建时间：8:30 ~ 9:00
            create_seconds = 8 * 3600 + 30 * 60 + random.randint(0, 30 * 60)
            dt_create = base_date + timedelta(seconds=create_seconds)
            
            # 修改时间：创建时间 + 30~60 分钟
            offset_seconds = random.randint(30 * 60, 60 * 60)
            dt_modify = dt_create + timedelta(seconds=offset_seconds)
            
            # 访问时间：修改时间 + 0~2 分钟
            access_offset = random.randint(0, 2 * 60)
            dt_access = dt_modify + timedelta(seconds=access_offset)
            
            self.logger.debug(f"  文件时间: 创建 {dt_create.strftime('%H:%M:%S')}, "
                            f"修改 {dt_modify.strftime('%H:%M:%S')}, "
                            f"访问 {dt_access.strftime('%H:%M:%S')}")
            
            ct = pywintypes.Time(dt_create)
            mt = pywintypes.Time(dt_modify)
            at = pywintypes.Time(dt_access)
            
            handle = win32file.CreateFile(
                file_path,
                win32file.GENERIC_WRITE,
                win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
                None,
                win32file.OPEN_EXISTING,
                win32file.FILE_ATTRIBUTE_NORMAL,
                None
            )
            try:
                win32file.SetFileTime(handle, ct, at, mt)
            finally:
                handle.Close()
        except Exception as e:
            self.logger.warning(f"设置文件时间失败: {e}")

    # ============================================================
    # 主生成方法
    # ============================================================
    
    def generate(self, date: datetime, output_filename: str = None) -> Optional[str]:
        """
        生成一个文件
        
        Args:
            date: 表格显示的日期
            output_filename: 输出文件名（不含路径），None 则自动生成
        
        Returns:
            str: 输出文件完整路径，失败则返回 None
        """
        if not self.template_path or not os.path.exists(self.template_path):
            raise ValueError("模板文件未设置或不存在")
        if not self.output_dir:
            raise ValueError("输出目录未设置")
        
        if output_filename is None:
            output_filename = f"{self.product_name}_{self.template_type}检验表_{date.strftime('%Y%m%d')}.xlsx"
        
        output_path = os.path.join(self.output_dir, output_filename)
        
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(os.path.abspath(self.template_path))
            
            # 应用规则到所有工作表
            for ws in wb.Worksheets:
                self._apply_rules(ws, date)
            
            wb.SaveAs(os.path.abspath(output_path))
            wb.Close()
            excel.Quit()
            excel = None
            
            # 设置文件时间
            self._set_file_time(output_path, date)
            
            return output_path
            
        except Exception as e:
            raise e
        finally:
            try:
                if excel:
                    excel.Quit()
            except:
                pass