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

import os
import sys
import random
import re
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
    """Excel生成器"""
    
    NIGHT_SHIFT_INSPECTORS = {
        'first_half': '颜大丰',   # 1-15日
        'second_half': '张志宇',  # 16-31日
    }

    def __init__(self):
        self.logger = get_logger('ExcelGenerator')
        self.template_path = None
        self.output_dir = None
        self.rules: List[Rule] = []
        self.product_name = ''
        self.template_type = ''

    def set_template(self, path: str) -> None:
        self.template_path = path

    def set_output_dir(self, path: str) -> None:
        self.output_dir = path
        os.makedirs(path, exist_ok=True)

    def set_rules(self, rules: List[Rule]) -> None:
        self.rules = [r for r in rules if r.enabled]

    def set_product_info(self, product_name: str, template_type: str) -> None:
        self.product_name = product_name
        self.template_type = template_type

    def _get_night_shift_inspector(self, date: datetime) -> str:
        if date.day <= 15:
            return self.NIGHT_SHIFT_INSPECTORS['first_half']
        else:
            return self.NIGHT_SHIFT_INSPECTORS['second_half']

    def _parse_cell_range(self, target: str) -> List[tuple]:
        """解析单元格范围"""
        def col_letter_to_index(letters: str) -> int:
            letters = letters.upper()
            index = 0
            for ch in letters:
                index = index * 26 + (ord(ch) - ord('A') + 1)
            return index
        
        result = []
        
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
                    result.append((left_row, left_col, right_row, right_col))
                elif left_col == right_col:
                    result.append((left_row, left_col, right_row, right_col))
                else:
                    for r in range(left_row, right_row + 1):
                        for c in range(left_col, right_col + 1):
                            result.append((r, c, r, c))
                return result
        
        match = re.match(r'([A-Z]+)(\d+)', target, re.IGNORECASE)
        if match:
            col_letters, row_str = match.groups()
            row = int(row_str)
            col = col_letter_to_index(col_letters)
            result.append((row, col, row, col))
        
        return result

    def _apply_random_rule(self, ws, rule: Rule, date: datetime) -> None:
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
        
        使用 Excel 日期序列号赋值，从单元格 NumberFormat 读取格式。
        兼容 EXE 环境中 cell.Value 返回数值而非 datetime 的情况。
        """
        import datetime as dt
        EXCEL_DATE_BASE = dt.datetime(1899, 12, 30)
        
        cells = self._parse_cell_range(rule.target)
        for row1, col1, row2, col2 in cells:
            for r in range(row1, row2 + 1):
                for c in range(col1, col2 + 1):
                    cell = ws.Cells(r, c)
                    old_val = cell.Value
                    
                    # ============================================================
                    # 从单元格的 NumberFormat 获取日期显示格式
                    # ============================================================
                    fmt = cell.NumberFormat
                    
                    # 如果格式无效或为常规格式，使用默认格式
                    if not fmt or fmt == 'General':
                        fmt = 'yyyy-mm-dd'
                    elif '/' in fmt:
                        # 保留原始分隔符格式
                        fmt = 'yyyy/mm/dd'
                    elif '.' in fmt:
                        fmt = 'yyyy.mm.dd'
                    elif '年' in fmt and '月' in fmt and '日' in fmt:
                        fmt = 'yyyy年mm月dd日'
                    else:
                        # 如果有其他日期格式但未匹配，保留原格式
                        # 但某些格式可能包含 'm-d' 等，直接使用
                        pass
                    
                    # 计算 Excel 日期序列号并赋值
                    excel_date_number = (date - EXCEL_DATE_BASE).days
                    cell.Value = float(excel_date_number)
                    cell.NumberFormat = fmt
                    
                    self.logger.debug(f"  {cell.Address}: '{old_val}' → '{date.strftime(fmt)}'")

    def _apply_text_with_date_rule(self, ws, rule: Rule, date: datetime) -> None:
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
                    
                    new_text = old_text
                    for date_obj, date_str, start, end in reversed(matches):
                        new_date_str = format_date_for_template(date, date_str)
                        new_text = new_text[:start] + new_date_str + new_text[end:]
                    
                    cell.NumberFormat = '@'
                    cell.Value = new_text
                    self.logger.debug(f"  {cell.Address}: 替换 {len(matches)} 处日期")

    def _apply_night_shift_rule(self, ws, rule: Rule, date: datetime) -> None:
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
                    
                    date_obj, date_str, prefix, suffix = extract_first_date_from_text(old_text)
                    if date_obj is not None:
                        new_date_str = format_date_for_template(date, date_str)
                        new_text = f"{prefix}{new_date_str}{suffix}"
                    
                    match = re.search(r'夜班检验员[：:]\s*([^\s]+)', new_text)
                    if match:
                        old_name = match.group(1)
                        new_text = new_text.replace(old_name, night_inspector)
                    
                    cell.NumberFormat = '@'
                    cell.Value = new_text

    def _apply_rules(self, ws, date: datetime) -> None:
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

    def _set_file_time(self, file_path: str, date: datetime) -> None:
        try:
            base_date = date.replace(hour=0, minute=0, second=0, microsecond=0)
            
            create_seconds = 8 * 3600 + 30 * 60 + random.randint(0, 30 * 60)
            dt_create = base_date + timedelta(seconds=create_seconds)
            
            offset_seconds = random.randint(30 * 60, 60 * 60)
            dt_modify = dt_create + timedelta(seconds=offset_seconds)
            
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

    def generate(self, date: datetime, output_filename: str = None) -> Optional[str]:
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
            
            for ws in wb.Worksheets:
                self._apply_rules(ws, date)
            
            wb.SaveAs(os.path.abspath(output_path))
            wb.Close()
            excel.Quit()
            excel = None
            
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