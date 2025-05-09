#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
工具函数模块
提供Excel范围解析等辅助功能
"""

from openpyxl.utils import get_column_letter, column_index_from_string


def parse_range_string(range_str):
    """
    解析范围字符串
    支持格式：
    - 单个值：1 或 A
    - 多个值：1,2,3 或 A,B,C
    - 范围值：1:3 或 A:C
    - 混合格式：1,3:5 或 A,C:E
    - 支持中文符号：1，3：5 或 A，C：E
    """
    # 将中文符号转换为英文符号
    range_str = range_str.replace('，', ',').replace('：', ':')
    
    # 分割多个范围
    parts = [p.strip() for p in range_str.split(',')]
    result = []
    
    is_number = parts[0].replace(':', '').isdigit()
    
    for part in parts:
        if ':' in part:
            start, end = [p.strip() for p in part.split(':')]
            if is_number:
                try:
                    start = int(start)
                    end = int(end)
                    if start <= 0:
                        raise ValueError("行号必须大于0")
                    if start > end:
                        raise ValueError("起始行必须小于或等于结束行")
                except ValueError as e:
                    if "行号必须大于0" in str(e):
                        raise ValueError("行号必须大于0")
                    raise ValueError("行号必须是数字")
            else:
                if not (start.isalpha() and end.isalpha()):
                    raise ValueError("列标识必须是字母")
                start = start.upper()
                end = end.upper()
                if start > end:
                    raise ValueError("起始列必须在结束列之前")
            result.append((start, end))
        else:
            if is_number:
                try:
                    value = int(part)
                    if value <= 0:
                        raise ValueError("行号必须大于0")
                except ValueError:
                    raise ValueError("行号必须是数字")
            else:
                if not part.isalpha():
                    raise ValueError("列标识必须是字母")
                part = part.upper()
            result.append(part)
    
    return result


def convert_to_column_index(column_str):
    """
    将列标识转换为列索引
    
    Args:
        column_str: 列标识（如'A'、'B'等）
    
    Returns:
        int: 列索引（从1开始）
    """
    return column_index_from_string(column_str)


def convert_to_column_letter(column_index):
    """
    将列索引转换为列标识
    
    Args:
        column_index: 列索引（从1开始）
    
    Returns:
        str: 列标识（如'A'、'B'等）
    """
    return get_column_letter(column_index)


def parse_cell_range(range_str):
    """
    解析单元格范围字符串
    支持格式：'A1:B2'、'A1'
    
    Args:
        range_str: 单元格范围字符串
    
    Returns:
        tuple: (min_row, min_col, max_row, max_col)
    """
    if ':' in range_str:
        start, end = range_str.split(':')
    else:
        start = end = range_str
    
    # 分离列标识和行号
    start_col = ''.join(c for c in start if c.isalpha())
    start_row = int(''.join(c for c in start if c.isdigit()))
    end_col = ''.join(c for c in end if c.isalpha())
    end_row = int(''.join(c for c in end if c.isdigit()))
    
    # 转换列标识为列索引
    start_col_idx = column_index_from_string(start_col)
    end_col_idx = column_index_from_string(end_col)
    
    return (start_row, start_col_idx, end_row, end_col_idx)


def validate_position_input(position_str, is_row=True):
    """验证行列位置输入的有效性
    
    Args:
        position_str: 用户输入的位置字符串
        is_row: 是否为行验证(True为行，False为列)
        
    Returns:
        tuple: (是否有效, 错误信息)
    """
    if not position_str:
        return False, "输入不能为空"
    
    # 将中文符号转换为英文符号
    position_str = position_str.replace('，', ',').replace('：', ':')
    
    try:
        if is_row:
            # 行验证逻辑
            parts = position_str.split(',')
            for part in parts:
                if ':' in part:
                    start, end = part.split(':')
                    if not (start.isdigit() and end.isdigit()):
                        return False, "行号必须为数字"
                    if int(start) > int(end):
                        return False, "起始行号不能大于结束行号"
                else:
                    if not part.isdigit():
                        return False, "行号必须为数字"
        else:
            # 列验证逻辑
            parts = position_str.split(',')
            for part in parts:
                if ':' in part:
                    start, end = part.split(':')
                    if not (start.isalpha() and end.isalpha()):
                        return False, "列标识必须为字母"
                    if len(start) != len(end):
                        return False, "列标识长度必须一致"
                    if start > end:
                        return False, "起始列不能大于结束列"
                else:
                    if not part.isalpha():
                        return False, "列标识必须为字母"
                        
        return True, ""
    except Exception as e:
        return False, f"验证出错: {str(e)}"