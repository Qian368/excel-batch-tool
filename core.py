#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel处理核心模块
提供Excel文件的各种批量处理功能
"""

import os
import traceback
from pathlib import Path
import shutil
from datetime import datetime

import openpyxl
from openpyxl.utils import get_column_letter


class ExcelProcessor:
    """
    Excel处理核心类，提供各种Excel批量处理功能
    """
    
    def __init__(self):
        self.workbooks = {}
        self.output_dir = None
        self.temp_files = {}
        
        # 清理可能存在的临时文件
        self._cleanup_temp_files()
    
    def set_output_dir(self, output_dir):
        """
        设置Excel文件的输出目录
        
        Args:
            output_dir: 输出目录路径
        """
        if not output_dir:
            raise ValueError("输出目录不能为空")
            
        output_path = Path(output_dir)
        if not output_path.exists():
            output_path.mkdir(parents=True)
            
        self.output_dir = str(output_path)
    
    def load_workbooks(self, file_paths):
        """
        加载Excel工作簿，并复制到临时工作区
        
        Args:
            file_paths: Excel文件路径列表
        """
        for file_path in file_paths:
            try:
                # 创建临时文件副本
                temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                shutil.copy2(file_path, temp_path)
                self.temp_files[file_path] = temp_path
                
                # 加载临时文件的工作簿
                wb = openpyxl.load_workbook(temp_path, data_only=False)
                self.workbooks[file_path] = wb
            except Exception as e:
                print(f"加载文件 {file_path} 失败: {str(e)}")
                raise
    
    def save_workbooks(self):
        """
        保存所有工作簿到指定的输出目录
        """
        if not self.output_dir:
            raise ValueError("未指定输出目录")
            
        # 确保输出目录存在且可写
        output_path = Path(self.output_dir)
        if not output_path.exists():
            try:
                output_path.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                raise PermissionError(f"无法创建输出目录 {self.output_dir}: {str(e)}")
                
        if not os.access(self.output_dir, os.W_OK):
            raise PermissionError(f"输出目录 {self.output_dir} 不可写")
            
        for file_path, wb in self.workbooks.items():
            try:
                # 获取文件名并构建新的保存路径
                file_name = Path(file_path).name
                output_path = Path(self.output_dir) / file_name
                
                # 检查目标文件是否已存在且可写
                if output_path.exists() and not os.access(str(output_path), os.W_OK):
                    raise PermissionError(f"目标文件 {output_path} 已存在且不可写")
                    
                # 保存工作簿
                wb.save(output_path)
                
            except PermissionError as e:
                print(f"权限错误: {str(e)}")
                raise
            except Exception as e:
                print(f"保存文件 {file_path} 失败: {str(e)}")
                raise
            finally:
                # 清理临时文件
                self._cleanup_temp_file(file_path)
    
    def convert_formulas_to_values(self, file_paths):
        """
        将公式转换为值
        
        Args:
            file_paths: 要处理的Excel文件路径列表
        """
        for file_path in file_paths:
            try:
                # 确保使用工作簿副本，而不是直接修改原始工作簿
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 先用data_only=True加载以计算公式
                    wb_data = openpyxl.load_workbook(temp_path, data_only=True)
                    
                    # 再用data_only=False加载以获取原始内容
                    wb_formula = openpyxl.load_workbook(temp_path, data_only=False)
                else:
                    # 如果已经有工作簿副本，则使用它
                    # 获取临时文件路径
                    temp_path = self.temp_files.get(file_path)
                    if not temp_path:
                        # 如果没有临时文件，创建一个
                        temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                        shutil.copy2(file_path, temp_path)
                        self.temp_files[file_path] = temp_path
                    
                    # 先用data_only=True加载以计算公式
                    wb_data = openpyxl.load_workbook(temp_path, data_only=True)
                    
                    # 使用已有的工作簿
                    wb_formula = wb
                
                for sheet_name in wb_formula.sheetnames:
                    sheet_formula = wb_formula[sheet_name]
                    sheet_data = wb_data[sheet_name]
                    
                    for row in range(1, sheet_formula.max_row + 1):
                        for col in range(1, sheet_formula.max_column + 1):
                            cell_formula = sheet_formula.cell(row=row, column=col)
                            cell_data = sheet_data.cell(row=row, column=col)
                            
                            if cell_formula.data_type == 'f':  # 如果是公式
                                cell_formula.value = cell_data.value
                
                # 保存修改后的工作簿到临时文件
                if temp_path:
                    wb_formula.save(temp_path)
                self.workbooks[file_path] = wb_formula
                
            except Exception as e:
                print(f"处理文件 {file_path} 失败: {str(e)}")
                raise
    
    def _cleanup_temp_files(self):
        """
        清理所有临时文件
        """
        for file_path, temp_path in list(self.temp_files.items()):
            try:
                if Path(temp_path).exists():
                    Path(temp_path).unlink()
                del self.temp_files[file_path]
            except Exception as e:
                print(f"清理临时文件 {temp_path} 失败: {str(e)}")
    
    def _cleanup_temp_file(self, file_path):
        """
        清理指定文件的临时文件
        
        Args:
            file_path: 原始文件路径
        """
        if file_path in self.temp_files:
            try:
                temp_path = self.temp_files[file_path]
                if Path(temp_path).exists():
                    Path(temp_path).unlink()
                del self.temp_files[file_path]
            except Exception as e:
                print(f"清理临时文件 {self.temp_files[file_path]} 失败: {str(e)}")
    
    def _get_data_range(self, sheet):
        """
        获取工作表中有数据的范围
        
        Args:
            sheet: 工作表对象
            
        Returns:
            tuple: (min_row, max_row, min_col, max_col)
        """
        min_row = min_col = float('inf')
        max_row = max_col = 0
        
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value is not None:
                    min_row = min(min_row, row)
                    max_row = max(row, max_row)
                    min_col = min(min_col, col)
                    max_col = max(col, max_col)
        
        return (min_row if min_row != float('inf') else 1,
                max_row if max_row != 0 else sheet.max_row,
                min_col if min_col != float('inf') else 1,
                max_col if max_col != 0 else sheet.max_column)
    
    def get_intersection(self, range1, range2):
        """
        计算两个单元格范围的交集
        
        Args:
            range1: 元组 (min_row1, min_col1, max_row1, max_col1)
            range2: 元组 (min_row2, min_col2, max_row2, max_col2)
        
        Returns:
            tuple: 交集区域的范围 (min_row, min_col, max_row, max_col) 或 None（如果没有交集）
        """
        min_row1, min_col1, max_row1, max_col1 = range1
        min_row2, min_col2, max_row2, max_col2 = range2
        
        # 如果一个范围在另一个范围的完全左边或右边，则没有交集
        if max_col1 < min_col2 or min_col1 > max_col2:
            return None
        
        # 如果一个范围在另一个范围的完全上边或下边，则没有交集
        if max_row1 < min_row2 or min_row1 > max_row2:
            return None
        
        # 计算交集区域
        intersection_min_row = max(min_row1, min_row2)
        intersection_min_col = max(min_col1, min_col2)
        intersection_max_row = min(max_row1, max_row2)
        intersection_max_col = min(max_col1, max_col2)
        
        return (intersection_min_row, intersection_min_col, intersection_max_row, intersection_max_col)
    
    def get_range_from_reference(self, range_ref):
        """
        从单元格引用获取范围的行列数值
        
        Args:
            range_ref: 字符串，如'A1:B2'或'A1'
            
        Returns:
            tuple: (min_row, min_col, max_row, max_col)
        """
        from openpyxl.utils import range_boundaries
        
        try:
            # 使用openpyxl的函数解析单元格范围
            min_col, min_row, max_col, max_row = range_boundaries(range_ref)
            return (min_row, min_col, max_row, max_col)
        except Exception:
            return None
    
    def find_intersections(self, ws, target_range):
        """
        查找与目标范围相交的合并单元格
        
        Args:
            ws: openpyxl工作表对象
            target_range: 目标单元格范围的引用字符串
        
        Returns:
            list: 相交的合并单元格列表，每个元素为(合并单元格范围, 交集范围)元组
        """
        target = self.get_range_from_reference(target_range)
        
        if not target:
            return []
            
        # 查找与目标范围有交集的合并单元格
        intersections = []
        for merged_range in ws.merged_cells.ranges:
            current_range = (
                merged_range.min_row,
                merged_range.min_col,
                merged_range.max_row,
                merged_range.max_col
            )
            
            intersection = self.get_intersection(target, current_range)
            if intersection:
                # 将交集区域转换为Excel引用格式
                from openpyxl.utils import get_column_letter
                intersection_start = f"{get_column_letter(intersection[1])}{intersection[0]}"
                intersection_end = f"{get_column_letter(intersection[3])}{intersection[2]}"
                intersection_range = intersection_start if intersection_start == intersection_end else f"{intersection_start}:{intersection_end}"
                
                intersections.append((merged_range, intersection_range))
        
        return intersections
    
    def _get_merged_cells_in_range(self, sheet, start_row, end_row, start_col, end_col):
        """
        获取指定范围内的合并单元格
        
        Args:
            sheet: 工作表对象
            start_row: 起始行
            end_row: 结束行
            start_col: 起始列
            end_col: 结束列
            
        Returns:
            list: 合并单元格范围列表
        """
        merged_ranges = []
        for merged_range in sheet.merged_cells.ranges:
            # 检查合并单元格是否与指定范围有交集
            if not (merged_range.max_row < start_row or 
                    merged_range.min_row > end_row or 
                    merged_range.max_col < start_col or 
                    merged_range.min_col > end_col):
                merged_ranges.append(merged_range)
        return merged_ranges
    
    def process_merged_cells(self, file_paths, action='unmerge', mode='all', range_str=None):
        """
        处理合并单元格（合并或拆分）
        
        Args:
            file_paths: 要处理的Excel文件路径列表
            action: 操作类型，'merge'表示合并，'unmerge'表示拆分，'keep_value'表示保留值但拆分
            mode: 处理模式，'all'表示处理整个工作表，'specific'表示处理指定范围
            range_str: 当mode为'specific'时，指定要处理的单元格范围
        """
        for file_path in file_paths:
            try:
                # 确保使用工作簿副本，而不是直接修改原始工作簿
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                for sheet_name in wb.sheetnames:
                    sheet = wb[sheet_name]
                    
                    if mode == 'specific' and range_str:
                        # 处理指定范围的合并单元格
                        # 解析单元格范围
                        try:
                            # 处理单个单元格的情况
                            if ':' not in range_str:
                                # 解析单个单元格坐标
                                cell_coord = openpyxl.utils.cell.coordinate_to_tuple(range_str)
                                min_row, min_col = cell_coord
                                max_row, max_col = cell_coord
                            else:
                                # 处理范围的情况
                                start_cell, end_cell = range_str.split(':')
                                start_coord = openpyxl.utils.cell.coordinate_to_tuple(start_cell)
                                end_coord = openpyxl.utils.cell.coordinate_to_tuple(end_cell)
                                
                                min_row, min_col = start_coord
                                max_row, max_col = end_coord
                            
                            # 找出所有与指定范围有交集的合并单元格
                            overlapping_ranges = []
                            for merged_range in list(sheet.merged_cells.ranges):
                                merged_min_row, merged_min_col = merged_range.min_row, merged_range.min_col
                                merged_max_row, merged_max_col = merged_range.max_row, merged_range.max_col
                                
                                # 检查是否有重叠
                                if not (merged_max_row < min_row or merged_min_row > max_row or 
                                        merged_max_col < min_col or merged_min_col > max_col):
                                    overlapping_ranges.append(merged_range)
                            
                            # 如果没有找到任何重叠的合并单元格，则报告错误
                            # 但是不要立即抛出异常，而是返回一个标志，让调用者决定如何处理
                            if not overlapping_ranges:
                                # 这里不再抛出异常，而是返回一个空列表，表示没有找到重叠的合并单元格
                                # 这样processing.py中的代码可以继续处理其他文件
                                continue
                            
                            # 拆分所有重叠的合并单元格
                            for merged_range in overlapping_ranges:
                                # 如果需要保留值，先获取左上角单元格的值
                                top_left_value = None
                                if action == 'keep_value':
                                    top_left_value = sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
                                
                                # 拆分合并单元格
                                sheet.unmerge_cells(str(merged_range))
                                
                                # 如果需要保留值，填充到所有单元格
                                if action == 'keep_value' and top_left_value is not None:
                                    for row in range(merged_range.min_row, merged_range.max_row + 1):
                                        for col in range(merged_range.min_col, merged_range.max_col + 1):
                                            sheet.cell(row=row, column=col).value = top_left_value
                                
                        except ValueError as ve:
                            # 传递自定义的ValueError
                            raise ve
                        except Exception as e:
                            # 如果解析单元格范围出错，抛出更具体的错误
                            raise ValueError(f"无效的单元格范围格式: {range_str}, 错误: {str(e)}")
                        
                        # 注意：合并单元格的值保留逻辑已经在拆分合并单元格的循环中处理了
                    else:
                        # 处理整个工作表的所有合并单元格
                        merged_ranges = list(sheet.merged_cells.ranges)
                        
                        for merged_range in merged_ranges:
                            # 获取左上角单元格的值
                            top_left_value = sheet.cell(
                                row=merged_range.min_row, 
                                column=merged_range.min_col
                            ).value
                            
                            # 拆分合并单元格
                            sheet.unmerge_cells(str(merged_range))
                            
                            if action == 'keep_value':
                                # 填充到所有单元格
                                for row in range(merged_range.min_row, merged_range.max_row + 1):
                                    for col in range(merged_range.min_col, merged_range.max_col + 1):
                                        sheet.cell(row=row, column=col).value = top_left_value
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                else:
                    # 如果没有临时文件路径，这是一个错误情况
                    raise ValueError(f"找不到文件 {file_path} 的临时文件路径")
                
            except Exception as e:
                print(f"处理文件 {file_path} 失败: {str(e)}")
                raise
    
    def _process_merged_cells_in_range(self, sheet, merged_ranges, merge_mode='ignore'):
        """
        根据不同模式处理指定范围内的合并单元格
        
        Args:
            sheet: 工作表对象
            merged_ranges: 合并单元格范围列表
            merge_mode: 合并单元格处理模式，可选值：
                       - ignore: 忽略合并单元格，直接执行操作
                       - unmerge_only: 仅拆分合并单元格，不保留值
                       - unmerge_keep_value: 拆分合并单元格并保留值
        """
        if merge_mode == 'ignore':
            return
            
        for merged_range in merged_ranges:
            # 获取左上角单元格的值（仅在需要保留值时获取）
            top_left_value = None
            if merge_mode == 'unmerge_keep_value':
                top_left_value = sheet.cell(
                    row=merged_range.min_row,
                    column=merged_range.min_col
                ).value
            
            # 拆分合并单元格
            sheet.unmerge_cells(str(merged_range))
            
            # 如果是unmerge_keep_value模式，则填充值到所有单元格
            if merge_mode == 'unmerge_keep_value' and top_left_value is not None:
                for row in range(merged_range.min_row, merged_range.max_row + 1):
                    for col in range(merged_range.min_col, merged_range.max_col + 1):
                        sheet.cell(row=row, column=col).value = top_left_value
            # 如果是unmerge_only模式，则清空所有单元格的值
            elif merge_mode == 'unmerge_only':
                for row in range(merged_range.min_row, merged_range.max_row + 1):
                    for col in range(merged_range.min_col, merged_range.max_col + 1):
                        sheet.cell(row=row, column=col).value = None
    
    def insert_rows(self, file_paths, sheet_indexes, position, count=1):
        """
        插入行
        
        Args:
            file_paths: Excel文件路径列表
            sheet_indexes: 工作表索引列表
            position: 插入位置（行号）
            count: 插入行数
        """
        for file_path in file_paths:
            try:
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                for sheet_index in sheet_indexes:
                    if isinstance(sheet_index, int) and 0 <= sheet_index < len(wb.sheetnames):
                        sheet = wb[wb.sheetnames[sheet_index]]
                        sheet.insert_rows(position, count)
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                else:
                    # 如果没有临时文件路径，这是一个错误情况
                    raise ValueError(f"找不到文件 {file_path} 的临时文件路径")
                
            except Exception as e:
                print(f"在文件 {file_path} 中插入行失败: {str(e)}")
                raise
    
    def delete_rows(self, file_paths, sheet_indexes, position, count=1, merge_mode='ignore'):
        """
        删除行
        
        Args:
            file_paths: Excel文件路径列表
            sheet_indexes: 工作表索引列表
            position: 删除位置（行号）
            count: 删除行数
            merge_mode: 合并单元格处理模式
        """
        for file_path in file_paths:
            try:
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                for sheet_index in sheet_indexes:
                    if isinstance(sheet_index, int) and 0 <= sheet_index < len(wb.sheetnames):
                        sheet = wb[wb.sheetnames[sheet_index]]
                        
                        # 获取数据范围
                        data_range = self._get_data_range(sheet)
                        
                        # 获取要删除的行范围内的合并单元格
                        merged_ranges = self._get_merged_cells_in_range(
                            sheet, position, position + count - 1,
                            data_range[2], data_range[3]
                        )
                        
                        # 处理合并单元格
                        self._process_merged_cells_in_range(sheet, merged_ranges, merge_mode)
                        
                        # 删除行
                        sheet.delete_rows(position, count)
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                else:
                    # 如果没有临时文件路径，这是一个错误情况
                    raise ValueError(f"找不到文件 {file_path} 的临时文件路径")
                
            except Exception as e:
                print(f"在文件 {file_path} 中删除行失败: {str(e)}")
                raise
    
    def insert_columns(self, file_paths, sheet_indexes, position, count=1):
        """
        插入列
        
        Args:
            file_paths: Excel文件路径列表
            sheet_indexes: 工作表索引列表
            position: 插入位置（列号）
            count: 插入列数
        """
        for file_path in file_paths:
            try:
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                for sheet_index in sheet_indexes:
                    if isinstance(sheet_index, int) and 0 <= sheet_index < len(wb.sheetnames):
                        sheet = wb[wb.sheetnames[sheet_index]]
                        sheet.insert_cols(position, count)
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                else:
                    # 如果没有临时文件路径，这是一个错误情况
                    raise ValueError(f"找不到文件 {file_path} 的临时文件路径")
                
            except Exception as e:
                print(f"在文件 {file_path} 中插入列失败: {str(e)}")
                raise
    
    def delete_columns(self, file_paths, sheet_indexes, position, count=1, merge_mode='ignore'):
        """
        删除列
        
        Args:
            file_paths: Excel文件路径列表
            sheet_indexes: 工作表索引列表
            position: 删除位置（列号）
            count: 删除列数
            merge_mode: 合并单元格处理模式
        """
        for file_path in file_paths:
            try:
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                for sheet_index in sheet_indexes:
                    if isinstance(sheet_index, int) and 0 <= sheet_index < len(wb.sheetnames):
                        sheet = wb[wb.sheetnames[sheet_index]]
                        
                        # 获取数据范围
                        data_range = self._get_data_range(sheet)
                        
                        # 获取要删除的列范围内的合并单元格
                        merged_ranges = self._get_merged_cells_in_range(
                            sheet, data_range[0], data_range[1],
                            position, position + count - 1
                        )
                        
                        # 处理合并单元格
                        self._process_merged_cells_in_range(sheet, merged_ranges, merge_mode)
                        
                        # 删除列
                        sheet.delete_cols(position, count)
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                else:
                    # 如果没有临时文件路径，这是一个错误情况
                    raise ValueError(f"找不到文件 {file_path} 的临时文件路径")
                
            except Exception as e:
                print(f"在文件 {file_path} 中删除列失败: {str(e)}")
                raise
    
    def hide_rows(self, file_paths, sheet_indexes, position, count=1):
        """
        隐藏行
        
        Args:
            file_paths: Excel文件路径列表
            sheet_indexes: 工作表索引列表
            position: 隐藏位置（行号）
            count: 隐藏行数
        """
        for file_path in file_paths:
            try:
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                for sheet_index in sheet_indexes:
                    if isinstance(sheet_index, int) and 0 <= sheet_index < len(wb.sheetnames):
                        sheet = wb[wb.sheetnames[sheet_index]]
                        
                        # 隐藏指定范围的行
                        for row in range(position, position + count):
                            sheet.row_dimensions[row].hidden = True
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                else:
                    # 如果没有临时文件路径，这是一个错误情况
                    raise ValueError(f"找不到文件 {file_path} 的临时文件路径")
                
            except Exception as e:
                print(f"在文件 {file_path} 中隐藏行失败: {str(e)}")
                raise
    
    def unhide_rows(self, file_paths, sheet_indexes, position, count=1):
        """
        显示行
        
        Args:
            file_paths: Excel文件路径列表
            sheet_indexes: 工作表索引列表
            position: 显示位置（行号）
            count: 显示行数
        """
        for file_path in file_paths:
            try:
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                for sheet_index in sheet_indexes:
                    if isinstance(sheet_index, int) and 0 <= sheet_index < len(wb.sheetnames):
                        sheet = wb[wb.sheetnames[sheet_index]]
                        
                        # 显示指定范围的行
                        for row in range(position, position + count):
                            sheet.row_dimensions[row].hidden = False
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                else:
                    # 如果没有临时文件路径，这是一个错误情况
                    raise ValueError(f"找不到文件 {file_path} 的临时文件路径")
                
            except Exception as e:
                print(f"在文件 {file_path} 中显示行失败: {str(e)}")
                raise
    
    def hide_columns(self, file_paths, sheet_indexes, position, count=1):
        """
        隐藏列
        
        Args:
            file_paths: Excel文件路径列表
            sheet_indexes: 工作表索引列表
            position: 隐藏位置（列号）
            count: 隐藏列数
        """
        for file_path in file_paths:
            try:
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                for sheet_index in sheet_indexes:
                    if isinstance(sheet_index, int) and 0 <= sheet_index < len(wb.sheetnames):
                        sheet = wb[wb.sheetnames[sheet_index]]
                        
                        # 隐藏指定范围的列
                        for col in range(position, position + count):
                            col_letter = get_column_letter(col)
                            sheet.column_dimensions[col_letter].hidden = True
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                else:
                    # 如果没有临时文件路径，这是一个错误情况
                    raise ValueError(f"找不到文件 {file_path} 的临时文件路径")
                
            except Exception as e:
                print(f"在文件 {file_path} 中隐藏列失败: {str(e)}")
                raise
    
    def unhide_columns(self, file_paths, sheet_indexes, position, count=1):
        """
        显示列
        
        Args:
            file_paths: Excel文件路径列表
            sheet_indexes: 工作表索引列表
            position: 显示位置（列号）
            count: 显示列数
        """
        for file_path in file_paths:
            try:
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                for sheet_index in sheet_indexes:
                    if isinstance(sheet_index, int) and 0 <= sheet_index < len(wb.sheetnames):
                        sheet = wb[wb.sheetnames[sheet_index]]
                        
                        # 显示指定范围的列
                        for col in range(position, position + count):
                            col_letter = get_column_letter(col)
                            sheet.column_dimensions[col_letter].hidden = False
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                else:
                    # 如果没有临时文件路径，这是一个错误情况
                    raise ValueError(f"找不到文件 {file_path} 的临时文件路径")
                
            except Exception as e:
                print(f"在文件 {file_path} 中显示列失败: {str(e)}")
                raise
    
    def create_worksheet(self, file_paths, sheet_name):
        """
        创建工作表
        
        Args:
            file_paths: Excel文件路径列表
            sheet_name: 要创建的工作表名称
        """
        for file_path in file_paths:
            try:
                # 确保使用工作簿副本，而不是直接修改原始工作簿
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                # 创建工作表
                wb.create_sheet(title=sheet_name)
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                
            except Exception as e:
                print(f"在文件 {file_path} 中创建工作表失败: {str(e)}")
                raise
    
    def delete_worksheet(self, file_paths, sheet_name):
        """
        删除工作表
        
        Args:
            file_paths: Excel文件路径列表
            sheet_name: 要删除的工作表名称
        """
        for file_path in file_paths:
            try:
                # 确保使用工作簿副本，而不是直接修改原始工作簿
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                # 检查工作表是否存在
                if sheet_name in wb.sheetnames:
                    # 删除工作表
                    del wb[sheet_name]
                    
                    # 保存修改后的工作簿到临时文件
                    temp_path = self.temp_files.get(file_path)
                    if temp_path:
                        wb.save(temp_path)
                else:
                    print(f"工作表 {sheet_name} 在文件 {file_path} 中不存在")
                    continue                
                
            except Exception as e:
                print(f"在文件 {file_path} 中删除工作表失败: {str(e)}")
                raise
    
    def merge_cells(self, file_paths, sheet_indexes, range_str):
        """
        合并单元格，始终保留左上角单元格的值
        
        与Excel默认行为一致，当指定范围内有多个单元格包含值时，
        合并后只保留左上角单元格的值，其他单元格的值将被丢弃。
        
        Args:
            file_paths: Excel文件路径列表
            sheet_indexes: 工作表索引列表，可以是整数索引或工作表名称
            range_str: 要合并的单元格范围，例如'A1:B3'
        
        Raises:
            ValueError: 当工作表索引无效时
            Exception: 当合并操作失败时
        """
        from openpyxl.utils import range_boundaries
        
        for file_path in file_paths:
            try:
                # 确保使用工作簿副本，而不是直接修改原始工作簿
                wb = self.workbooks.get(file_path)
                if not wb:
                    # 创建临时文件副本
                    temp_path = Path(file_path).parent / f"temp_{Path(file_path).name}"
                    shutil.copy2(file_path, temp_path)
                    self.temp_files[file_path] = temp_path
                    
                    # 加载临时文件的工作簿
                    wb = openpyxl.load_workbook(temp_path)
                    self.workbooks[file_path] = wb
                
                for sheet_index in sheet_indexes:
                    # 支持通过索引或名称访问工作表
                    if isinstance(sheet_index, int) and 0 <= sheet_index < len(wb.sheetnames):
                        sheet = wb[wb.sheetnames[sheet_index]]
                    elif isinstance(sheet_index, str) and sheet_index in wb.sheetnames:
                        sheet = wb[sheet_index]
                    else:
                        raise ValueError(f"无效的工作表索引或名称: {sheet_index}")
                        
                    # 获取合并范围的边界
                    min_col, min_row, max_col, max_row = range_boundaries(range_str)
                    
                    # 保存左上角单元格的值
                    top_left_value = sheet.cell(row=min_row, column=min_col).value
                    
                    # 检查范围内是否有其他非空值
                    has_other_values = False
                    non_empty_cells = []
                    
                    for row in range(min_row, max_row + 1):
                        for col in range(min_col, max_col + 1):
                            if row == min_row and col == min_col:
                                continue  # 跳过左上角单元格
                            cell = sheet.cell(row=row, column=col)
                            if cell.value is not None:
                                has_other_values = True
                                cell_addr = f"{get_column_letter(col)}{row}"
                                non_empty_cells.append(cell_addr)
                    
                    # 合并单元格前记录警告信息
                    if has_other_values:
                        cells_str = ", ".join(non_empty_cells[:5])
                        if len(non_empty_cells) > 5:
                            cells_str += f" 等共 {len(non_empty_cells)} 个单元格"
                        print(f"警告: 在文件 {file_path} 的工作表 {sheet.title} 中，"
                              f"合并范围 {range_str} 内存在多个非空值({cells_str})，"
                              f"仅保留左上角单元格 {range_str.split(':')[0]} 的值。")
                    
                    # 合并单元格
                    sheet.merge_cells(range_str)
                    
                    # 确保合并后的单元格保留左上角的值
                    merged_cell = sheet.cell(row=min_row, column=min_col)
                    merged_cell.value = top_left_value
                
                # 保存修改后的工作簿到临时文件
                temp_path = self.temp_files.get(file_path)
                if temp_path:
                    wb.save(temp_path)
                else:
                    # 如果没有临时文件路径，这是一个错误情况
                    raise ValueError(f"找不到文件 {file_path} 的临时文件路径")
                
            except Exception as e:
                print(f"在文件 {file_path} 中合并单元格失败: {str(e)}")
                raise
    
    def close_workbooks(self):
        """
        关闭所有工作簿
        """
        for wb in self.workbooks.values():
            try:
                wb.close()
            except:
                pass
        self.workbooks.clear()