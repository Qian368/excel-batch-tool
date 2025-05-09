#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
行列操作模块
提供Excel批量处理工具的行列操作相关功能
"""

from PyQt5.QtWidgets import QMessageBox

class RowColOperationsMixin:
    """行列操作混入类，提供行列相关的操作方法"""
    
    def validate_input(self, text, is_column_operation):
        """
        验证输入文本的格式
        Args:
            text: 输入的文本
            is_column_operation: 是否为列操作
        Returns:
            (bool, str): (是否有效, 错误信息)
        """
        # 替换中文符号为英文符号
        text = text.replace('，', ',').replace('：', ':')
        
        # 分割多个范围
        ranges = text.split(',')
        for range_str in ranges:
            # 处理单个范围
            if ':' in range_str:
                start, end = range_str.split(':')
                start = start.strip()
                end = end.strip()
                
                if is_column_operation:
                    # 列操作验证：必须是英文字母
                    if not (start.isascii() and start.isalpha() and end.isascii() and end.isalpha()):
                        return False, "列位置只能输入英文字母，例如: A,C:E"
                else:
                    # 行操作验证：必须是数字
                    if not (start.isdigit() and end.isdigit()):
                        return False, "行位置只能输入数字，例如: 1,3:5"
            else:
                # 处理单个位置
                value = range_str.strip()
                if is_column_operation:
                    if not (value.isascii() and value.isalpha()):
                        return False, "列位置只能输入英文字母，例如: A,C:E"
                else:
                    if not value.isdigit():
                        return False, "行位置只能输入数字，例如: 1,3:5"
        
        return True, ""
    
    def get_current_operation(self):
        """获取当前选中的操作类型"""
        if self.insert_rows_radio.isChecked():
            return 'insert_rows'
        elif self.insert_cols_radio.isChecked():
            return 'insert_columns'
        elif self.delete_rows_radio.isChecked():
            return 'delete_rows'
        elif self.delete_cols_radio.isChecked():
            return 'delete_columns'
        elif self.hide_rows_radio.isChecked():
            return 'hide_rows'
        elif self.hide_cols_radio.isChecked():
            return 'hide_columns'
        elif self.unhide_rows_radio.isChecked():
            return 'unhide_rows'
        elif self.unhide_cols_radio.isChecked():
            return 'unhide_columns'
        return None
    
    def set_operation_radio(self, operation):
        """设置操作类型单选按钮"""
        if operation == 'insert_rows':
            self.insert_rows_radio.setChecked(True)
        elif operation == 'insert_columns':
            self.insert_cols_radio.setChecked(True)
        elif operation == 'delete_rows':
            self.delete_rows_radio.setChecked(True)
        elif operation == 'delete_columns':
            self.delete_cols_radio.setChecked(True)
        elif operation == 'hide_rows':
            self.hide_rows_radio.setChecked(True)
        elif operation == 'hide_columns':
            self.hide_cols_radio.setChecked(True)
        elif operation == 'unhide_rows':
            self.unhide_rows_radio.setChecked(True)
        elif operation == 'unhide_columns':
            self.unhide_cols_radio.setChecked(True)
    
    def add_row_col_step(self):
        """添加行列操作步骤"""
        try:
            operation = self.get_current_operation()
            position = self.position_edit.text().strip()
            
            if not position:
                QMessageBox.warning(self, "警告", "请输入位置信息！")
                return
                
            # 验证输入格式
            is_column_operation = operation in [
                'insert_columns', 'delete_columns',
                'hide_columns', 'unhide_columns'
            ]
            is_valid, error_msg = self.validate_input(position, is_column_operation)
            if not is_valid:
                QMessageBox.warning(self, "输入错误", error_msg)
                return
            
            # 替换中文符号为英文符号
            position = position.replace('，', ',').replace('：', ':')
            
            # 添加步骤
            params = {
                'operation': operation,
                'position': position,
                'sheet_indexes': [0]  # 默认处理第一个工作表
            }
            # 如果是删除操作，添加合并单元格处理模式
            if operation in ['delete_rows', 'delete_columns']:
                if self.delete_merge_ignore_radio.isChecked():
                    params['merge_mode'] = 'ignore'
                elif self.delete_merge_unmerge_only_radio.isChecked():
                    params['merge_mode'] = 'unmerge_only'
                elif self.delete_merge_unmerge_keep_value_radio.isChecked():
                    params['merge_mode'] = 'unmerge_keep_value'
                    
            self.add_step(operation, params)
            self.position_edit.clear()
            
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            print(f"添加行列操作步骤失败: {error_msg}")
            QMessageBox.critical(self, "错误", f"添加行列操作步骤失败: {str(e)}")
    
    def insert_row_col_step(self):
        """插入行列操作步骤"""
        try:
            operation = self.get_current_operation()
            position = self.position_edit.text().strip()
            
            if not position:
                QMessageBox.warning(self, "警告", "请输入位置信息！")
                return
                
            # 验证输入格式
            is_column_operation = operation in [
                'insert_columns', 'delete_columns',
                'hide_columns', 'unhide_columns'
            ]
            is_valid, error_msg = self.validate_input(position, is_column_operation)
            if not is_valid:
                QMessageBox.warning(self, "输入错误", error_msg)
                return
            
            # 替换中文符号为英文符号
            position = position.replace('，', ',').replace('：', ':')
            
            # 插入步骤
            params = {
                'operation': operation,
                'position': position,
                'sheet_indexes': [0]  # 默认处理第一个工作表
            }
            # 如果是删除操作，添加合并单元格处理模式
            if operation in ['delete_rows', 'delete_columns']:
                if self.delete_merge_ignore_radio.isChecked():
                    params['merge_mode'] = 'ignore'
                elif self.delete_merge_unmerge_only_radio.isChecked():
                    params['merge_mode'] = 'unmerge_only'
                elif self.delete_merge_unmerge_keep_value_radio.isChecked():
                    params['merge_mode'] = 'unmerge_keep_value'
                    
            self.insert_specific_step(operation, params)
            self.position_edit.clear()
            
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            print(f"插入行列操作步骤失败: {error_msg}")
            QMessageBox.critical(self, "错误", f"插入行列操作步骤失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"添加步骤失败: {str(e)}")
            import traceback
            traceback.print_exc()
            # 防止重复触发
            self.position_edit.returnPressed.disconnect()
            
            operation = self.get_current_operation()
            position = self.position_edit.text().strip()
            
            if not position:
                QMessageBox.warning(self, "警告", "请输入位置！")
                # 重新连接回车事件
                self.position_edit.returnPressed.connect(self.add_row_col_step)
                return
                
            # 获取当前选中的操作类型
            is_column_operation = operation in [
                'insert_columns', 'delete_columns',
                'hide_columns', 'unhide_columns'
            ]
            
            # 验证输入格式
            is_valid, error_msg = self.validate_input(position, is_column_operation)
            if not is_valid:
                QMessageBox.warning(self, "输入错误", error_msg)
                # 重新连接回车事件
                self.position_edit.returnPressed.connect(self.add_row_col_step)
                return
            
            # 替换中文符号为英文符号
            position = position.replace('，', ',').replace('：', ':')
            
            # 安全地添加步骤
            self.safe_add_step_with_validation(operation, {'position': position}, self.position_edit)
            
            # 重新连接回车事件
            self.position_edit.returnPressed.connect(self.add_row_col_step)