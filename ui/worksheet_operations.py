#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
工作表操作模块
提供Excel批量处理工具的工作表操作相关功能
"""

from PyQt5.QtWidgets import QMessageBox

class WorksheetOperationsMixin:
    """工作表操作相关的功能"""
    
    def add_merge_step(self):
        """添加合并单元格步骤"""
        try:
            range_str = self.merge_range_edit.text().strip()
            
            if not range_str:
                QMessageBox.warning(self, "输入错误", "请输入合并范围！")
                return
                
            # 替换中文冒号为英文冒号
            range_str = range_str.replace('：', ':')
            
            # 验证输入格式
            if not self.validate_cell_range(range_str):
                QMessageBox.warning(self, "输入错误", 
                    "请输入有效的单元格范围，格式如：A1:B50 或 A1：B50\n"
                    "示例：A1:D5 或 B10：E20")
                return
                
            self.add_step(f"合并单元格({range_str})", {'range_str': range_str})
            self.merge_range_edit.clear()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"添加步骤失败: {str(e)}")
            import traceback
            traceback.print_exc()

    def insert_merge_step(self):
        """插入合并单元格步骤，增加输入校验"""
        try:
            range_str = self.merge_range_edit.text().strip()
            if not range_str:
                QMessageBox.warning(self, "输入错误", "请输入合并范围！")
                return
            # 替换中文冒号为英文冒号
            range_str = range_str.replace('：', ':')
            # 验证输入格式
            if not self.validate_cell_range(range_str):
                QMessageBox.warning(self, "输入错误", 
                    "请输入有效的单元格范围，格式如：A1:B50 或 A1：B50\n"
                    "示例：A1:D5 或 B10：E20")
                return
            self.insert_specific_step(f"合并单元格({range_str})", {'range_str': range_str})
            self.merge_range_edit.clear()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"插入步骤失败: {str(e)}")
            import traceback
            traceback.print_exc()

    def add_create_worksheet_step(self):
        """添加新建工作表步骤"""
        sheet_name = self.create_ws_name_edit.text().strip()
        if not sheet_name:
            QMessageBox.warning(self, "警告", "请输入工作表名称！")
            return
        self.add_step(f"新建工作表({sheet_name})", {'sheet_name': sheet_name})
        self.create_ws_name_edit.clear()

    def insert_create_worksheet_step(self):
        """插入新建工作表步骤"""
        sheet_name = self.create_ws_name_edit.text().strip()
        if not sheet_name:
            QMessageBox.warning(self, "警告", "请输入工作表名称！")
            return
        self.insert_specific_step(f"新建工作表({sheet_name})", {'sheet_name': sheet_name})
        self.create_ws_name_edit.clear()

    def add_delete_worksheet_step(self):
        """添加删除工作表步骤"""
        sheet_name = self.delete_ws_name_edit.text().strip()
        if not sheet_name:
            QMessageBox.warning(self, "警告", "请输入工作表名称！")
            return
        self.add_step(f"删除工作表({sheet_name})", {'sheet_name': sheet_name})
        self.delete_ws_name_edit.clear()

    def insert_delete_worksheet_step(self):
        """插入删除工作表步骤"""
        sheet_name = self.delete_ws_name_edit.text().strip()
        if not sheet_name:
            QMessageBox.warning(self, "警告", "请输入工作表名称！")
            return
        self.insert_specific_step(f"删除工作表({sheet_name})", {'sheet_name': sheet_name})
        self.delete_ws_name_edit.clear()
        
    def set_worksheet_operation(self, operation, params):
        """设置工作表操作相关控件状态"""
        if operation == 'create_worksheet':
            # 设置为新建工作表
            if hasattr(self, 'create_ws_radio'):
                self.create_ws_radio.setChecked(True)
            if hasattr(self, 'create_ws_name_edit') and 'sheet_name' in params:
                self.create_ws_name_edit.setText(params['sheet_name'])
        elif operation == 'delete_worksheet':
            # 设置为删除工作表
            if hasattr(self, 'delete_ws_radio'):
                self.delete_ws_radio.setChecked(True)
            if hasattr(self, 'delete_ws_name_edit') and 'sheet_name' in params:
                self.delete_ws_name_edit.setText(params['sheet_name'])

    def validate_cell_range(self, range_str):
        """验证单元格范围格式，支持中文冒号"""
        try:
            import re
            # 替换中文冒号为英文冒号
            range_str = range_str.replace('：', ':')
            # 验证格式如A1:B50
            pattern = r'^[A-Za-z]+\d+:[A-Za-z]+\d+$'
            if not re.match(pattern, range_str):
                return False
            # 验证行列号是否有效
            start, end = range_str.split(':')
            return self.is_valid_cell(start) and self.is_valid_cell(end)
        except Exception:
            return False

    def is_valid_cell(self, cell_ref):
        """验证单个单元格引用是否有效"""
        import re
        return re.match(r'^[A-Za-z]+\d+$', cell_ref) is not None
        
    # def add_unmerge_step(self):
    #     """添加拆分合并单元格步骤"""
    #     try:
    #         # 获取并验证单元格范围或单个单元格
    #         range_str = self.unmerge_range_edit.text().strip()
    #         if not range_str:
    #             QMessageBox.warning(self, "警告", "请输入要拆分的单元格范围或单个单元格！")
    #             return
                
    #         # 处理中文冒号
    #         range_str = range_str.replace('：', ':')
            
    #         # 验证输入是有效单元格范围还是单个单元格
    #         if not (self.validate_cell_range(range_str) or self.is_valid_cell(range_str)):
    #             QMessageBox.warning(self, "输入错误", 
    #                 "请输入有效的单元格范围（如 A1:B50）或单个单元格（如 C3）！")
    #             return
                
    #         # 设置操作参数
    #         operation = 'process_merged_cells'
    #         params = {
    #             'action': 'keep_value' if self.unmerge_keep_value_radio.isChecked() else 'unmerge',
    #             'range_str': range_str
    #         }
    #         desc = f"拆分指定范围 {range_str} ({'保留值' if params['action'] == 'keep_value' else '仅拆分'})"
            
    #         # 添加步骤
    #         self.add_step(desc, {'operation': operation, 'params': params})
    #         self.unmerge_range_edit.clear()
    #         # QMessageBox.information(self, "成功", "已添加拆分合并单元格步骤")
                
    #     except Exception as e:
    #         import traceback
    #         error_msg = traceback.format_exc()
    #         print(f"添加拆分合并单元格步骤失败: {error_msg}")
    #         QMessageBox.critical(self, "错误", f"添加拆分合并单元格步骤失败: {str(e)}")
            
    # def insert_unmerge_step(self):
    #     """插入拆分合并单元格步骤"""
    #     try:
    #         # 获取并验证单元格范围或单个单元格
    #         range_str = self.unmerge_range_edit.text().strip()
    #         if not range_str:
    #             QMessageBox.warning(self, "警告", "请输入要拆分的单元格范围或单个单元格！")
    #             return
                
    #         # 处理中文冒号
    #         range_str = range_str.replace('：', ':')
            
    #         # 验证输入是有效单元格范围还是单个单元格
    #         if not (self.validate_cell_range(range_str) or self.is_valid_cell(range_str)):
    #             QMessageBox.warning(self, "输入错误", 
    #                 "请输入有效的单元格范围（如 A1:B50）或单个单元格（如 C3）！")
    #             return
                
    #         # 设置操作参数
    #         operation = 'process_merged_cells'
    #         params = {
    #             'action': 'keep_value' if self.unmerge_keep_value_radio.isChecked() else 'unmerge',
    #             'range_str': range_str
    #         }
    #         desc = f"拆分指定范围 {range_str} ({'保留值' if params['action'] == 'keep_value' else '仅拆分'})"
            
    #         # 插入步骤
    #         self.insert_specific_step(desc, {'operation': operation, 'params': params})
    #         self.unmerge_range_edit.clear()
    #         # QMessageBox.information(self, "成功", "已插入拆分合并单元格步骤")
                
    #     except Exception as e:
    #         import traceback
    #         error_msg = traceback.format_exc()
    #         print(f"插入拆分合并单元格步骤失败: {error_msg}")
    #         QMessageBox.critical(self, "错误", f"插入拆分合并单元格步骤失败: {str(e)}")
