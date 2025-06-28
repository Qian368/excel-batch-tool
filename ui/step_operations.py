#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
步骤操作模块
提供Excel批量处理工具的步骤操作相关功能
"""

import json
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from models import StepItem  

class StepOperationsMixin:
    """步骤操作混入类，提供步骤相关的操作方法"""
    
    def add_step(self, operation, params):
        """添加步骤到列表"""
        step = StepItem(operation, params)
        self.steps.append(step)
        self.update_steps_list()
    
    def safe_add_step_with_validation(self, operation, params, input_widget=None):
        """
        安全地添加步骤，包含输入验证
        Args:
            operation: 操作类型
            params: 操作参数
            input_widget: 输入控件（可选）
        """
        try:
            self.add_step(operation, params)
            if input_widget:
                input_widget.clear()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"添加步骤失败：{str(e)}")
    
    def insert_specific_step(self, operation, params):
        """
        在当前选中位置下方插入特定步骤
        Args:
            operation: 操作类型
            params: 操作参数
        """
        try:
            # 获取当前选中的行
            current_row = self.steps_list.currentRow()
            
            # 如果没有选中行，则默认添加到列表末尾
            if current_row < 0:
                self.add_step(operation, params)
                return
            
            # 创建步骤
            step = StepItem(operation, params)
            
            # 将步骤插入到选中位置的下方
            self.steps.insert(current_row + 1, step)
            
            # 更新步骤列表显示
            self.update_steps_list()
            
            # 选中插入的步骤
            self.steps_list.setCurrentRow(current_row + 1)
            
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            print(f"插入特定步骤失败: {error_msg}")
            QMessageBox.critical(self, "错误", f"插入特定步骤失败: {str(e)}")
    
    def get_selected_operation_type(self):
        """获取当前选中的操作类型"""
        # 根据当前选中的选项卡确定操作类型
        current_tab_index = self.tab_widget.currentIndex()
        
        if current_tab_index == 0:  # 公式转值
            return 'convert_formulas_to_values'
        elif current_tab_index == 1:  # 合并单元格处理
            # 检查是否是合并单元格操作
            # 注意：main_window.py中没有merge_radio，需要根据UI结构判断
            if hasattr(self, 'merge_range_edit') and self.merge_range_edit.text().strip():
                return 'merge_cells'
            
            # 检查是哪种拆分合并单元格操作
            if hasattr(self, 'unmerge_mode_group'):
                # 获取选中的模式
                if hasattr(self, 'unmerge_all_radio') and self.unmerge_all_radio.isChecked():
                    return 'process_merged_cells_all'
                elif hasattr(self, 'unmerge_specific_radio') and self.unmerge_specific_radio.isChecked():
                    return 'process_merged_cells_specific'
            
            # 默认处理所有合并单元格
            return 'process_merged_cells_all'
        elif current_tab_index == 2:  # 工作表操作
            # 根据输入框内容判断是创建还是删除工作表
            if hasattr(self, 'create_ws_name_edit') and self.create_ws_name_edit.text().strip():
                return 'create_worksheet'
            elif hasattr(self, 'delete_ws_name_edit') and self.delete_ws_name_edit.text().strip():
                return 'delete_worksheet'
            # 默认返回创建工作表
            return 'create_worksheet'
        elif current_tab_index == 3:  # 行列操作
            # 使用row_col_operations.py中定义的方法获取当前操作类型
            if hasattr(self, 'get_current_operation'):
                return self.get_current_operation()
            # 如果没有该方法，则手动判断
            if hasattr(self, 'insert_rows_radio') and self.insert_rows_radio.isChecked():
                return 'insert_rows'
            elif hasattr(self, 'insert_cols_radio') and self.insert_cols_radio.isChecked():
                return 'insert_columns'
            elif hasattr(self, 'delete_rows_radio') and self.delete_rows_radio.isChecked():
                return 'delete_rows'
            elif hasattr(self, 'delete_cols_radio') and self.delete_cols_radio.isChecked():
                return 'delete_columns'
            elif hasattr(self, 'hide_rows_radio') and self.hide_rows_radio.isChecked():
                return 'hide_rows'
            elif hasattr(self, 'hide_cols_radio') and self.hide_cols_radio.isChecked():
                return 'hide_columns'
            elif hasattr(self, 'unhide_rows_radio') and self.unhide_rows_radio.isChecked():
                return 'unhide_rows'
            elif hasattr(self, 'unhide_cols_radio') and self.unhide_cols_radio.isChecked():
                return 'unhide_columns'
        
        return None
    
    def get_selected_operation_params(self):
        """获取当前选中操作的参数"""
        operation_type = self.get_selected_operation_type()
        if not operation_type:
            return None
            
        params = {}
        
        # 根据操作类型获取相应参数
        if operation_type == 'convert_formulas_to_values':
            # 公式转值不需要额外参数
            pass
        elif operation_type == 'process_merged_cells_all':
            # 处理所有合并单元格
            if hasattr(self, 'unmerge_keep_value_radio'):
                params['action'] = 'keep_value' if self.unmerge_keep_value_radio.isChecked() else 'unmerge'
            # 添加unmerge参数，表示处理所有单元格
            params['unmerge'] = 'all'

        elif operation_type == 'process_merged_cells_specific':
            # 处理指定范围的合并单元格
            if hasattr(self, 'unmerge_keep_value_radio'):
                params['action'] = 'keep_value' if self.unmerge_keep_value_radio.isChecked() else 'unmerge'
            
            # 获取指定的单元格范围
            if hasattr(self, 'unmerge_range_edit'):
                range_str = self.unmerge_range_edit.text().strip()
                if not range_str:
                    QMessageBox.warning(self, "警告", "请输入要拆分的单元格范围！")
                    return None
                
                # 处理中文冒号
                range_str = range_str.replace('：', ':')
                params['range_str'] = range_str
            else:
                QMessageBox.warning(self, "警告", "未找到单元格范围输入框！")
                return None
        elif operation_type == 'merge_cells':
            # 合并单元格
            if hasattr(self, 'merge_range_edit'):
                range_str = self.merge_range_edit.text().strip()
                if not range_str:
                    QMessageBox.warning(self, "警告", "请输入合并范围！")
                    return None
                # 处理中文冒号
                range_str = range_str.replace('：', ':')
                params['range_str'] = range_str
        elif operation_type in ['create_worksheet', 'delete_worksheet']:
            # 工作表操作
            sheet_name = ""
            if operation_type == 'create_worksheet' and hasattr(self, 'create_ws_name_edit'):
                sheet_name = self.create_ws_name_edit.text().strip()
            elif operation_type == 'delete_worksheet' and hasattr(self, 'delete_ws_name_edit'):
                sheet_name = self.delete_ws_name_edit.text().strip()
                
            if not sheet_name:
                QMessageBox.warning(self, "警告", "请输入工作表名称！")
                return None
            params['sheet_name'] = sheet_name
        elif operation_type in ['insert_rows', 'delete_rows', 'insert_columns', 'delete_columns', 
                               'hide_rows', 'hide_columns', 'unhide_rows', 'unhide_columns']:
            # 行列操作
            if hasattr(self, 'position_edit'):
                position = self.position_edit.text().strip()
                if not position:
                    QMessageBox.warning(self, "警告", "请输入位置！")
                    return None
                # 处理中文符号
                position = position.replace('，', ',').replace('：', ':')
                params['position'] = position
                
                # 如果是删除操作，获取合并单元格处理模式
                if operation_type in ['delete_rows', 'delete_columns'] and hasattr(self, 'delete_merge_action_group'):
                    if hasattr(self, 'delete_merge_ignore_radio') and self.delete_merge_ignore_radio.isChecked():
                        params['merge_mode'] = 'ignore'
                    elif hasattr(self, 'delete_merge_unmerge_only_radio') and self.delete_merge_unmerge_only_radio.isChecked():
                        params['merge_mode'] = 'unmerge_only'
                    elif hasattr(self, 'delete_merge_unmerge_keep_value_radio') and self.delete_merge_unmerge_keep_value_radio.isChecked():
                        params['merge_mode'] = 'unmerge_keep_value'
        
        return params
    
    def edit_step(self):
        """
        编辑当前选中的步骤
        - 获取选中步骤的信息
        - 根据步骤类型切换到对应选项卡
        - 填充相应的参数数据
        - 设置焦点到相关输入控件
        """
        try:
            # 获取当前选中的行
            current_row = self.steps_list.currentRow()
            if current_row < 0:
                QMessageBox.warning(self, "警告", "请先选择要编辑的步骤！")
                return
            
            # 获取选中的步骤
            step = self.steps[current_row]
            operation = step.operation
            params = step.params
            
            # 处理嵌套参数结构的情况
            # 如果params中包含operation和params字段，说明是嵌套结构
            if isinstance(params, dict) and 'operation' in params and 'params' in params:
                # 提取真正的操作类型和参数
                real_operation = params['operation']
                real_params = params['params']
                print(f"[DEBUG] 检测到嵌套参数结构，原始操作: {operation}, 实际操作: {real_operation}")
                operation = real_operation
                params = real_params
            
            print(f"[DEBUG] 开始编辑步骤: {operation}, 参数: {params}")
            print(f"[DEBUG] 当前步骤索引: {current_row}, 总步骤数: {len(self.steps)}")  
            
            # 根据操作类型切换到相应的选项卡并填充数据
            if operation == 'convert_formulas_to_values' or operation == '公式转值':
                # 公式转值，切换到第一个选项卡
                self.tab_widget.setCurrentIndex(0)
                
            elif operation in ['process_merged_cells_all', 'process_merged_cells'] or operation.startswith('拆分所有合并单元格'):
                # 处理所有合并单元格，切换到第二个选项卡
                self.tab_widget.setCurrentIndex(1)
                
                # 设置操作类型单选按钮
                if hasattr(self, 'unmerge_all_radio'):
                    self.unmerge_all_radio.setChecked(True)
                
                # 设置合并单元格处理模式
                # 从不同格式的参数中提取action
                action = 'unmerge'  # 默认值
                
                # 直接从params中获取action
                if 'action' in params:
                    action = params['action']
                # 如果是从描述中提取的操作，可能包含在操作名称中
                elif operation.startswith('拆分所有合并单元格'):
                    if '保留值' in operation:
                        action = 'keep_value'
                
                print(f"[DEBUG] 设置拆分模式: {action}")
                
                if hasattr(self, 'unmerge_keep_value_radio') and action == 'keep_value':
                    self.unmerge_keep_value_radio.setChecked(True)
                elif hasattr(self, 'unmerge_only_radio'):
                    self.unmerge_only_radio.setChecked(True)
                    
            elif operation == 'process_merged_cells_specific' or operation.startswith('拆分指定范围'):
                # 处理指定范围的合并单元格，切换到第二个选项卡
                self.tab_widget.setCurrentIndex(1)
                
                # 设置操作类型单选按钮
                if hasattr(self, 'unmerge_specific_radio'):
                    self.unmerge_specific_radio.setChecked(True)
                    # 确保启用范围输入框
                    if hasattr(self, 'unmerge_range_edit'):
                        self.unmerge_range_edit.setEnabled(True)
                
                # 设置合并单元格处理模式
                action = params.get('action', 'unmerge')
                if hasattr(self, 'unmerge_keep_value_radio') and action == 'keep_value':
                    self.unmerge_keep_value_radio.setChecked(True)
                elif hasattr(self, 'unmerge_only_radio'):
                    self.unmerge_only_radio.setChecked(True)
                
                # 填充单元格范围
                # 从不同格式的参数中提取range_str
                range_str = ''
                if 'range_str' in params:
                    range_str = params['range_str']
                # 如果是从描述中提取的操作，可能包含在操作名称中
                elif operation.startswith('拆分指定范围'):
                    # 尝试从操作名称中提取范围，格式如：拆分指定范围 A1:B2 (仅拆分)
                    import re
                    match = re.search(r'拆分指定范围\s+([A-Za-z0-9:]+)', operation)
                    if match:
                        range_str = match.group(1)
                
                if hasattr(self, 'unmerge_range_edit'):
                    self.unmerge_range_edit.setText(range_str)
                    self.unmerge_range_edit.setFocus()
                    
                print(f"[DEBUG] 设置拆分范围: {range_str}")
                    
            elif operation == 'merge_cells' or operation.startswith('合并单元格'):
                # 合并单元格，切换到第二个选项卡
                self.tab_widget.setCurrentIndex(1)
                
                # 填充合并范围
                if hasattr(self, 'merge_range_edit'):
                    range_str = params.get('range_str', '')
                    self.merge_range_edit.setText(range_str)
                    self.merge_range_edit.setFocus()
                    
            elif operation == 'create_worksheet' or operation.startswith('新建工作表'):
                # 新建工作表，切换到第三个选项卡
                self.tab_widget.setCurrentIndex(2)
                # 填充工作表名称
                if 'sheet_name' in params and hasattr(self, 'create_ws_name_edit'):
                    self.create_ws_name_edit.setText(params['sheet_name'])
                    self.create_ws_name_edit.setFocus()
                    
            elif operation == 'delete_worksheet' or operation.startswith('删除工作表'):
                # 删除工作表，切换到第三个选项卡
                self.tab_widget.setCurrentIndex(2)
                # 填充工作表名称
                if 'sheet_name' in params and hasattr(self, 'delete_ws_name_edit'):
                    self.delete_ws_name_edit.setText(params['sheet_name'])
                    self.delete_ws_name_edit.setFocus()
                    
            elif operation in ['insert_rows', 'insert_columns', 'delete_rows', 'delete_columns', 
                              'hide_rows', 'hide_columns', 'unhide_rows', 'unhide_columns']:
                # 行列操作，切换到第四个选项卡
                self.tab_widget.setCurrentIndex(3)
                # 设置操作类型单选按钮
                if hasattr(self, 'set_operation_radio'):
                    self.set_operation_radio(operation)
                # 设置位置
                if 'position' in params and hasattr(self, 'position_edit'):
                    self.position_edit.setText(params['position'])
                    self.position_edit.setFocus()
                
                # 如果是删除操作，设置合并单元格处理模式
                if operation in ['delete_rows', 'delete_columns']:
                    merge_mode = params.get('merge_mode', 'ignore')
                    if merge_mode == 'ignore' and hasattr(self, 'delete_merge_ignore_radio'):
                        self.delete_merge_ignore_radio.setChecked(True)
                    elif merge_mode == 'unmerge_only' and hasattr(self, 'delete_merge_unmerge_only_radio'):
                        self.delete_merge_unmerge_only_radio.setChecked(True)
                    elif merge_mode == 'unmerge_keep_value' and hasattr(self, 'delete_merge_unmerge_keep_value_radio'):
                        self.delete_merge_unmerge_keep_value_radio.setChecked(True)
            elif operation == 'change_font_color' or operation.startswith('修改字体颜色'):
                # 字体颜色操作，切换到字体颜色选项卡
                self.tab_widget.setCurrentIndex(4)  # 假设字体颜色是第5个选项卡
                
                # 设置颜色
                if 'color' in params and hasattr(self, 'font_color_combo'):
                    color_index = self.font_color_combo.findText(params['color'])
                    if color_index >= 0:
                        self.font_color_combo.setCurrentIndex(color_index)
                
                # 设置范围模式
                range_mode = params.get('range_mode', 'specific')
                if range_mode == 'entire_sheet' and hasattr(self, 'font_entire_sheet_radio'):
                    self.font_entire_sheet_radio.setChecked(True)
                elif hasattr(self, 'font_specific_radio'):
                    self.font_specific_radio.setChecked(True)
                
                # 填充单元格范围
                if range_mode == 'specific' and 'range_str' in params and hasattr(self, 'font_range_edit'):
                    self.font_range_edit.setText(params['range_str'])
                    self.font_range_edit.setFocus()
                    
            elif operation == 'change_fill_color' or operation.startswith('修改填充颜色'):
                # 填充颜色操作，切换到填充颜色选项卡
                self.tab_widget.setCurrentIndex(5)  # 假设填充颜色是第6个选项卡
                
                # 设置颜色
                if 'color' in params and hasattr(self, 'fill_color_combo'):
                    color_index = self.fill_color_combo.findText(params['color'])
                    if color_index >= 0:
                        self.fill_color_combo.setCurrentIndex(color_index)
                
                # 设置范围模式
                range_mode = params.get('range_mode', 'specific')
                if range_mode == 'entire_sheet' and hasattr(self, 'fill_entire_sheet_radio'):
                    self.fill_entire_sheet_radio.setChecked(True)
                elif hasattr(self, 'fill_specific_radio'):
                    self.fill_specific_radio.setChecked(True)
                
                # 填充单元格范围
                if range_mode == 'specific' and 'range_str' in params and hasattr(self, 'fill_range_edit'):
                    self.fill_range_edit.setText(params['range_str'])
                    self.fill_range_edit.setFocus()
                    
            elif operation == 'add_border' or operation == 'remove_border' or operation.startswith('添加单元格边框') or operation.startswith('移除单元格边框'):
                # 边框操作，切换到边框选项卡
                self.tab_widget.setCurrentIndex(6)  # 假设边框是第7个选项卡
                
                # 设置边框模式
                if (operation == 'add_border' or operation.startswith('添加单元格边框')) and hasattr(self, 'add_border_radio'):
                    self.add_border_radio.setChecked(True)
                elif hasattr(self, 'remove_border_radio'):
                    self.remove_border_radio.setChecked(True)
                
                # 填充单元格范围
                if 'range_str' in params and hasattr(self, 'border_range_edit'):
                    self.border_range_edit.setText(params['range_str'])
                    self.border_range_edit.setFocus()
                    
            elif operation == 'modify_cell_content' or operation.startswith('修改单元格内容'):
                # 单元格内容修改，切换到单元格内容选项卡
                self.tab_widget.setCurrentIndex(7)  # 假设单元格内容是第8个选项卡
                
                # 填充单元格位置
                if 'position' in params and hasattr(self, 'cell_position_edit'):
                    self.cell_position_edit.setText(params['position'])
                
                # 填充新内容
                if 'content' in params and hasattr(self, 'cell_content_edit'):
                    self.cell_content_edit.setText(params['content'])
                    self.cell_content_edit.setFocus()
                    
            else:
                # 未知操作类型
                QMessageBox.warning(self, "警告", f"未知操作类型: {operation}")
                return
            
            # 记录原始步骤信息用于调试
            original_step_info = f"操作: {step.operation}, 参数: {step.params}"
            
            # 删除当前步骤
            self.steps.pop(current_row)
            
            # 更新步骤列表
            self.update_steps_list()
            
            print(f"[DEBUG] 步骤编辑完成，剩余步骤数: {len(self.steps)}")
            print(f"[DEBUG] 原始步骤信息: {original_step_info}")
            print("")
            
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            print(f"[ERROR] 编辑步骤失败: {error_msg}")
            QMessageBox.critical(self, "错误", f"编辑步骤失败: {str(e)}")
            print("")
    
    def delete_step(self):
        """删除选中的步骤"""
        try:
            # 获取所有选中的行
            selected_items = self.steps_list.selectedItems()
            if not selected_items:
                QMessageBox.information(self, "提示", "请先选择要删除的步骤！")
                return
                
            # 确认删除
            if len(selected_items) > 1:
                confirm = QMessageBox.question(self, "确认删除", 
                                            f"确定要删除选中的 {len(selected_items)} 个步骤吗？",
                                            QMessageBox.Yes | QMessageBox.No)
            else:
                confirm = QMessageBox.question(self, "确认删除", 
                                            "确定要删除选中的步骤吗？",
                                            QMessageBox.Yes | QMessageBox.No)
                
            if confirm != QMessageBox.Yes:
                return
                
            # 获取所有选中行的索引
            selected_rows = [self.steps_list.row(item) for item in selected_items]
            # 按照从大到小的顺序排序，以便从后往前删除，避免索引变化
            selected_rows.sort(reverse=True)
            
            # 从后往前删除选中的步骤
            for row in selected_rows:
                if 0 <= row < len(self.steps):
                    del self.steps[row]
                    
            # 更新步骤列表
            self.update_steps_list()
            
            # 提示删除成功
            if len(selected_items) > 1:
                QMessageBox.information(self, "成功", f"已删除 {len(selected_items)} 个步骤")
            else:
                QMessageBox.information(self, "成功", "已删除选中的步骤")
                
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            print(f"删除步骤失败: {error_msg}")
            QMessageBox.critical(self, "错误", f"删除步骤失败: {str(e)}")

    
    def clear_steps(self):
        """清空步骤列表"""
        self.steps.clear()
        self.update_steps_list()
    
    def move_step_up(self):
        """将选中的步骤向上移动"""
        current_row = self.steps_list.currentRow()
        if current_row > 0:
            self.steps[current_row], self.steps[current_row - 1] = \
                self.steps[current_row - 1], self.steps[current_row]
            self.update_steps_list()
            self.steps_list.setCurrentRow(current_row - 1)
    
    def move_step_down(self):
        """将选中的步骤向下移动"""
        current_row = self.steps_list.currentRow()
        if current_row >= 0 and current_row < len(self.steps) - 1:
            self.steps[current_row], self.steps[current_row + 1] = \
                self.steps[current_row + 1], self.steps[current_row]
            self.update_steps_list()
            self.steps_list.setCurrentRow(current_row + 1)

    def export_steps(self):
        """导出当前步骤列表到文件"""
        if not self.steps:
            QMessageBox.information(self, "提示", "没有步骤可以导出。")
            return

        # 准备要导出的数据
        export_data = []
        for step in self.steps:
            export_data.append({
                'operation': step.operation,
                'params': step.params
            })

        # 打开文件保存对话框（去掉DontUseNativeDialog选项）
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "导出步骤到文件", "", "JSON Files (*.json);;All Files (*)", options=options)

        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(export_data, f, ensure_ascii=False, indent=4)
                QMessageBox.information(self, "成功", f"步骤已成功导出到 {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出步骤失败: {str(e)}")

    def import_steps(self):
        """从文件导入步骤列表"""
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "从文件导入步骤", "", "JSON Files (*.json);;All Files (*)", options=options)

        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    import_data = json.load(f)

                if not isinstance(import_data, list):
                    raise ValueError("导入的文件格式不正确，应为步骤列表。")

                # 验证导入数据的结构
                new_steps = []
                for item in import_data:
                    if not isinstance(item, dict) or 'operation' not in item or 'params' not in item:
                        raise ValueError("导入的数据项格式不正确。")
                    # 可以添加更详细的参数验证
                    new_steps.append(StepItem(item['operation'], item['params']))

                # 清空现有步骤并加载新步骤
                self.steps.clear()
                self.steps.extend(new_steps)
                self.update_steps_list()
                QMessageBox.information(self, "成功", f"步骤已成功从 {file_path} 导入")

            except json.JSONDecodeError:
                QMessageBox.critical(self, "错误", "导入失败：文件不是有效的JSON格式。")
            except ValueError as ve:
                QMessageBox.critical(self, "错误", f"导入失败：{str(ve)}")
            except Exception as e:
                import traceback
                error_msg = traceback.format_exc()
                print(f"导入步骤失败: {error_msg}")
                QMessageBox.critical(self, "错误", f"导入步骤失败: {str(e)}")

    def init_merge_cells_ui(self):
        """
        初始化合并单元格UI组件
        此方法已迁移到main_window.py中实现
        """
        pass
        
        # # 指定范围输入框
        # unmerge_range_layout = QHBoxLayout()
        # unmerge_range_layout.addWidget(QLabel("单元格范围:"))
        # self.unmerge_range_edit = QLineEdit()
        # self.unmerge_range_edit.setPlaceholderText("例如: A1:B50 或 A1：B50 或单个单元格 B50")
        # unmerge_range_layout.addWidget(self.unmerge_range_edit)
        
        # # 指定范围的保留值选项
        # unmerge_specific_value_layout = QHBoxLayout()
        # self.unmerge_specific_keep_value_radio = QRadioButton("保留值")
        # self.unmerge_specific_only_radio = QRadioButton("仅拆分")
        # self.unmerge_specific_value_group = QButtonGroup(self)
        # self.unmerge_specific_value_group.addButton(self.unmerge_specific_keep_value_radio)
        # self.unmerge_specific_value_group.addButton(self.unmerge_specific_only_radio)
        # self.unmerge_specific_only_radio.setChecked(True)  # 默认选中仅拆分
        # unmerge_specific_value_layout.addWidget(self.unmerge_specific_keep_value_radio)
        # unmerge_specific_value_layout.addWidget(self.unmerge_specific_only_radio)
        
        # # 组装指定范围布局
        # unmerge_specific_layout.addLayout(unmerge_range_layout)
        # unmerge_specific_layout.addLayout(unmerge_specific_value_layout)
        
        # # 将三种模式添加到模式组中
        # self.unmerge_mode_group_btns.addButton(self.unmerge_all_radio)
        # self.unmerge_mode_group_btns.addButton(self.unmerge_row_col_radio)
        # self.unmerge_mode_group_btns.addButton(self.unmerge_specific_radio)
        # self.unmerge_all_radio.setChecked(True)  # 默认选中处理整个工作表
        
        # # 组装模式布局
        # unmerge_mode_layout.addWidget(self.unmerge_all_radio)
        # unmerge_mode_layout.addLayout(unmerge_all_layout)
        # unmerge_mode_layout.addWidget(self.unmerge_row_col_radio)
        # unmerge_mode_layout.addLayout(unmerge_row_col_layout)
        # unmerge_mode_layout.addWidget(self.unmerge_specific_radio)
        # unmerge_mode_layout.addLayout(unmerge_specific_layout)
        # self.unmerge_mode_group.setLayout(unmerge_mode_layout)
        
        # # 将操作类型添加到操作类型组中
        # self.merge_type_group.addButton(self.merge_radio)
        # merge_type_layout.addWidget(self.merge_radio)
        # merge_type_layout.addLayout(merge_range_layout)
        # merge_type_layout.addWidget(self.unmerge_mode_group)
        # merge_type_group.setLayout(merge_type_layout)
        
        # # 添加到合并单元格选项卡
        # self.merge_cells_tab_layout.addWidget(merge_type_group)
        
        # # 连接信号槽
        # self.merge_radio.toggled.connect(self.on_merge_type_changed)
        # self.unmerge_all_radio.toggled.connect(self.on_unmerge_mode_changed)
        # self.unmerge_row_col_radio.toggled.connect(self.on_unmerge_mode_changed)
        # self.unmerge_specific_radio.toggled.connect(self.on_unmerge_mode_changed)
    
    def on_merge_type_changed(self, checked):
        """合并单元格操作类型变更处理"""
        if checked:
            # 合并单元格被选中，启用合并范围输入框，禁用拆分模式组
            self.merge_range_edit.setEnabled(True)
            self.unmerge_mode_group.setEnabled(False)
        else:
            # 拆分合并单元格被选中，禁用合并范围输入框，启用拆分模式组
            self.merge_range_edit.setEnabled(False)
            self.unmerge_mode_group.setEnabled(True)
    
    # def on_unmerge_mode_changed(self, checked):
    #     """拆分合并单元格模式变更处理"""
    #     if not checked:
    #         return
            
    #     # 根据选中的模式启用/禁用相应的控件
    #     self.unmerge_keep_value_radio.setEnabled(self.unmerge_all_radio.isChecked())
    #     self.unmerge_only_radio.setEnabled(self.unmerge_all_radio.isChecked())
        
    #     self.unmerge_row_col_keep_value_radio.setEnabled(self.unmerge_row_col_radio.isChecked())
    #     self.unmerge_row_col_only_radio.setEnabled(self.unmerge_row_col_radio.isChecked())
        
    #     self.unmerge_range_edit.setEnabled(self.unmerge_specific_radio.isChecked())
    #     self.unmerge_specific_keep_value_radio.setEnabled(self.unmerge_specific_radio.isChecked())
    #     self.unmerge_specific_only_radio.setEnabled(self.unmerge_specific_radio.isChecked())
   
    def update_steps_list(self):
        """更新步骤列表显示，并添加步骤编号"""
        self.steps_list.clear()
        for i, step in enumerate(self.steps, 1):
            # 添加步骤编号，格式：[编号] 步骤描述
            self.steps_list.addItem(f"[{i}] {str(step)}")