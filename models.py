#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
模型类模块
提供步骤项等数据模型的定义
"""


class StepItem:
    """步骤项，用于记录操作步骤"""
    # 操作类型映射到中文描述
    operation_desc = {
            'convert_formulas_to_values': '公式转值',
            'process_merged_cells_all': '拆分所有合并单元格',
            'process_merged_cells_specific': '拆分指定范围',
            # 'process_merged_cells': '拆分合并单元格',  # 新增：支持指定范围拆分
            'merge_cells': '合并单元格',
            'create_worksheet': '新建工作表',
            'delete_worksheet': '删除工作表',
            'insert_rows': '插入行',
            'delete_rows': '删除行',
            'hide_rows': '隐藏行',
            'unhide_rows': '取消隐藏行',
            'insert_columns': '插入列',
            'delete_columns': '删除列',
            'hide_columns': '隐藏列',
            'unhide_columns': '取消隐藏列'
        }
        
    def __init__(self, operation, params):
        self.operation = operation
        self.params = params

    def __str__(self):
        # 获取操作的中文描述
        desc = self.operation_desc.get(self.operation, self.operation)
        
        # 处理参数的中文描述
        params_desc = []
        if self.params:
            if 'action' in self.params:
                action_desc = '保留值' if self.params['action'] == 'keep_value' else '仅拆分'
                params_desc.append(f'方式：{action_desc}')
            if 'mode' in self.params:
                mode_desc = {
                    'all': '所有单元格',
                    'row_col': '行列操作时',
                    'specific': '指定范围'
                }.get(self.params['mode'], self.params['mode'])
                params_desc.append(f'模式：{mode_desc}')
            if 'range_str' in self.params:
                # 将中文符号转换为英文符号
                range_str = self.params["range_str"].replace('，', ',').replace('：', ':')
                params_desc.append(f'单元格：{range_str}')
            if 'sheet_name' in self.params:
                params_desc.append(f'工作表：{self.params["sheet_name"]}')
            if 'position' in self.params:
                # 将中文符号转换为英文符号
                position = self.params["position"].replace('，', ',').replace('：', ':')
                params_desc.append(f'位置：{position}')
            # 添加对删除行列操作中合并单元格处理模式的描述
            if self.operation in ['delete_rows', 'delete_columns'] and 'merge_mode' in self.params:
                merge_mode_desc = {
                    'ignore': '不处理',
                    'unmerge_only': '仅拆分',
                    'unmerge_keep_value': '拆分并保留值'
                }.get(self.params['merge_mode'], '未知') # 添加默认值以防万一
                params_desc.append(f'合并处理：{merge_mode_desc}')
        
        # 组合最终描述
        if params_desc:
            return f"{desc}（{', '.join(params_desc)}）"
        return desc