#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
主窗口模块
整合所有UI功能模块，提供完整的Excel批量处理工具界面
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, 
    QPushButton, QLabel, QButtonGroup, QRadioButton,
    QLineEdit, QGridLayout, QMessageBox  # 添加QMessageBox
)
from PyQt5.QtCore import Qt, QEvent

from ui.base_window import BaseWindow
from ui.file_operations import FileOperationsMixin
from ui.step_operations import StepOperationsMixin
from ui.worksheet_operations import WorksheetOperationsMixin
from ui.row_col_operations import RowColOperationsMixin
from ui.cell_format_operations import CellFormatOperationsMixin

class MainWindow(BaseWindow, FileOperationsMixin, StepOperationsMixin,
                WorksheetOperationsMixin, RowColOperationsMixin, CellFormatOperationsMixin):
    """主窗口类，整合所有功能模块"""
    
    def __init__(self):
        super().__init__()
        self.setup_connections()
        self.setup_tabs()
        # 安装事件过滤器
        self.position_edit.installEventFilter(self)
        self.unmerge_range_edit.installEventFilter(self)
        self.merge_range_edit.installEventFilter(self)
        self.create_ws_name_edit.installEventFilter(self)
        self.delete_ws_name_edit.installEventFilter(self)
        # 为单元格格式操作相关控件安装事件过滤器
        if hasattr(self, 'font_range_edit'):
            self.font_range_edit.installEventFilter(self)
        if hasattr(self, 'fill_range_edit'):
            self.fill_range_edit.installEventFilter(self)
        if hasattr(self, 'border_range_edit'):
            self.border_range_edit.installEventFilter(self)
        if hasattr(self, 'cell_position_edit'):
            self.cell_position_edit.installEventFilter(self)
    
    def setup_connections(self):
        """设置信号连接"""
        # 文件操作按钮连接
        self.add_files_btn.clicked.connect(self.add_files)
        self.add_folder_btn.clicked.connect(self.add_folder)
        self.remove_files_btn.clicked.connect(self.remove_selected_files)
        self.clear_files_btn.clicked.connect(self.clear_files)
        
        # 步骤操作按钮连接
        self.move_up_btn.clicked.connect(self.move_step_up)
        self.move_down_btn.clicked.connect(self.move_step_down)
        # 移除"插入步骤"按钮的信号连接
        # self.insert_step_btn.clicked.connect(self.insert_step)
        self.edit_step_btn.clicked.connect(self.edit_step)  # 确保这行代码存在
        self.delete_step_btn.clicked.connect(self.delete_step)
        self.clear_steps_btn.clicked.connect(self.clear_steps)
        self.export_steps_btn.clicked.connect(self.export_steps) # 连接导出按钮信号
        self.import_steps_btn.clicked.connect(self.import_steps) # 连接导入按钮信号
    
    def setup_tabs(self):
        """设置功能选项卡"""
        # 公式转值选项卡
        self.setup_formula_tab()
        
        # 合并单元格处理选项卡
        self.setup_merge_tab()
        
        # 工作表管理选项卡
        self.setup_worksheet_tab()
        
        # 行列操作选项卡
        self.setup_row_col_tab()
        
        # 单元格格式操作选项卡
        self.setup_cell_format_tabs()

        # 连接删除行列单选按钮的信号，用于控制合并单元格处理选项的可见性
        self.delete_rows_radio.toggled.connect(self.toggle_delete_merge_options)
        self.delete_cols_radio.toggled.connect(self.toggle_delete_merge_options)
        self.delete_hidden_rows_radio.toggled.connect(self.toggle_delete_merge_options)
        self.delete_hidden_cols_radio.toggled.connect(self.toggle_delete_merge_options)
        # 连接其他行列操作单选按钮的信号，用于隐藏合并单元格处理选项
        self.insert_rows_radio.toggled.connect(self.toggle_delete_merge_options)
        self.insert_cols_radio.toggled.connect(self.toggle_delete_merge_options)
        self.hide_rows_radio.toggled.connect(self.toggle_delete_merge_options)
        self.hide_cols_radio.toggled.connect(self.toggle_delete_merge_options)
        self.unhide_rows_radio.toggled.connect(self.toggle_delete_merge_options)
        self.unhide_cols_radio.toggled.connect(self.toggle_delete_merge_options)
        
        # 连接所有行列操作单选按钮的信号，用于控制位置输入框的启用/禁用状态
        self.insert_rows_radio.toggled.connect(self.toggle_position_input)
        self.insert_cols_radio.toggled.connect(self.toggle_position_input)
        self.delete_rows_radio.toggled.connect(self.toggle_position_input)
        self.delete_cols_radio.toggled.connect(self.toggle_position_input)
        self.delete_hidden_rows_radio.toggled.connect(self.toggle_position_input)
        self.delete_hidden_cols_radio.toggled.connect(self.toggle_position_input)
        self.hide_rows_radio.toggled.connect(self.toggle_position_input)
        self.hide_cols_radio.toggled.connect(self.toggle_position_input)
        self.unhide_rows_radio.toggled.connect(self.toggle_position_input)
        self.unhide_cols_radio.toggled.connect(self.toggle_position_input)
    
    def setup_formula_tab(self):
        """设置公式转值选项卡"""
        formula_tab = QWidget()
        formula_layout = QVBoxLayout(formula_tab)
        
        formula_desc = QLabel("将Excel中的公式转换为实际值，保留格式。")
        formula_desc.setWordWrap(True)
        formula_layout.addWidget(formula_desc)
        
        add_formula_btn = QPushButton("添加到步骤")
        add_formula_btn.clicked.connect(lambda: self.add_step('convert_formulas_to_values', {}))
        formula_layout.addWidget(add_formula_btn)
        
        # 添加"插入到当前步骤下方"按钮
        insert_formula_btn = QPushButton("插入到当前步骤下方")
        insert_formula_btn.clicked.connect(lambda: self.insert_specific_step('convert_formulas_to_values', {}))
        formula_layout.addWidget(insert_formula_btn)
        
        formula_layout.addStretch()
        
        self.tab_widget.addTab(formula_tab, "公式转值")
    
    def setup_merge_tab(self):
        """设置合并单元格处理选项卡"""
        merge_tab = QWidget()
        merge_layout = QVBoxLayout(merge_tab)
        
        # 合并单元格的处理方式
        unmerge_group = QGroupBox("拆分合并单元格（合并单元格与指定范围有交集即拆分）")
        unmerge_layout = QVBoxLayout()
        
        # 拆分模式选择
        self.unmerge_mode_group = QButtonGroup()
        self.unmerge_all_radio = QRadioButton("拆分所有合并单元格")
        self.unmerge_all_radio.setChecked(True)
        unmerge_layout.addWidget(self.unmerge_all_radio)
        
        # 处理指定范围和输入框放在同一行
        specific_range_layout = QHBoxLayout()
        self.unmerge_specific_radio = QRadioButton("拆分指定范围")
        specific_range_layout.addWidget(self.unmerge_specific_radio)
        
        # 范围输入框
        self.unmerge_range_edit = QLineEdit()
        self.unmerge_range_edit.setPlaceholderText("例如: A1:B50 或 A1：B50 或单个单元格 B50 (支持中文冒号)")
        specific_range_layout.addWidget(self.unmerge_range_edit)
        
        self.unmerge_mode_group.addButton(self.unmerge_all_radio)
        self.unmerge_mode_group.addButton(self.unmerge_specific_radio)
        unmerge_layout.addLayout(specific_range_layout)
        
        # 拆分处理方式选择
        action_layout = QHBoxLayout()
        self.unmerge_action_group = QButtonGroup()
        self.unmerge_only_radio = QRadioButton("仅拆分")
        self.unmerge_only_radio.setChecked(True)
        self.unmerge_keep_value_radio = QRadioButton("拆分并保留值")
        self.unmerge_action_group.addButton(self.unmerge_only_radio)
        self.unmerge_action_group.addButton(self.unmerge_keep_value_radio)
        action_layout.addWidget(self.unmerge_only_radio)
        action_layout.addWidget(self.unmerge_keep_value_radio)
        unmerge_layout.addLayout(action_layout)
        
        # 添加按钮
        add_unmerge_btn = QPushButton("添加到步骤")
        add_unmerge_btn.clicked.connect(self.add_unmerge_step)
        unmerge_layout.addWidget(add_unmerge_btn)
        
        # 添加"插入到当前步骤下方"按钮
        insert_unmerge_btn = QPushButton("插入到当前步骤下方")
        insert_unmerge_btn.clicked.connect(self.insert_unmerge_step)
        unmerge_layout.addWidget(insert_unmerge_btn)
        
        unmerge_group.setLayout(unmerge_layout)
        
        # 合并单元格组
        merge_group = QGroupBox("合并单元格（默认保留左上角单元格的值）")
        merge_inner_layout = QVBoxLayout()
        
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel("合并范围:"))
        self.merge_range_edit = QLineEdit()
        self.merge_range_edit.setPlaceholderText("例如: A1:B50 或 A1：B50 (支持中文冒号)")
        range_layout.addWidget(self.merge_range_edit)
        
        add_merge_btn = QPushButton("添加到步骤")
        add_merge_btn.clicked.connect(self.add_merge_step)
        
        # 添加"插入到当前步骤下方"按钮
        insert_merge_btn = QPushButton("插入到当前步骤下方")
        insert_merge_btn.clicked.connect(self.insert_merge_step)
        
        merge_inner_layout.addLayout(range_layout)
        merge_inner_layout.addWidget(add_merge_btn)
        merge_inner_layout.addWidget(insert_merge_btn)
        merge_group.setLayout(merge_inner_layout)
        
        merge_layout.addWidget(unmerge_group)
        merge_layout.addWidget(merge_group)
        merge_layout.addStretch()
        
        self.tab_widget.addTab(merge_tab, "合并单元格处理")
        
        # 连接单选按钮信号以控制范围输入框的可见性
        self.unmerge_specific_radio.toggled.connect(
            lambda checked: self.unmerge_range_edit.setEnabled(checked))
        self.unmerge_range_edit.setEnabled(False)  # 初始状态下禁用
    
    def setup_worksheet_tab(self):
        """设置工作表管理选项卡"""
        worksheet_tab = QWidget()
        worksheet_layout = QVBoxLayout(worksheet_tab)
        
        # 新建工作表组
        create_ws_group = QGroupBox("新建工作表")
        create_ws_layout = QVBoxLayout()
        
        create_ws_name_layout = QHBoxLayout()
        create_ws_name_layout.addWidget(QLabel("名称:"))
        self.create_ws_name_edit = QLineEdit()
        self.create_ws_name_edit.setPlaceholderText("输入新工作表名称")
        create_ws_name_layout.addWidget(self.create_ws_name_edit)
        
        add_create_ws_btn = QPushButton("添加到步骤")
        add_create_ws_btn.clicked.connect(self.add_create_worksheet_step)
        
        # 添加"插入到当前步骤下方"按钮
        insert_create_ws_btn = QPushButton("插入到当前步骤下方")
        insert_create_ws_btn.clicked.connect(self.insert_create_worksheet_step)
        
        create_ws_layout.addLayout(create_ws_name_layout)
        create_ws_layout.addWidget(add_create_ws_btn)
        create_ws_layout.addWidget(insert_create_ws_btn)
        create_ws_group.setLayout(create_ws_layout)
        
        # 删除工作表组
        delete_ws_group = QGroupBox("删除工作表")
        delete_ws_layout = QVBoxLayout()
        
        delete_ws_name_layout = QHBoxLayout()
        delete_ws_name_layout.addWidget(QLabel("名称:"))
        self.delete_ws_name_edit = QLineEdit()
        self.delete_ws_name_edit.setPlaceholderText("输入要删除的工作表名称")
        delete_ws_name_layout.addWidget(self.delete_ws_name_edit)
        
        add_delete_ws_btn = QPushButton("添加到步骤")
        add_delete_ws_btn.clicked.connect(self.add_delete_worksheet_step)
        
        # 添加"插入到当前步骤下方"按钮
        insert_delete_ws_btn = QPushButton("插入到当前步骤下方")
        insert_delete_ws_btn.clicked.connect(self.insert_delete_worksheet_step)
        
        delete_ws_layout.addLayout(delete_ws_name_layout)
        delete_ws_layout.addWidget(add_delete_ws_btn)
        delete_ws_layout.addWidget(insert_delete_ws_btn)
        delete_ws_group.setLayout(delete_ws_layout)
        
        worksheet_layout.addWidget(create_ws_group)
        worksheet_layout.addWidget(delete_ws_group)
        worksheet_layout.addStretch()
        
        self.tab_widget.addTab(worksheet_tab, "工作表管理")
    
    def setup_row_col_tab(self):
        """设置行列操作选项卡"""
        row_col_tab = QWidget()
        row_col_layout = QVBoxLayout(row_col_tab)
        
        # 操作类型选择组
        operation_group = QGroupBox("操作类型（注意插入和删除行列会影响后续步骤的行列编号）")
        operation_layout = QGridLayout()
        
        self.operation_group = QButtonGroup()
        self.insert_rows_radio = QRadioButton("插入行（建议靠后处理）")
        self.insert_rows_radio.setChecked(True)
        self.insert_cols_radio = QRadioButton("插入列（建议靠后处理）")
        self.delete_rows_radio = QRadioButton("删除行（建议靠后处理）")
        self.delete_cols_radio = QRadioButton("删除列（建议靠后处理）")
        self.delete_hidden_rows_radio = QRadioButton("删除所有隐藏行（建议靠后处理）")
        self.delete_hidden_cols_radio = QRadioButton("删除所有隐藏列（建议靠后处理）")
        self.hide_rows_radio = QRadioButton("隐藏行")
        self.hide_cols_radio = QRadioButton("隐藏列")
        self.unhide_rows_radio = QRadioButton("取消隐藏行")
        self.unhide_cols_radio = QRadioButton("取消隐藏列")
        
        # 添加单选按钮到按钮组
        self.operation_group.addButton(self.insert_rows_radio)
        self.operation_group.addButton(self.insert_cols_radio)
        self.operation_group.addButton(self.delete_rows_radio)
        self.operation_group.addButton(self.delete_cols_radio)
        self.operation_group.addButton(self.delete_hidden_rows_radio)
        self.operation_group.addButton(self.delete_hidden_cols_radio)
        self.operation_group.addButton(self.hide_rows_radio)
        self.operation_group.addButton(self.hide_cols_radio)
        self.operation_group.addButton(self.unhide_rows_radio)
        self.operation_group.addButton(self.unhide_cols_radio)
        
        # 布局单选按钮
        operation_layout.addWidget(self.insert_rows_radio, 0, 0)
        operation_layout.addWidget(self.insert_cols_radio, 0, 1)
        operation_layout.addWidget(self.delete_rows_radio, 1, 0)
        operation_layout.addWidget(self.delete_cols_radio, 1, 1)
        operation_layout.addWidget(self.delete_hidden_rows_radio, 2, 0)
        operation_layout.addWidget(self.delete_hidden_cols_radio, 2, 1)
        operation_layout.addWidget(self.hide_rows_radio, 3, 0)
        operation_layout.addWidget(self.hide_cols_radio, 3, 1)
        operation_layout.addWidget(self.unhide_rows_radio, 4, 0)
        operation_layout.addWidget(self.unhide_cols_radio, 4, 1)

        
        operation_group.setLayout(operation_layout)
        
        # 位置输入组
        position_group = QGroupBox("位置")
        position_layout = QVBoxLayout()
        
        position_desc = QLabel("输入行号或列字母，多个位置用逗号分隔，范围用冒号表示。\n例如：行：1,3,5:7 列：A,C,E:G")
        position_desc.setWordWrap(True)
        
        position_input_layout = QHBoxLayout()
        position_input_layout.addWidget(QLabel("位置:"))
        self.position_edit = QLineEdit()
        position_input_layout.addWidget(self.position_edit)
        
        # 添加回车事件处理
        self.position_edit.returnPressed.connect(self.add_row_col_step)
        self.position_edit.installEventFilter(self)
        
        position_layout.addWidget(position_desc)
        position_layout.addLayout(position_input_layout)
        position_group.setLayout(position_layout)
        
        add_row_col_btn = QPushButton("添加到步骤")
        add_row_col_btn.clicked.connect(self.add_row_col_step)
        
        # 添加"插入到当前步骤下方"按钮
        insert_row_col_btn = QPushButton("插入到当前步骤下方")
        insert_row_col_btn.clicked.connect(self.insert_row_col_step)
        
        position_layout.addLayout(position_input_layout)
        position_layout.addWidget(add_row_col_btn)
        position_layout.addWidget(insert_row_col_btn)
        position_group.setLayout(position_layout)
        
        row_col_layout.addWidget(operation_group)
        row_col_layout.addWidget(position_group)

        # --- 新增：删除时合并单元格处理 --- 
        self.delete_merge_group = QGroupBox("删除时合并单元格处理")
        delete_merge_layout = QVBoxLayout()
        self.delete_merge_group.setLayout(delete_merge_layout)

        self.delete_merge_action_group = QButtonGroup(self) # 确保按钮组有父对象

        self.delete_merge_ignore_radio = QRadioButton("不处理 (直接删除)")
        self.delete_merge_ignore_radio.setChecked(True) # 默认选中不处理
        self.delete_merge_action_group.addButton(self.delete_merge_ignore_radio)
        delete_merge_layout.addWidget(self.delete_merge_ignore_radio)

        self.delete_merge_unmerge_only_radio = QRadioButton("仅拆分")
        self.delete_merge_action_group.addButton(self.delete_merge_unmerge_only_radio)
        delete_merge_layout.addWidget(self.delete_merge_unmerge_only_radio)

        self.delete_merge_unmerge_keep_value_radio = QRadioButton("拆分并保留值")
        self.delete_merge_action_group.addButton(self.delete_merge_unmerge_keep_value_radio)
        delete_merge_layout.addWidget(self.delete_merge_unmerge_keep_value_radio)

        row_col_layout.addWidget(self.delete_merge_group)
        self.delete_merge_group.setVisible(False) # 初始隐藏
        # --- 结束：删除时合并单元格处理 ---

        row_col_layout.addStretch()
        
        self.tab_widget.addTab(row_col_tab, "行列操作")

    def toggle_delete_merge_options(self, checked):
        """切换删除时合并单元格处理选项的可见性"""
        # 只有当触发信号的按钮被选中时，才判断是否显示
        # 并且只有当删除行或删除列被选中时才显示
        # 使用 self.sender() 获取触发信号的对象，检查它是否是删除按钮且被选中
        sender = self.sender()
        show = False
        delete_operations = [self.delete_rows_radio, self.delete_cols_radio, 
                           self.delete_hidden_rows_radio, self.delete_hidden_cols_radio]
        
        if sender in delete_operations and checked:
            show = True
        elif sender not in delete_operations and checked:
             # 如果是其他按钮被选中，则隐藏
            show = False
        else:
            # 如果是删除按钮被取消选中，也需要重新判断是否还有其他删除按钮被选中
            show = any(radio.isChecked() for radio in delete_operations)
            
        self.delete_merge_group.setVisible(show)

    def toggle_position_input(self, checked):
        """切换位置输入框的启用/禁用状态"""
        # 只有当触发信号的按钮被选中时，才判断是否启用位置输入框
        sender = self.sender()
        
        # 删除隐藏行列操作不需要位置输入
        no_position_operations = [self.delete_hidden_rows_radio, self.delete_hidden_cols_radio]
        
        if sender in no_position_operations and checked:
            # 如果选中的是删除隐藏行列操作，禁用位置输入框
            self.position_edit.setEnabled(False)
            self.position_edit.setPlaceholderText("此操作不需要位置参数")
        elif checked:
            # 如果选中的是其他操作，启用位置输入框
            self.position_edit.setEnabled(True)
            self.position_edit.setPlaceholderText("")
     
    def add_unmerge_step(self):
        """添加拆分合并单元格步骤"""
        try:
            # 确定操作类型和参数
            operation = None
            params = {}
            
            # 根据选中的模式确定操作类型和参数
            if self.unmerge_all_radio.isChecked():
                operation = 'process_merged_cells_all'
                params['action'] = 'keep_value' if self.unmerge_keep_value_radio.isChecked() else 'unmerge'
                params['unmerge'] = 'all'
                desc = f"拆分所有合并单元格({'保留值' if params['action'] == 'keep_value' else '仅拆分'})"
            elif self.unmerge_specific_radio.isChecked():
                operation = 'process_merged_cells_specific'
                params['action'] = 'keep_value' if self.unmerge_keep_value_radio.isChecked() else 'unmerge'

                
                # 获取并验证单元格范围或单个单元格
                range_str = self.unmerge_range_edit.text().strip()
                if not range_str:
                    QMessageBox.warning(self, "警告", "请输入要拆分的单元格范围或单个单元格！")
                    return
                    
                # 处理中文冒号
                range_str = range_str.replace('：', ':')
                
                # 验证输入是有效单元格范围还是单个单元格
                if not (self.validate_cell_range(range_str) or self.is_valid_cell(range_str)):
                    QMessageBox.warning(self, "输入错误", 
                        "请输入有效的单元格范围（如 A1:B50）或单个单元格（如 C3）！")
                    return
                    
                params['range_str'] = range_str
                desc = f"拆分指定范围 {range_str} ({'保留值' if params['action'] == 'keep_value' else '仅拆分'})"
            
            # 添加步骤
            if operation:
                self.add_step(desc, {'operation': operation, 'params': params})
                # 清空输入框（如果适用）
                if self.unmerge_specific_radio.isChecked():
                    self.unmerge_range_edit.clear()
                # QMessageBox.information(self, "成功", "已添加拆分合并单元格步骤")
            else:
                QMessageBox.warning(self, "警告", "请选择拆分合并单元格的模式！")
                
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            print(f"添加拆分合并单元格步骤失败: {error_msg}")
            QMessageBox.critical(self, "错误", f"添加拆分合并单元格步骤失败: {str(e)}")

    def insert_unmerge_step(self):
        """插入拆分合并单元格步骤"""
        try:
            # 确定操作类型和参数
            operation = None
            params = {}
            
            # 根据选中的模式确定操作类型和参数
            if self.unmerge_all_radio.isChecked():
                operation = 'process_merged_cells_all'
                params['action'] = 'keep_value' if self.unmerge_keep_value_radio.isChecked() else 'unmerge'

                desc = f"拆分所有合并单元格({'保留值' if params['action'] == 'keep_value' else '仅拆分'})"
            elif self.unmerge_specific_radio.isChecked():
                operation = 'process_merged_cells_specific'
                params['action'] = 'keep_value' if self.unmerge_keep_value_radio.isChecked() else 'unmerge'

                
                # 获取并验证单元格范围或单个单元格
                range_str = self.unmerge_range_edit.text().strip()
                if not range_str:
                    QMessageBox.warning(self, "警告", "请输入要拆分的单元格范围或单个单元格！")
                    return
                    
                # 处理中文冒号
                range_str = range_str.replace('：', ':')
                
                # 验证输入是有效单元格范围还是单个单元格
                if not (self.validate_cell_range(range_str) or self.is_valid_cell(range_str)):
                    QMessageBox.warning(self, "输入错误", 
                        "请输入有效的单元格范围（如 A1:B50）或单个单元格（如 C3）！")
                    return
                    
                params['range_str'] = range_str
                desc = f"拆分指定范围 {range_str} ({'保留值' if params['action'] == 'keep_value' else '仅拆分'})"
            
            # 插入步骤
            if operation:
                self.insert_specific_step(desc, {'operation': operation, 'params': params})
                # 清空输入框（如果适用）
                if self.unmerge_specific_radio.isChecked():
                    self.unmerge_range_edit.clear()
                # QMessageBox.information(self, "成功", "已插入拆分合并单元格步骤")
            else:
                QMessageBox.warning(self, "警告", "请选择拆分合并单元格的模式！")
                
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            print(f"插入拆分合并单元格步骤失败: {error_msg}")
            QMessageBox.critical(self, "错误", f"插入拆分合并单元格步骤失败: {str(e)}")
 

    def eventFilter(self, obj, event):
        """事件过滤器，处理回车键事件"""
        if event.type() == QEvent.KeyPress and event.key() == Qt.Key_Return:
            if obj == self.position_edit:
                self.add_row_col_step()
                return True
            elif obj == self.unmerge_range_edit:
                self.add_unmerge_step()
            elif obj == self.merge_range_edit:
                self.add_merge_step()                
                return True
            elif obj == self.create_ws_name_edit:
                self.add_create_worksheet_step()
                return True
            elif obj == self.delete_ws_name_edit:
                self.add_delete_worksheet_step()
                return True
            elif hasattr(self, 'font_range_edit') and obj == self.font_range_edit:
                self.add_font_color_step()
                return True
            elif hasattr(self, 'fill_range_edit') and obj == self.fill_range_edit:
                self.add_fill_color_step()
                return True
            elif hasattr(self, 'border_range_edit') and obj == self.border_range_edit:
                self.add_border_step()
                return True
            elif hasattr(self, 'cell_position_edit') and obj == self.cell_position_edit:
                self.add_cell_content_step()
                return True
        return super().eventFilter(obj, event)


