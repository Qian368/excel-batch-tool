#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
步骤编辑器模块
提供编辑已添加步骤的对话框界面
"""

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QComboBox, QRadioButton, QButtonGroup, QGroupBox,
    QMessageBox, QTextEdit
)
from PyQt5.QtCore import Qt

from messages import (
    COLOR_RED, COLOR_GREEN, COLOR_BLUE, COLOR_BLACK, COLOR_WHITE, 
    COLOR_YELLOW, COLOR_PURPLE, COLOR_ORANGE, COLOR_GRAY,
    COLOR_LABEL, RANGE_MODE_LABEL, RANGE_MODE_SPECIFIC, RANGE_MODE_ENTIRE_SHEET,
    BORDER_MODE_LABEL, BORDER_MODE_ADD, BORDER_MODE_REMOVE,
    CELL_RANGE_LABEL, CELL_RANGE_PLACEHOLDER,
    CELL_CONTENT_LABEL, CELL_NEW_CONTENT_LABEL, CELL_CONTENT_PLACEHOLDER, CELL_NEW_CONTENT_PLACEHOLDER
)


class StepEditorDialog(QDialog):
    """
    步骤编辑对话框，用于编辑已添加的操作步骤
    """
    
    def __init__(self, step, parent=None):
        super().__init__(parent)
        self.step = step
        self.operation = step.operation
        self.params = step.params
        self.result = None
        self.init_ui()
        self.load_step_data()
    
    def init_ui(self):
        """
        初始化用户界面
        """
        self.setWindowTitle("编辑步骤")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # 根据操作类型创建不同的编辑界面
        if self.operation == 'change_font_color':
            self.setup_font_color_ui(layout)
        elif self.operation == 'change_fill_color':
            self.setup_fill_color_ui(layout)
        elif self.operation == 'add_border' or self.operation == 'remove_border':
            self.setup_border_ui(layout)
        elif self.operation == 'modify_cell_content':
            self.setup_cell_content_ui(layout)
        elif self.operation in ['delete_hidden_rows', 'delete_hidden_columns']:
            self.setup_delete_hidden_ui(layout)
        else:
            # 对于其他类型的步骤，显示一个简单的信息
            layout.addWidget(QLabel(f"当前不支持编辑此类型的步骤: {self.operation}"))
        
        # 按钮布局
        button_layout = QHBoxLayout()
        
        self.ok_btn = QPushButton("确定")
        self.ok_btn.clicked.connect(self.accept)
        
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(self.ok_btn)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
    
    def setup_font_color_ui(self, layout):
        """
        设置字体颜色编辑界面
        """
        # 颜色选择
        color_layout = QHBoxLayout()
        color_layout.addWidget(QLabel(COLOR_LABEL))
        self.font_color_combo = QComboBox()
        self.font_color_combo.addItems([
            COLOR_RED, COLOR_GREEN, COLOR_BLUE, COLOR_BLACK, COLOR_WHITE,
            COLOR_YELLOW, COLOR_PURPLE, COLOR_ORANGE, COLOR_GRAY
        ])
        color_layout.addWidget(self.font_color_combo)
        layout.addLayout(color_layout)
        
        # 应用范围选择
        range_mode_layout = QHBoxLayout()
        range_mode_layout.addWidget(QLabel(RANGE_MODE_LABEL))
        self.font_range_mode_group = QButtonGroup()
        self.font_specific_radio = QRadioButton(RANGE_MODE_SPECIFIC)
        self.font_entire_sheet_radio = QRadioButton(RANGE_MODE_ENTIRE_SHEET)
        self.font_range_mode_group.addButton(self.font_specific_radio)
        self.font_range_mode_group.addButton(self.font_entire_sheet_radio)
        range_mode_layout.addWidget(self.font_specific_radio)
        range_mode_layout.addWidget(self.font_entire_sheet_radio)
        layout.addLayout(range_mode_layout)
        
        # 单元格范围输入
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel(CELL_RANGE_LABEL))
        self.font_range_edit = QLineEdit()
        self.font_range_edit.setPlaceholderText(CELL_RANGE_PLACEHOLDER)
        range_layout.addWidget(self.font_range_edit)
        layout.addLayout(range_layout)
        
        # 连接单选按钮信号以控制范围输入框的可见性
        self.font_specific_radio.toggled.connect(
            lambda checked: self.font_range_edit.setEnabled(checked))
    
    def setup_fill_color_ui(self, layout):
        """
        设置填充颜色编辑界面
        """
        # 颜色选择
        color_layout = QHBoxLayout()
        color_layout.addWidget(QLabel(COLOR_LABEL))
        self.fill_color_combo = QComboBox()
        self.fill_color_combo.addItems([
            COLOR_RED, COLOR_GREEN, COLOR_BLUE, COLOR_BLACK, COLOR_WHITE,
            COLOR_YELLOW, COLOR_PURPLE, COLOR_ORANGE, COLOR_GRAY
        ])
        color_layout.addWidget(self.fill_color_combo)
        layout.addLayout(color_layout)
        
        # 应用范围选择
        range_mode_layout = QHBoxLayout()
        range_mode_layout.addWidget(QLabel(RANGE_MODE_LABEL))
        self.fill_range_mode_group = QButtonGroup()
        self.fill_specific_radio = QRadioButton(RANGE_MODE_SPECIFIC)
        self.fill_entire_sheet_radio = QRadioButton(RANGE_MODE_ENTIRE_SHEET)
        self.fill_range_mode_group.addButton(self.fill_specific_radio)
        self.fill_range_mode_group.addButton(self.fill_entire_sheet_radio)
        range_mode_layout.addWidget(self.fill_specific_radio)
        range_mode_layout.addWidget(self.fill_entire_sheet_radio)
        layout.addLayout(range_mode_layout)
        
        # 单元格范围输入
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel(CELL_RANGE_LABEL))
        self.fill_range_edit = QLineEdit()
        self.fill_range_edit.setPlaceholderText(CELL_RANGE_PLACEHOLDER)
        range_layout.addWidget(self.fill_range_edit)
        layout.addLayout(range_layout)
        
        # 连接单选按钮信号以控制范围输入框的可见性
        self.fill_specific_radio.toggled.connect(
            lambda checked: self.fill_range_edit.setEnabled(checked))
    
    def setup_border_ui(self, layout):
        """
        设置边框编辑界面
        """
        # 边框操作选择
        if self.operation == 'add_border':
            # 对于添加边框，不需要显示边框模式选择
            pass
        elif self.operation == 'remove_border':
            # 对于移除边框，不需要显示边框模式选择
            pass
        
        # 单元格范围输入
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel(CELL_RANGE_LABEL))
        self.border_range_edit = QLineEdit()
        self.border_range_edit.setPlaceholderText(CELL_RANGE_PLACEHOLDER)
        range_layout.addWidget(self.border_range_edit)
        layout.addLayout(range_layout)
    
    def setup_cell_content_ui(self, layout):
        """
        设置单元格内容编辑界面
        """
        # 单元格位置输入
        position_layout = QHBoxLayout()
        position_layout.addWidget(QLabel(CELL_CONTENT_LABEL))
        self.cell_position_edit = QLineEdit()
        self.cell_position_edit.setPlaceholderText(CELL_CONTENT_PLACEHOLDER)
        position_layout.addWidget(self.cell_position_edit)
        layout.addLayout(position_layout)
        
        # 新内容输入
        content_layout = QHBoxLayout()
        content_layout.addWidget(QLabel(CELL_NEW_CONTENT_LABEL))
        self.cell_content_edit = QTextEdit()
        self.cell_content_edit.setPlaceholderText(CELL_NEW_CONTENT_PLACEHOLDER)
        content_layout.addWidget(self.cell_content_edit)
        layout.addLayout(content_layout)
    
    def load_step_data(self):
        """
        加载步骤数据到界面控件
        """
        if self.operation == 'change_font_color':
            # 设置颜色
            color = self.params.get('color', COLOR_BLACK)
            index = self.font_color_combo.findText(color)
            if index >= 0:
                self.font_color_combo.setCurrentIndex(index)
            
            # 设置范围模式
            range_mode = self.params.get('range_mode', 'specific')
            if range_mode == 'entire_sheet':
                self.font_entire_sheet_radio.setChecked(True)
            else:
                self.font_specific_radio.setChecked(True)
            
            # 设置范围
            range_str = self.params.get('range_str', '')
            self.font_range_edit.setText(range_str)
            self.font_range_edit.setEnabled(range_mode == 'specific')
            
        elif self.operation == 'change_fill_color':
            # 设置颜色
            color = self.params.get('color', COLOR_BLACK)
            index = self.fill_color_combo.findText(color)
            if index >= 0:
                self.fill_color_combo.setCurrentIndex(index)
            
            # 设置范围模式
            range_mode = self.params.get('range_mode', 'specific')
            if range_mode == 'entire_sheet':
                self.fill_entire_sheet_radio.setChecked(True)
            else:
                self.fill_specific_radio.setChecked(True)
            
            # 设置范围
            range_str = self.params.get('range_str', '')
            self.fill_range_edit.setText(range_str)
            self.fill_range_edit.setEnabled(range_mode == 'specific')
            
        elif self.operation == 'add_border' or self.operation == 'remove_border':
            # 设置范围
            range_str = self.params.get('range_str', '')
            self.border_range_edit.setText(range_str)
            
        elif self.operation == 'modify_cell_content':
            # 设置单元格位置
            cell_position = self.params.get('position', '')
            self.cell_position_edit.setText(cell_position)
            
            # 设置新内容
            new_content = self.params.get('content', '')
            self.cell_content_edit.setText(new_content)
            
        elif self.operation in ['delete_hidden_rows', 'delete_hidden_columns']:
            # 加载删除隐藏行列数据
            merge_mode = self.params.get('merge_mode', 'ignore')
            merge_mode_text_map = {
                'ignore': '忽略合并单元格',
                'unmerge_only': '仅拆分合并单元格',
                'unmerge_keep_value': '拆分并保留值'
            }
            merge_mode_text = merge_mode_text_map.get(merge_mode, '忽略合并单元格')
            if hasattr(self, 'merge_mode_combo'):
                index = self.merge_mode_combo.findText(merge_mode_text)
                if index >= 0:
                    self.merge_mode_combo.setCurrentIndex(index)
    
    def accept(self):
        """
        确认编辑，保存修改后的步骤数据
        """
        try:
            if self.operation == 'change_font_color':
                # 获取颜色
                color = self.font_color_combo.currentText()
                
                # 获取范围模式
                range_mode = 'entire_sheet' if self.font_entire_sheet_radio.isChecked() else 'specific'
                
                # 获取范围
                range_str = self.font_range_edit.text().strip()
                if range_mode == 'specific' and not range_str:
                    QMessageBox.warning(self, "警告", "请输入单元格范围")
                    return
                
                # 更新参数
                self.params = {
                    'color': color,
                    'range_mode': range_mode,
                    'range_str': range_str
                }
                
            elif self.operation == 'change_fill_color':
                # 获取颜色
                color = self.fill_color_combo.currentText()
                
                # 获取范围模式
                range_mode = 'entire_sheet' if self.fill_entire_sheet_radio.isChecked() else 'specific'
                
                # 获取范围
                range_str = self.fill_range_edit.text().strip()
                if range_mode == 'specific' and not range_str:
                    QMessageBox.warning(self, "警告", "请输入单元格范围")
                    return
                
                # 更新参数
                self.params = {
                    'color': color,
                    'range_mode': range_mode,
                    'range_str': range_str
                }
                
            elif self.operation == 'add_border' or self.operation == 'remove_border':
                # 获取范围
                range_str = self.border_range_edit.text().strip()
                if not range_str:
                    QMessageBox.warning(self, "警告", "请输入单元格范围")
                    return
                
                # 更新参数
                self.params = {
                    'range_str': range_str
                }
                if self.operation == 'add_border' and 'border_style' in self.params:
                    # 保留原有的边框样式
                    self.params['border_style'] = self.params.get('border_style', 'thin')
                
            elif self.operation == 'modify_cell_content':
                # 获取单元格位置
                cell_position = self.cell_position_edit.text().strip()
                if not cell_position:
                    QMessageBox.warning(self, "警告", "请输入单元格位置")
                    return
                
                # 获取新内容
                new_content = self.cell_content_edit.toPlainText()
                
                # 更新参数
                self.params = {
                    'position': cell_position,
                    'content': new_content
                }
            
            elif self.operation in ['delete_hidden_rows', 'delete_hidden_columns']:
                # 获取合并单元格处理模式
                if hasattr(self, 'merge_mode_combo'):
                    merge_mode_map = {
                        '忽略合并单元格': 'ignore',
                        '仅拆分合并单元格': 'unmerge_only', 
                        '拆分并保留值': 'unmerge_keep_value'
                    }
                    merge_mode = merge_mode_map.get(self.merge_mode_combo.currentText(), 'ignore')
                    self.params = {'merge_mode': merge_mode}
                else:
                    self.params = {'merge_mode': 'ignore'}
            
            # 保存结果
            self.result = (self.operation, self.params)
            super().accept()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存步骤数据失败: {str(e)}")
    
    def setup_delete_hidden_ui(self, layout):
        """
        设置删除隐藏行列编辑界面
        """
        # 操作说明
        if self.operation == 'delete_hidden_rows':
            operation_label = QLabel("删除所有隐藏行")
        else:
            operation_label = QLabel("删除所有隐藏列")
        layout.addWidget(operation_label)
        
        # 合并单元格处理模式
        merge_layout = QHBoxLayout()
        merge_layout.addWidget(QLabel("合并单元格处理:"))
        self.merge_mode_combo = QComboBox()
        self.merge_mode_combo.addItems([
            '忽略合并单元格',
            '仅拆分合并单元格', 
            '拆分并保留值'
        ])
        merge_layout.addWidget(self.merge_mode_combo)
        layout.addLayout(merge_layout)
    
    def get_result(self):
        """
        获取编辑结果
        
        Returns:
            tuple: (operation, params) 或 None（如果取消）
        """
        return self.result