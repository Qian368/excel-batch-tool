#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
单元格格式操作UI模块
提供单元格字体颜色、填充颜色、边框和内容修改等功能的UI界面
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, 
    QPushButton, QLabel, QButtonGroup, QRadioButton,
    QLineEdit, QGridLayout, QMessageBox, QComboBox, QTextEdit
)
from PyQt5.QtCore import Qt

from messages import (
    FONT_COLOR_TAB_TITLE, FILL_COLOR_TAB_TITLE, 
    BORDER_TAB_TITLE, CELL_CONTENT_TAB_TITLE,
    CELL_RANGE_PLACEHOLDER, CELL_RANGE_LABEL, CELL_RANGE_ERROR, CELL_RANGE_EMPTY,
    COLOR_RED, COLOR_GREEN, COLOR_BLUE, COLOR_BLACK, COLOR_WHITE, 
    COLOR_YELLOW, COLOR_PURPLE, COLOR_ORANGE, COLOR_GRAY,
    COLOR_LABEL, RANGE_MODE_LABEL, RANGE_MODE_SPECIFIC, RANGE_MODE_ENTIRE_SHEET,
    BORDER_MODE_LABEL, BORDER_MODE_ADD, BORDER_MODE_REMOVE,
    CELL_CONTENT_LABEL, CELL_CONTENT_PLACEHOLDER, 
    CELL_NEW_CONTENT_LABEL, CELL_NEW_CONTENT_PLACEHOLDER,
    CELL_CONTENT_ERROR, CELL_CONTENT_EMPTY,
    BTN_ADD_TO_STEPS, BTN_INSERT_TO_STEPS
)


class CellFormatOperationsMixin:
    """单元格格式操作混入类，提供单元格格式相关的UI和操作方法"""
    
    def setup_cell_format_tabs(self):
        """设置单元格格式操作相关的选项卡"""
        # 字体颜色选项卡
        self.setup_font_color_tab()
        
        # 填充颜色选项卡
        self.setup_fill_color_tab()
        
        # 单元格边框选项卡
        self.setup_border_tab()
        
        # 单元格内容修改选项卡
        self.setup_cell_content_tab()
    
    def setup_font_color_tab(self):
        """设置字体颜色选项卡"""
        font_color_tab = QWidget()
        font_color_layout = QVBoxLayout(font_color_tab)
        
        # 颜色选择组
        color_group = QGroupBox("字体颜色设置")
        color_layout = QVBoxLayout()
        
        # 颜色选择下拉框
        color_select_layout = QHBoxLayout()
        color_select_layout.addWidget(QLabel(COLOR_LABEL))
        self.font_color_combo = QComboBox()
        self.font_color_combo.addItems([
            COLOR_RED, COLOR_GREEN, COLOR_BLUE, COLOR_BLACK, COLOR_WHITE,
            COLOR_YELLOW, COLOR_PURPLE, COLOR_ORANGE, COLOR_GRAY
        ])
        color_select_layout.addWidget(self.font_color_combo)
        color_layout.addLayout(color_select_layout)
        
        # 应用范围选择
        range_mode_layout = QHBoxLayout()
        range_mode_layout.addWidget(QLabel(RANGE_MODE_LABEL))
        self.font_range_mode_group = QButtonGroup()
        self.font_specific_radio = QRadioButton(RANGE_MODE_SPECIFIC)
        self.font_specific_radio.setChecked(True)
        self.font_entire_sheet_radio = QRadioButton(RANGE_MODE_ENTIRE_SHEET)
        self.font_range_mode_group.addButton(self.font_specific_radio)
        self.font_range_mode_group.addButton(self.font_entire_sheet_radio)
        range_mode_layout.addWidget(self.font_specific_radio)
        range_mode_layout.addWidget(self.font_entire_sheet_radio)
        color_layout.addLayout(range_mode_layout)
        
        # 单元格范围输入
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel(CELL_RANGE_LABEL))
        self.font_range_edit = QLineEdit()
        self.font_range_edit.setPlaceholderText(CELL_RANGE_PLACEHOLDER)
        range_layout.addWidget(self.font_range_edit)
        color_layout.addLayout(range_layout)
        
        # 连接单选按钮信号以控制范围输入框的可见性
        self.font_specific_radio.toggled.connect(
            lambda checked: self.font_range_edit.setEnabled(checked))
        self.font_entire_sheet_radio.toggled.connect(
            lambda checked: self.font_range_edit.setEnabled(not checked))
        
        # 添加按钮
        add_font_color_btn = QPushButton(BTN_ADD_TO_STEPS)
        add_font_color_btn.clicked.connect(self.add_font_color_step)
        color_layout.addWidget(add_font_color_btn)
        
        # 插入按钮
        insert_font_color_btn = QPushButton(BTN_INSERT_TO_STEPS)
        insert_font_color_btn.clicked.connect(self.insert_font_color_step)
        color_layout.addWidget(insert_font_color_btn)
        
        color_group.setLayout(color_layout)
        font_color_layout.addWidget(color_group)
        font_color_layout.addStretch()
        
        self.tab_widget.addTab(font_color_tab, FONT_COLOR_TAB_TITLE)
    
    def setup_fill_color_tab(self):
        """设置填充颜色选项卡"""
        fill_color_tab = QWidget()
        fill_color_layout = QVBoxLayout(fill_color_tab)
        
        # 颜色选择组
        color_group = QGroupBox("填充颜色设置")
        color_layout = QVBoxLayout()
        
        # 颜色选择下拉框
        color_select_layout = QHBoxLayout()
        color_select_layout.addWidget(QLabel(COLOR_LABEL))
        self.fill_color_combo = QComboBox()
        self.fill_color_combo.addItems([
            COLOR_RED, COLOR_GREEN, COLOR_BLUE, COLOR_BLACK, COLOR_WHITE,
            COLOR_YELLOW, COLOR_PURPLE, COLOR_ORANGE, COLOR_GRAY
        ])
        color_select_layout.addWidget(self.fill_color_combo)
        color_layout.addLayout(color_select_layout)
        
        # 应用范围选择
        range_mode_layout = QHBoxLayout()
        range_mode_layout.addWidget(QLabel(RANGE_MODE_LABEL))
        self.fill_range_mode_group = QButtonGroup()
        self.fill_specific_radio = QRadioButton(RANGE_MODE_SPECIFIC)
        self.fill_specific_radio.setChecked(True)
        self.fill_entire_sheet_radio = QRadioButton(RANGE_MODE_ENTIRE_SHEET)
        self.fill_range_mode_group.addButton(self.fill_specific_radio)
        self.fill_range_mode_group.addButton(self.fill_entire_sheet_radio)
        range_mode_layout.addWidget(self.fill_specific_radio)
        range_mode_layout.addWidget(self.fill_entire_sheet_radio)
        color_layout.addLayout(range_mode_layout)
        
        # 单元格范围输入
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel(CELL_RANGE_LABEL))
        self.fill_range_edit = QLineEdit()
        self.fill_range_edit.setPlaceholderText(CELL_RANGE_PLACEHOLDER)
        range_layout.addWidget(self.fill_range_edit)
        color_layout.addLayout(range_layout)
        
        # 连接单选按钮信号以控制范围输入框的可见性
        self.fill_specific_radio.toggled.connect(
            lambda checked: self.fill_range_edit.setEnabled(checked))
        self.fill_entire_sheet_radio.toggled.connect(
            lambda checked: self.fill_range_edit.setEnabled(not checked))
        
        # 添加按钮
        add_fill_color_btn = QPushButton(BTN_ADD_TO_STEPS)
        add_fill_color_btn.clicked.connect(self.add_fill_color_step)
        color_layout.addWidget(add_fill_color_btn)
        
        # 插入按钮
        insert_fill_color_btn = QPushButton(BTN_INSERT_TO_STEPS)
        insert_fill_color_btn.clicked.connect(self.insert_fill_color_step)
        color_layout.addWidget(insert_fill_color_btn)
        
        color_group.setLayout(color_layout)
        fill_color_layout.addWidget(color_group)
        fill_color_layout.addStretch()
        
        self.tab_widget.addTab(fill_color_tab, FILL_COLOR_TAB_TITLE)
    
    def setup_border_tab(self):
        """设置单元格边框选项卡"""
        border_tab = QWidget()
        border_layout = QVBoxLayout(border_tab)
        
        # 边框设置组
        border_group = QGroupBox("边框设置")
        border_inner_layout = QVBoxLayout()
        
        # 边框操作选择
        border_mode_layout = QHBoxLayout()
        border_mode_layout.addWidget(QLabel(BORDER_MODE_LABEL))
        self.border_mode_group = QButtonGroup()
        self.add_border_radio = QRadioButton(BORDER_MODE_ADD)
        self.add_border_radio.setChecked(True)
        self.remove_border_radio = QRadioButton(BORDER_MODE_REMOVE)
        self.border_mode_group.addButton(self.add_border_radio)
        self.border_mode_group.addButton(self.remove_border_radio)
        border_mode_layout.addWidget(self.add_border_radio)
        border_mode_layout.addWidget(self.remove_border_radio)
        border_inner_layout.addLayout(border_mode_layout)
        
        # 单元格范围输入
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel(CELL_RANGE_LABEL))
        self.border_range_edit = QLineEdit()
        self.border_range_edit.setPlaceholderText(CELL_RANGE_PLACEHOLDER)
        range_layout.addWidget(self.border_range_edit)
        border_inner_layout.addLayout(range_layout)
        
        # 添加按钮
        add_border_btn = QPushButton(BTN_ADD_TO_STEPS)
        add_border_btn.clicked.connect(self.add_border_step)
        border_inner_layout.addWidget(add_border_btn)
        
        # 插入按钮
        insert_border_btn = QPushButton(BTN_INSERT_TO_STEPS)
        insert_border_btn.clicked.connect(self.insert_border_step)
        border_inner_layout.addWidget(insert_border_btn)
        
        border_group.setLayout(border_inner_layout)
        border_layout.addWidget(border_group)
        border_layout.addStretch()
        
        self.tab_widget.addTab(border_tab, BORDER_TAB_TITLE)
    
    def setup_cell_content_tab(self):
        """设置单元格内容修改选项卡"""
        content_tab = QWidget()
        content_layout = QVBoxLayout(content_tab)
        
        # 内容修改组
        content_group = QGroupBox("修改单元格内容")
        content_inner_layout = QVBoxLayout()
        
        # 单元格位置输入
        cell_layout = QHBoxLayout()
        cell_layout.addWidget(QLabel(CELL_CONTENT_LABEL))
        self.cell_position_edit = QLineEdit()
        self.cell_position_edit.setPlaceholderText(CELL_CONTENT_PLACEHOLDER)
        cell_layout.addWidget(self.cell_position_edit)
        content_inner_layout.addLayout(cell_layout)
        
        # 新内容输入
        content_inner_layout.addWidget(QLabel(CELL_NEW_CONTENT_LABEL))
        self.cell_content_edit = QTextEdit()
        self.cell_content_edit.setPlaceholderText(CELL_NEW_CONTENT_PLACEHOLDER)
        content_inner_layout.addWidget(self.cell_content_edit)
        
        # 添加按钮
        add_content_btn = QPushButton(BTN_ADD_TO_STEPS)
        add_content_btn.clicked.connect(self.add_cell_content_step)
        content_inner_layout.addWidget(add_content_btn)
        
        # 插入按钮
        insert_content_btn = QPushButton(BTN_INSERT_TO_STEPS)
        insert_content_btn.clicked.connect(self.insert_cell_content_step)
        content_inner_layout.addWidget(insert_content_btn)
        
        content_group.setLayout(content_inner_layout)
        content_layout.addWidget(content_group)
        content_layout.addStretch()
        
        self.tab_widget.addTab(content_tab, CELL_CONTENT_TAB_TITLE)
    
    def add_font_color_step(self):
        """添加字体颜色修改步骤"""
        try:
            params = self._get_font_color_params()
            if params:
                self.add_step('change_font_color', params)
                # 清空文本框
                self.font_range_edit.clear()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"添加字体颜色步骤失败：{str(e)}")
    
    def insert_font_color_step(self):
        """插入字体颜色修改步骤"""
        try:
            params = self._get_font_color_params()
            if params:
                self.insert_specific_step('change_font_color', params)
                # 清空文本框
                self.font_range_edit.clear()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"插入字体颜色步骤失败：{str(e)}")
    
    def _get_font_color_params(self):
        """获取字体颜色参数"""
        # 获取颜色
        color = self.font_color_combo.currentText()
        
        # 获取范围模式
        range_mode = 'entire_sheet' if self.font_entire_sheet_radio.isChecked() else 'specific'
        
        params = {
            'color': color,
            'range_mode': range_mode
        }
        
        # 如果是指定范围模式，验证并获取范围
        if range_mode == 'specific':
            range_str = self.font_range_edit.text().strip()
            if not range_str:
                QMessageBox.warning(self, "警告", CELL_RANGE_EMPTY)
                return None
            
            # 处理中文符号
            range_str = range_str.replace('，', ',').replace('：', ':')
            
            # 验证单元格范围
            if not self.validate_cell_range(range_str) and not self.is_valid_cell(range_str):
                QMessageBox.warning(self, "警告", CELL_RANGE_ERROR)
                return None
            
            params['range_str'] = range_str
        
        return params
    
    def add_fill_color_step(self):
        """添加填充颜色修改步骤"""
        try:
            params = self._get_fill_color_params()
            if params:
                self.add_step('change_fill_color', params)
                # 清空文本框
                self.fill_range_edit.clear()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"添加填充颜色步骤失败：{str(e)}")
    
    def insert_fill_color_step(self):
        """插入填充颜色修改步骤"""
        try:
            params = self._get_fill_color_params()
            if params:
                self.insert_specific_step('change_fill_color', params)
                # 清空文本框
                self.fill_range_edit.clear()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"插入填充颜色步骤失败：{str(e)}")
    
    def _get_fill_color_params(self):
        """获取填充颜色修改参数"""
        # 获取颜色
        color = self.fill_color_combo.currentText()
        
        # 获取范围模式
        range_mode = 'entire_sheet' if self.fill_entire_sheet_radio.isChecked() else 'specific'
        
        params = {
            'color': color,
            'range_mode': range_mode
        }
        
        # 如果是指定范围模式，验证并获取范围
        if range_mode == 'specific':
            range_str = self.fill_range_edit.text().strip()
            if not range_str:
                QMessageBox.warning(self, "警告", CELL_RANGE_EMPTY)
                return None
            
            # 处理中文符号
            range_str = range_str.replace('，', ',').replace('：', ':')
            
            # 验证单元格范围
            if not self.validate_cell_range(range_str) and not self.is_valid_cell(range_str):
                QMessageBox.warning(self, "警告", CELL_RANGE_ERROR)
                return None
            
            params['range_str'] = range_str
        
        return params
    
    def add_border_step(self):
        """添加边框操作步骤"""
        try:
            params = self._get_border_params()
            if params:
                operation = 'add_border' if self.add_border_radio.isChecked() else 'remove_border'
                self.add_step(operation, params)
                # 清空文本框
                self.border_range_edit.clear()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"添加边框操作步骤失败：{str(e)}")
    
    def insert_border_step(self):
        """插入边框操作步骤"""
        try:
            params = self._get_border_params()
            if params:
                operation = 'add_border' if self.add_border_radio.isChecked() else 'remove_border'
                self.insert_specific_step(operation, params)
                # 清空文本框
                self.border_range_edit.clear()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"插入边框操作步骤失败：{str(e)}")
    
    def _get_border_params(self):
        """获取边框操作参数"""
        # 获取单元格范围
        range_str = self.border_range_edit.text().strip()
        if not range_str:
            QMessageBox.warning(self, "警告", CELL_RANGE_EMPTY)
            return None
        
        # 处理中文符号
        range_str = range_str.replace('，', ',').replace('：', ':')
        
        # 验证单元格范围
        if not self.validate_cell_range(range_str) and not self.is_valid_cell(range_str):
            QMessageBox.warning(self, "警告", CELL_RANGE_ERROR)
            return None
        
        return {'range_str': range_str}
    
    def add_cell_content_step(self):
        """添加单元格内容修改步骤"""
        try:
            params = self._get_cell_content_params()
            if params:
                self.add_step('modify_cell_content', params)
                # 清空文本框
                self.cell_position_edit.clear()
                self.cell_content_edit.clear()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"添加单元格内容修改步骤失败：{str(e)}")
    
    def insert_cell_content_step(self):
        """插入单元格内容修改步骤"""
        try:
            params = self._get_cell_content_params()
            if params:
                self.insert_specific_step('modify_cell_content', params)
                # 清空文本框
                self.cell_position_edit.clear()
                self.cell_content_edit.clear()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"插入单元格内容修改步骤失败：{str(e)}")
    
    def _get_cell_content_params(self):
        """获取单元格内容修改参数"""
        # 获取单元格位置
        cell_position = self.cell_position_edit.text().strip()
        if not cell_position:
            QMessageBox.warning(self, "警告", CELL_CONTENT_EMPTY)
            return None
        
        # 验证单元格位置
        if not self.is_valid_cell(cell_position):
            QMessageBox.warning(self, "警告", CELL_CONTENT_ERROR)
            return None
        
        # 获取新内容
        new_content = self.cell_content_edit.toPlainText()
        
        return {
            'position': cell_position,
            'content': new_content
        }
    
    def is_valid_cell(self, cell_ref):
        """验证单元格引用是否有效"""
        import re
        # 匹配单个单元格引用，如A1, AB123等
        return bool(re.match(r'^[A-Za-z]+[0-9]+$', cell_ref))
    
    def validate_cell_range(self, range_str):
        """验证单元格范围是否有效"""
        # 处理多个范围，用逗号分隔
        ranges = range_str.split(',')
        for r in ranges:
            r = r.strip()
            if ':' in r:
                # 处理范围 (如 A1:B5)
                parts = r.split(':')
                if len(parts) != 2:
                    return False
                if not (self.is_valid_cell(parts[0].strip()) and self.is_valid_cell(parts[1].strip())):
                    return False
            elif not self.is_valid_cell(r):
                # 处理单个单元格 (如 A1)
                return False
        return True