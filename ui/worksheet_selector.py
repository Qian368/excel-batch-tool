#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
工作表选择器模块
提供Excel批量处理工具的工作表选择功能
"""

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QCheckBox, QMessageBox
)
from PyQt5.QtCore import Qt
import os
import openpyxl

class WorksheetSelectorDialog(QDialog):
    """
    工作表选择对话框，允许用户选择要操作的工作表
    """
    
    def __init__(self, file_paths, parent=None):
        super().__init__(parent)
        self.file_paths = file_paths
        self.selected_worksheets = {}
        self.init_ui()
        self.load_worksheets()
    
    def init_ui(self):
        """
        初始化用户界面
        """
        self.setWindowTitle("选择工作表")
        self.setMinimumWidth(600)
        self.setMinimumHeight(400)
        
        layout = QVBoxLayout(self)
        
        # 说明标签
        info_label = QLabel("请选择要操作的工作表：")
        layout.addWidget(info_label)
        
        # 工作表列表
        self.worksheet_list = QListWidget()
        self.worksheet_list.setSelectionMode(QListWidget.ExtendedSelection)
        layout.addWidget(self.worksheet_list)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.clicked.connect(self.select_all_worksheets)
        
        self.deselect_all_btn = QPushButton("取消全选")
        self.deselect_all_btn.clicked.connect(self.deselect_all_worksheets)
        
        self.ok_btn = QPushButton("确定")
        self.ok_btn.clicked.connect(self.accept)
        
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(self.select_all_btn)
        button_layout.addWidget(self.deselect_all_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.ok_btn)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
    
    def load_worksheets(self):
        """
        加载所有Excel文件的工作表
        """
        try:
            for file_path in self.file_paths:
                try:
                    # 检查文件是否存在
                    if not os.path.exists(file_path):
                        continue
                    
                    # 加载工作簿
                    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                    
                    # 获取文件名（不含路径）
                    file_name = os.path.basename(file_path)
                    
                    # 为每个工作表创建一个带复选框的列表项
                    for sheet_name in wb.sheetnames:
                        item_text = f"{file_name} - {sheet_name}"
                        item = QListWidgetItem(item_text)
                        item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                        item.setCheckState(Qt.Checked)  # 默认选中
                        
                        # 存储文件路径和工作表名称
                        item.setData(Qt.UserRole, (file_path, sheet_name))
                        
                        self.worksheet_list.addItem(item)
                        
                        # 初始化选中状态
                        if file_path not in self.selected_worksheets:
                            self.selected_worksheets[file_path] = []
                        self.selected_worksheets[file_path].append(sheet_name)
                    
                    # 关闭工作簿
                    wb.close()
                    
                except Exception as e:
                    print(f"加载文件 {file_path} 的工作表失败: {str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载工作表失败: {str(e)}")
    
    def select_all_worksheets(self):
        """
        选择所有工作表
        """
        for i in range(self.worksheet_list.count()):
            item = self.worksheet_list.item(i)
            item.setCheckState(Qt.Checked)
            
            # 更新选中状态
            file_path, sheet_name = item.data(Qt.UserRole)
            if file_path not in self.selected_worksheets:
                self.selected_worksheets[file_path] = []
            if sheet_name not in self.selected_worksheets[file_path]:
                self.selected_worksheets[file_path].append(sheet_name)
    
    def deselect_all_worksheets(self):
        """
        取消选择所有工作表
        """
        for i in range(self.worksheet_list.count()):
            item = self.worksheet_list.item(i)
            item.setCheckState(Qt.Unchecked)
        
        # 清空选中状态
        self.selected_worksheets = {}
    
    def get_selected_worksheets(self):
        """
        获取用户选择的工作表
        
        Returns:
            dict: 文件路径到工作表名称列表的映射
        """
        # 重新扫描列表，确保返回最新的选择状态
        self.selected_worksheets = {}
        
        for i in range(self.worksheet_list.count()):
            item = self.worksheet_list.item(i)
            if item.checkState() == Qt.Checked:
                file_path, sheet_name = item.data(Qt.UserRole)
                
                if file_path not in self.selected_worksheets:
                    self.selected_worksheets[file_path] = []
                
                self.selected_worksheets[file_path].append(sheet_name)
        
        return self.selected_worksheets