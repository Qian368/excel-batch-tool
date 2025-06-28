#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文件操作模块
提供Excel批量处理工具的文件操作相关功能
"""

from PyQt5.QtWidgets import QFileDialog

class FileOperationsMixin:
    """文件操作混入类，提供文件相关的操作方法"""
    
    def add_files(self):
        """添加文件到列表"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        if files:
            self.file_paths.extend(files)
            self.update_file_list()
    
    def add_folder(self):
        """添加文件夹中的所有Excel文件到列表"""
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            import os
            for root, _, files in os.walk(folder):
                for file in files:
                    if file.endswith(('.xlsx', '.xls')):
                        self.file_paths.append(os.path.join(root, file))
            self.update_file_list()
    
    def clear_files(self):
        """清空文件列表"""
        self.file_paths.clear()
        self.update_file_list()
    
    def remove_selected_files(self):
        """从列表中移除选中的文件"""
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            return
            
        # 获取所有选中项的文本（文件路径）
        selected_paths = [item.text() for item in selected_items]
        
        # 从文件路径列表中移除选中的文件
        self.file_paths = [path for path in self.file_paths if path not in selected_paths]
        
        # 更新文件列表显示
        self.update_file_list()
    
    def update_file_list(self):
        """更新文件列表显示"""
        self.file_list.clear()
        for path in self.file_paths:
            self.file_list.addItem(path)