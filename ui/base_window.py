#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
基础窗口模块
提供Excel批量处理工具的基础窗口界面实现
"""

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
    QPushButton, QListWidget, QTabWidget, QLabel, QButtonGroup,
    QRadioButton, QLineEdit, QGridLayout, QFileDialog, QMessageBox,
    QProgressDialog
)
from PyQt5.QtCore import Qt
import os

from models import StepItem
from core import ExcelProcessor
from processing import ProcessingThread
from execution import ExecutionMixin
from PyQt5.QtGui import QIcon

class BaseWindow(QMainWindow, ExecutionMixin):
    """基础窗口类，提供基本UI框架"""
    
    def __init__(self):
        super().__init__()
        self.processor = ExcelProcessor()
        self.file_paths = []
        self.steps = []
        self.init_ui()
        self.init_execution()
    
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("Excel批量处理工具")
        self.setGeometry(100, 100, 1000, 800)
        # 设置窗口图标（使用相对路径，基于当前文件位置）
        icon_path = os.path.join(os.path.dirname(__file__), "image", "icon.ico")
        self.setWindowIcon(QIcon(icon_path))
                
        # 创建中央部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QHBoxLayout(central_widget)
        
        # 初始化左侧面板
        self.init_left_panel()
        
        # 初始化右侧面板
        self.init_right_panel()
        
        self.main_layout.setStretch(0, 1)
        self.main_layout.setStretch(1, 2)
    
    def init_left_panel(self):
        """初始化左侧面板"""
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        # 文件选择区域
        file_group = self.create_file_group()
        
        # 步骤列表区域
        steps_group = self.create_steps_group()
        
        left_layout.addWidget(file_group)
        left_layout.addWidget(steps_group)
        
        self.main_layout.addWidget(left_panel)
    
    def create_file_group(self):
        """创建文件选择组件"""
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()
        
        # 按钮布局
        file_btn_layout = QHBoxLayout()
        self.add_files_btn = QPushButton("添加文件")
        self.add_folder_btn = QPushButton("添加文件夹")
        self.clear_files_btn = QPushButton("清空列表")
        
        file_btn_layout.addWidget(self.add_files_btn)
        file_btn_layout.addWidget(self.add_folder_btn)
        file_btn_layout.addWidget(self.clear_files_btn)
        
        # 文件列表
        self.file_list = QListWidget()
        
        file_layout.addLayout(file_btn_layout)
        file_layout.addWidget(self.file_list)
        file_group.setLayout(file_layout)
        
        return file_group
    
    def create_steps_group(self):
        """创建步骤列表组件"""
        steps_group = QGroupBox("操作步骤列表")
        steps_layout = QVBoxLayout()
        
        # 步骤列表
        self.steps_list = QListWidget()
        
        # 步骤操作按钮 - 第一行
        steps_btn_layout1 = QHBoxLayout()
        self.move_up_btn = QPushButton("上移")
        self.move_down_btn = QPushButton("下移")
        self.edit_step_btn = QPushButton("编辑步骤")
        self.delete_step_btn = QPushButton("删除步骤")
        self.clear_steps_btn = QPushButton("清空步骤")

        steps_btn_layout1.addWidget(self.move_up_btn)
        steps_btn_layout1.addWidget(self.move_down_btn)
        steps_btn_layout1.addWidget(self.edit_step_btn)
        steps_btn_layout1.addWidget(self.delete_step_btn)
        steps_btn_layout1.addWidget(self.clear_steps_btn)

        # 步骤操作按钮 - 第二行 (导入/导出)
        steps_btn_layout2 = QHBoxLayout()
        self.export_steps_btn = QPushButton("导出步骤")
        self.import_steps_btn = QPushButton("导入步骤")
        
        # 设置按钮自动填充宽度
        from PyQt5.QtWidgets import QSizePolicy  # 添加必要的import语句
        self.export_steps_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.import_steps_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        steps_btn_layout2.addWidget(self.export_steps_btn)
        steps_btn_layout2.addWidget(self.import_steps_btn)

        steps_layout.addWidget(self.steps_list)
        steps_layout.addLayout(steps_btn_layout1)
        steps_layout.addLayout(steps_btn_layout2)
        steps_group.setLayout(steps_layout)
        return steps_group
    
    def init_right_panel(self):
        """初始化右侧面板"""
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        # 选项卡组件
        self.tab_widget = QTabWidget()
        right_layout.addWidget(self.tab_widget)
        
        # 执行按钮
        self.execute_btn = QPushButton("执行")
        right_layout.addWidget(self.execute_btn)
        
        self.main_layout.addWidget(right_panel)
