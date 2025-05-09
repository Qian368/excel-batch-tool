#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
执行模块
提供Excel批量处理工具的执行功能实现
"""

from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QFileDialog
from PyQt5.QtCore import Qt

from core import ExcelProcessor
from processing import ProcessingThread
from models import StepItem
from message_utils import format_result_message


class ExecutionMixin:
    """执行功能混入类，提供执行相关的功能实现"""
    
    def init_execution(self):
        """初始化执行功能"""
        # 初始化处理器
        self.processor = ExcelProcessor()
        
        # 连接执行按钮的点击事件
        if hasattr(self, 'execute_btn'):
            self.execute_btn.clicked.connect(self.execute_steps)
    
    def execute_steps(self):
        """执行所有步骤"""
        # 检查是否选择了文件
        if not self.file_paths:
            QMessageBox.warning(self, "警告", "请先选择要处理的Excel文件！")
            return
        
        # 检查是否有步骤
        if not self.steps:
            QMessageBox.warning(self, "警告", "请先添加要执行的操作步骤！")
            return
        
        # 获取输出目录
        output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if not output_dir:
            return
            
        # 设置输出目录
        self.processor.set_output_dir(output_dir)
        
        # 创建进度对话框
        self.progress_dialog = QProgressDialog("正在处理...", "取消", 0, 100, self)
        self.progress_dialog.setWindowTitle("执行进度")
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setAutoReset(False)
        
        # 创建处理线程
        self.processing_thread = ProcessingThread(self.processor, self.steps, self.file_paths)
        
        # 连接信号
        self.processing_thread.progress_updated.connect(self.update_progress)
        self.processing_thread.operation_complete.connect(self.handle_operation_complete)
        self.processing_thread.step_results_updated.connect(self.show_step_results)
        
        # 打印调试信息
        print(f"开始执行 {len(self.steps)} 个步骤:")
        for i, step in enumerate(self.steps, 1):
            print(f"步骤 {i}: {step.operation} - {step.params}")
        
        # 启动线程
        self.processing_thread.start()
    
    def update_progress(self, value):
        """更新进度条"""
        if self.progress_dialog.wasCanceled():
            # 如果用户取消了操作，终止线程
            self.processing_thread.terminate()
            self.processing_thread.wait()
            return
        
        self.progress_dialog.setValue(value)
    
    def handle_operation_complete(self, success, message):
        """处理操作完成事件"""
        self.progress_dialog.close()
        
        if success:
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.critical(self, "错误", message)
            
    def show_step_results(self, step_results):
        """显示步骤执行结果对话框，并询问用户是否生成Excel报告"""
        # 在控制台输出结果
        print("执行结果详情:")
        for result in step_results:
            status = "✓" if result['success'] else "✗"
            print(f"{status} {result['message']}")
        
        # 创建结果对话框
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QPushButton, QHBoxLayout
        from PyQt5.QtCore import Qt
        
        dialog = QDialog(self)
        dialog.setWindowTitle("执行结果")
        dialog.setMinimumWidth(800)
        dialog.setMinimumHeight(400)
        
        layout = QVBoxLayout()
        
        # 添加标题
        title_label = QLabel("步骤执行结果")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(12)
        font.setBold(True)
        title_label.setFont(font)
        layout.addWidget(title_label)
        
        # 创建表格
        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["步骤", "操作", "结果", "详细信息"])
        table.setRowCount(len(step_results))
        
        # 使用StepItem类中定义的操作类型映射
        operation_name_map = StepItem.operation_desc
        
        # 设置列宽
        table.setColumnWidth(0, 60)   # 步骤
        table.setColumnWidth(1, 150)  # 操作
        table.setColumnWidth(2, 80)   # 结果
        table.setColumnWidth(3, 450)  # 详细信息
        
        # 填充数据
        for i, result in enumerate(step_results):
            # 步骤
            step_item = QTableWidgetItem(str(result['step']))
            step_item.setTextAlignment(Qt.AlignCenter)
            table.setItem(i, 0, step_item)
            
            # 操作 - 显示中文名称
            operation_name = result['operation']
            # 如果操作名在映射表中，则使用中文名称
            if operation_name in operation_name_map:
                operation_name = operation_name_map[operation_name]
            operation_item = QTableWidgetItem(operation_name)
            table.setItem(i, 1, operation_item)
            
            # 结果
            result_text = "成功" if result['success'] else "失败"
            result_item = QTableWidgetItem(result_text)
            result_item.setTextAlignment(Qt.AlignCenter)
            # 设置背景色
            if result['success']:
                result_item.setBackground(Qt.green)
            else:
                result_item.setBackground(Qt.red)
            table.setItem(i, 2, result_item)
            
            # 使用公共函数处理消息格式
            message = format_result_message(result)
            message_item = QTableWidgetItem(message)
            table.setItem(i, 3, message_item)
        
        layout.addWidget(table)
        
        # 添加按钮
        button_layout = QHBoxLayout()
        
        generate_report_btn = QPushButton("生成Excel报告")
        close_btn = QPushButton("关闭")
        
        button_layout.addWidget(generate_report_btn)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        dialog.setLayout(layout)
        
        # 连接按钮信号
        generate_report_btn.clicked.connect(lambda: self.generate_excel_report(step_results, dialog))
        close_btn.clicked.connect(dialog.accept)
        
        # 显示对话框
        dialog.exec_()
    
    def generate_excel_report(self, step_results, dialog=None):
        """生成Excel报告"""
        try:
            from report import generate_report
            report_path = generate_report(step_results, self.processor.output_dir)
            
            # 显示成功消息
            QMessageBox.information(self, "报告生成成功", f"执行报告已生成: {report_path}")
            
            # 如果对话框存在，关闭它
            if dialog:
                dialog.accept()
                
        except Exception as e:
            QMessageBox.critical(self, "报告生成失败", f"生成报告时出错: {str(e)}")