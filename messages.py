#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
消息文本模块
集中管理用户可见的文本字符串
"""

# 通用消息
MSG_SUCCESS = "成功"
MSG_WARNING = "警告"
MSG_ERROR = "错误"
MSG_INFO = "信息"

# 按钮文本
BTN_ADD_TO_STEPS = "添加到步骤"
BTN_INSERT_TO_STEPS = "插入到当前步骤下方"

# 单元格格式操作相关文本
FONT_COLOR_TAB_TITLE = "字体颜色"
FILL_COLOR_TAB_TITLE = "填充颜色"
BORDER_TAB_TITLE = "单元格边框"
CELL_CONTENT_TAB_TITLE = "单元格内容"

# 单元格范围提示
CELL_RANGE_PLACEHOLDER = "例如: A1:B50 或 A1：B50 或单个单元格 B50 (支持中文冒号)"
CELL_RANGE_LABEL = "单元格范围:"
CELL_RANGE_ERROR = "请输入有效的单元格范围（如 A1:B50）或单个单元格（如 C3）！"
CELL_RANGE_EMPTY = "请输入单元格范围或单个单元格！"

# 颜色选择相关文本
COLOR_RED = "红色"
COLOR_GREEN = "绿色"
COLOR_BLUE = "蓝色"
COLOR_BLACK = "黑色"
COLOR_WHITE = "白色"
COLOR_YELLOW = "黄色"
COLOR_PURPLE = "紫色"
COLOR_ORANGE = "橙色"
COLOR_GRAY = "灰色"
COLOR_LABEL = "选择颜色:"

# 范围选择相关文本
RANGE_MODE_LABEL = "应用范围:"
RANGE_MODE_SPECIFIC = "指定范围"
RANGE_MODE_ENTIRE_SHEET = "整个工作表"

# 边框操作相关文本
BORDER_MODE_LABEL = "边框操作:"
BORDER_MODE_ADD = "添加所有边框"
BORDER_MODE_REMOVE = "移除所有边框"

# 单元格内容修改相关文本
CELL_CONTENT_LABEL = "单元格位置:"
CELL_CONTENT_PLACEHOLDER = "输入单元格位置，例如: A1"
CELL_NEW_CONTENT_LABEL = "新内容（支持自动获取并修改合并单元格的值）:"
CELL_NEW_CONTENT_PLACEHOLDER = "输入要设置的新内容 (按Enter换行)"
CELL_CONTENT_ERROR = "请输入有效的单元格位置！"
CELL_CONTENT_EMPTY = "请输入单元格位置！"

# 操作描述
OPERATION_FONT_COLOR = "修改字体颜色"
OPERATION_FILL_COLOR = "修改填充颜色"
OPERATION_ADD_BORDER = "添加单元格边框"
OPERATION_REMOVE_BORDER = "移除单元格边框"
OPERATION_MODIFY_CONTENT = "修改单元格内容"

# 范围描述
RANGE_DESC_SPECIFIC = "指定范围"
RANGE_DESC_ENTIRE_SHEET = "整个工作表"

# 错误消息
ERROR_ADD_STEP = "添加步骤失败"
ERROR_INSERT_STEP = "插入步骤失败"