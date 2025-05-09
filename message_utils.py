#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
消息处理工具模块
提供处理执行结果消息的通用函数
"""


def format_result_message(result):
    """
    格式化执行结果消息，移除重复信息，统一显示格式
    
    Args:
        result: 包含step、success和message字段的结果字典
        
    Returns:
        str: 格式化后的消息
    """
    message = result['message']
    
    # 移除"步骤X: 操作名称(参数)" 格式的前缀，确保不会显示重复的错误信息
    if f"步骤{result['step']}: " in message:
        # 统一处理所有步骤的详细信息
        if " 执行成功" in message:
            # 如果消息中包含额外信息（在 - 之后），则保留
            if " - " in message:
                extra_info = message.split(" - ", 1)[1]
                message = f"执行成功 - {extra_info}"
            else:
                message = "执行成功"
        elif " 执行失败" in message:
            # 统一简化失败信息，只显示"执行失败: 错误原因"格式
            error_parts = message.split(" 执行失败: ", 1)
            if len(error_parts) > 1 and error_parts[1]:
                # 检查错误信息中是否包含重复的步骤信息
                error_msg = error_parts[1]
                # 检查是否包含类似"步骤X: 删除工作表（工作表：XXX）执行失败: "的重复内容
                if "步骤" in error_msg and " 执行失败: " in error_msg:
                    # 提取最后一个错误原因
                    final_error = error_msg.split(" 执行失败: ", 1)[1]
                    message = "执行失败: " + final_error
                # 检查是否包含类似"（工作表：XXX）"的内容，这是删除工作表操作特有的格式
                elif "（工作表：" in error_msg and "）" in error_msg:
                    # 尝试提取括号后面的实际错误信息
                    parts = error_msg.split("）", 1)
                    if len(parts) > 1 and parts[1].strip():
                        # 如果括号后有内容，则使用括号后的内容
                        message = "执行失败: " + parts[1].strip()
                    else:
                        # 否则使用完整错误信息
                        message = "执行失败: " + error_msg
                else:
                    message = "执行失败: " + error_msg
            else:
                # 尝试其他可能的分隔模式
                error_parts = message.split(": ", 1)
                if len(error_parts) > 1 and error_parts[1]:
                    message = "执行失败: " + error_parts[1]
                else:
                    message = "执行失败"
        else:
            # 如果没有明确的成功/失败标记，则尝试移除前缀
            parts = message.split(" ", 1)
            if len(parts) > 1:
                message = parts[1]
    
    return message