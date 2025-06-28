# Excel批量处理工具

## 项目简介

这是一个基于Python开发的Excel批量处理工具，提供图形用户界面，可以对多个Excel文件进行批量操作，包括公式转值、合并单元格处理、行列操作等功能。

## 功能特点

- **公式转值**：将Excel中的公式转换为实际值，保留单元格格式
- **合并单元格处理**：可选择仅拆分合并单元格或拆分并保留值
- **工作表管理**：支持新建、删除、重命名工作表
- **行列操作**：支持批量插入/删除/隐藏/显示行或列
- **批量处理**：可同时处理多个Excel文件
- **自动备份**：操作前自动创建备份，确保数据安全
- **步骤管理**：支持添加、编辑、删除和调整步骤顺序
- **执行报告**：生成详细的操作执行报告

## V2.0版本主要更新

- 支持自定义单元格字体颜色、背景颜色
- 支持批量删除隐藏行列功能
- 优化遇到合并单元格的处理方式
- 优化批量文件处理逻辑，支持同时处理多个文件
- 完善错误处理机制，提高处理稳定性

## 项目结构
```
├── __main__.py                 # 程序入口点，负责启动应用程序
├── core.py                      # Excel处理核心功能实现
├── models.py                    # 数据模型定义 
├── processing.py                # 后台处理任务实现
├── execution.py                 # 执行操作的实现
├── message_utils.py             # 消息处理工具模块
├── report.py                    # 报告生成模块
├── utils.py                     # 通用工具函数
├── ui/                          # UI相关模块
│   ├── __init__.py              # 包初始化文件
│   ├── main_window.py           # 主窗口界面实现
│   ├── base_window.py           # 基础窗口类
│   ├── file_operations.py       # 文件操作界面
│   ├── step_operations.py       # 步骤操作界面
│   ├── worksheet_operations.py  # 工作表操作界面
│   └── row_col_operations.py    # 行列操作界面
├── image/                       # 图像资源文件夹
│   └── icon.ico                 # 应用图标
├── requirements.txt             # 项目依赖配置
├── excel_tool.spec              # PyInstaller打包配置文件
├── excel_tool.iss               # Inno Setup安装程序配置文件
└── build_excel_tool.ps1          # 打包自动化脚本
```

## 模块说明

### 1. __main__.py
程序的入口点，负责创建QApplication实例并启动主窗口。主要职责：
- 初始化Qt应用程序
- 创建并显示主窗口
- 进入应用程序主循环

### 2. main_window.py
实现了应用程序的主窗口界面，包含以下主要功能：
- 文件选择面板：支持添加单个文件或整个文件夹
- 操作步骤列表：可以添加、编辑、删除和调整步骤顺序
- 功能选项卡：
  - 公式转值
  - 合并单元格处理（合并/拆分）
  - 工作表管理（新建/删除）
  - 行列操作（插入/删除/隐藏/显示）

### 3. core.py
Excel处理的核心模块，提供了所有Excel操作的具体实现：
- 文件备份功能
- 工作簿的加载和保存
- 公式转值处理
- 合并单元格操作
- 工作表管理
- 行列操作处理

### 核心模块

- **core.py**: Excel处理核心模块，提供Excel文件的各种批量处理功能
- **cell_format_module.py**: 单元格格式化模块，提供Excel单元格格式化相关功能（字体颜色、填充颜色、边框等）
- **processing.py**: 处理模块，负责执行操作步骤
- **models.py**: 模型类模块，提供步骤项等数据模型的定义
- **excel_utils.py**: Excel工具模块，提供Excel相关的工具函数
- **cell_utils.py**: 单元格工具模块，提供单元格相关的工具函数

### 4. models.py
定义了程序中使用的数据模型，主要包括：
- 操作步骤的数据结构
- 相关配置信息的模型定义

### 5. processing.py
实现后台处理任务，负责：
- 在后台线程中执行Excel处理操作
- 提供进度更新和错误处理机制

### 6. execution.py
执行操作的实现模块，提供：
- 执行功能混入类，可被主窗口继承使用
- 步骤执行的具体实现
- 执行进度显示和结果处理

### 7. message_utils.py
消息处理工具模块，提供：
- 执行结果消息的格式化处理
- 移除重复信息，统一显示格式
- 错误信息的优化处理

### 8. report.py
报告生成模块，负责：
- 生成操作执行报告
- 记录处理结果和统计信息

### 9. utils.py
通用工具函数模块，提供：
- 文件和路径处理函数
- 日期时间处理函数
- 其他通用辅助功能

## 模块依赖关系
```
__main__.py
  └── ui/main_window.py
       ├── ui/base_window.py
       ├── ui/file_operations.py
       ├── ui/step_operations.py
       ├── ui/worksheet_operations.py
       ├── ui/row_col_operations.py
       ├── models.py
       ├── core.py
       ├── processing.py
       ├── execution.py
       ├── message_utils.py
       ├── report.py
       └── utils.py
       
打包流程:
excel_tool.spec → build_excel_tool.ps1 → dist/excel_tool_deps
  └── excel_tool.iss → output/ExcelToolSetup.exe
```

- __main__.py 依赖 ui/main_window.py 来创建主窗口
- ui/main_window.py 依赖：
  - ui/base_window.py 基础窗口功能
  - ui/file_operations.py 文件操作界面
  - ui/step_operations.py 步骤操作界面
  - ui/worksheet_operations.py 工作表操作界面
  - ui/row_col_operations.py 行列操作界面
  - models.py 获取数据模型
  - core.py 调用Excel处理功能
  - processing.py 执行后台处理任务
- processing.py 依赖：
  - core.py 执行具体的Excel操作
  - execution.py 执行操作的实现
  - message_utils.py 处理执行结果消息
- report.py 依赖 utils.py 生成报告

## 开发环境要求
- Python 3.6+
- PyQt5
- openpyxl

## 使用说明
1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 运行程序：
```bash
python -m __main__
```
或直接运行：
```bash
python __main__.py
```

3. 使用步骤：
   1. 在左侧面板添加需要处理的Excel文件
   2. 在右侧选项卡中选择要执行的操作
   3. 设置操作参数并添加到步骤列表
   4. 调整步骤顺序（如需要）
   5. 点击"执行"按钮开始处理

## 注意事项
- 程序会自动备份原始文件
- 建议在处理重要文件前先进行测试
- 大量文件处理时请耐心等待


## 安装说明

### 环境要求

- Python 3.6+
- Windows操作系统

### 安装步骤

1. 安装依赖包：

```bash
pip install -r requirements.txt
```

## 使用说明

1. 运行程序：

```bash
python -m __main__
```
或直接运行：
```bash
python __main__.py
```

2. 添加要处理的Excel文件：
   - 点击"添加文件"按钮选择单个或多个Excel文件
   - 或点击"添加文件夹"按钮选择包含Excel文件的文件夹

3. 选择要执行的操作：
   - **公式转值**：将所有公式转换为实际值
   - **合并单元格处理**：选择仅拆分或拆分并保留值
   - **工作表管理**：新建、删除或重命名工作表
   - **行列操作**：选择插入/删除/隐藏/显示行/列，并设置位置和数量

4. 点击"执行操作"按钮开始处理

5. 处理完成后，会显示操作结果和备份位置

## 注意事项

- 所有操作前会自动创建备份，备份文件保存在用户选择的输出目录中
- 工作表索引从0开始，即第一个工作表的索引为0
- 行列位置从1开始，即第一行/列的位置为1
- 可以通过调整步骤顺序来优化处理流程
- 执行完成后会生成详细的操作报告

## 开发说明

- 使用PyQt5构建图形界面
- 使用openpyxl库处理Excel文件
- 采用多线程设计，避免界面卡顿

## 打包说明

### 打包工具准备
1. 安装PyInstaller和Inno Setup:
```bash
pip install pyinstaller
```
2. 下载并安装Inno Setup编译器

### 打包流程
1. 使用PyInstaller生成可执行文件:
```bash
pyinstaller excel_tool.spec
```
2. 使用Inno Setup创建安装程序:
```bash
iscc excel_tool.iss
```

### 打包相关文件说明
- **excel_tool.spec**: PyInstaller配置文件，定义打包参数和依赖项
- **excel_tool.iss**: Inno Setup脚本文件，定义安装程序配置
- **build_excel_tool.ps1**: 自动化打包脚本，一键执行iss文件并生成安装程序和调试结果

### 注意事项
- 打包前请确保所有依赖已正确安装