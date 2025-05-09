# Excel批量处理工具 API接口文档

## 1. 核心模块接口

### core.py

#### 主要功能
核心模块提供了所有Excel操作的具体实现，包括文件备份、工作簿加载和保存、公式转值、合并单元格操作等。

#### 主要方法

| 方法名 | 功能描述 | 参数 | 返回值 |
| ----- | ------- | ---- | ----- |
| `backup_file` | 创建Excel文件备份 | `file_path`: 文件路径<br>`backup_dir`: 备份目录 | 备份文件路径 |
| `load_workbook` | 加载Excel工作簿 | `file_path`: 文件路径<br>`data_only`: 是否只加载数据 | 工作簿对象 |
| `save_workbook` | 保存Excel工作簿 | `workbook`: 工作簿对象<br>`file_path`: 保存路径 | 成功/失败 |
| `convert_formulas_to_values` | 将公式转换为值 | `workbook`: 工作簿对象<br>`sheet_name`: 工作表名称 | 成功/失败 |
| `split_merged_cells` | 拆分合并单元格 | `workbook`: 工作簿对象<br>`sheet_name`: 工作表名称<br>`keep_value`: 是否保留值 | 成功/失败 |
| `merge_cells` | 合并单元格 | `workbook`: 工作簿对象<br>`sheet_name`: 工作表名称<br>`range_str`: 合并范围 | 成功/失败 |
| `add_worksheet` | 添加工作表 | `workbook`: 工作簿对象<br>`sheet_name`: 工作表名称 | 成功/失败 |
| `delete_worksheet` | 删除工作表 | `workbook`: 工作簿对象<br>`sheet_name`: 工作表名称 | 成功/失败 |
| `insert_rows` | 插入行 | `workbook`: 工作簿对象<br>`sheet_name`: 工作表名称<br>`row_idx`: 行索引<br>`amount`: 插入数量 | 成功/失败 |
| `delete_rows` | 删除行 | `workbook`: 工作簿对象<br>`sheet_name`: 工作表名称<br>`row_idx`: 行索引<br>`amount`: 删除数量 | 成功/失败 |
| `insert_columns` | 插入列 | `workbook`: 工作簿对象<br>`sheet_name`: 工作表名称<br>`col_idx`: 列索引<br>`amount`: 插入数量 | 成功/失败 |
| `delete_columns` | 删除列 | `workbook`: 工作簿对象<br>`sheet_name`: 工作表名称<br>`col_idx`: 列索引<br>`amount`: 删除数量 | 成功/失败 |

## 2. 处理模块接口

### processing.py

#### 主要功能
实现后台处理任务，负责在后台线程中执行Excel处理操作，提供进度更新和错误处理机制。

#### 主要方法

| 方法名 | 功能描述 | 参数 | 返回值 |
| ----- | ------- | ---- | ----- |
| `process_files` | 处理多个Excel文件 | `file_list`: 文件列表<br>`steps`: 操作步骤列表<br>`progress_callback`: 进度回调函数 | 处理结果列表 |
| `process_single_file` | 处理单个Excel文件 | `file_path`: 文件路径<br>`steps`: 操作步骤列表<br>`progress_callback`: 进度回调函数 | 处理结果 |

## 3. 执行模块接口

### execution.py

#### 主要功能
执行操作的实现，根据操作类型调用相应的核心功能。

#### 主要方法

| 方法名 | 功能描述 | 参数 | 返回值 |
| ----- | ------- | ---- | ----- |
| `execute_step` | 执行单个操作步骤 | `workbook`: 工作簿对象<br>`step`: 操作步骤<br>`step_index`: 步骤索引 | 执行结果 |
| `execute_formula_to_value` | 执行公式转值操作 | `workbook`: 工作簿对象<br>`params`: 操作参数 | 执行结果 |
| `execute_merge_cells` | 执行合并单元格操作 | `workbook`: 工作簿对象<br>`params`: 操作参数 | 执行结果 |
| `execute_split_merged_cells` | 执行拆分合并单元格操作 | `workbook`: 工作簿对象<br>`params`: 操作参数 | 执行结果 |
| `execute_add_worksheet` | 执行添加工作表操作 | `workbook`: 工作簿对象<br>`params`: 操作参数 | 执行结果 |
| `execute_delete_worksheet` | 执行删除工作表操作 | `workbook`: 工作簿对象<br>`params`: 操作参数 | 执行结果 |
| `execute_insert_rows` | 执行插入行操作 | `workbook`: 工作簿对象<br>`params`: 操作参数 | 执行结果 |
| `execute_delete_rows` | 执行删除行操作 | `workbook`: 工作簿对象<br>`params`: 操作参数 | 执行结果 |
| `execute_insert_columns` | 执行插入列操作 | `workbook`: 工作簿对象<br>`params`: 操作参数 | 执行结果 |
| `execute_delete_columns` | 执行删除列操作 | `workbook`: 工作簿对象<br>`params`: 操作参数 | 执行结果 |

## 4. 消息处理模块接口

### message_utils.py

#### 主要功能
提供处理执行结果消息的通用函数，格式化执行结果消息，移除重复信息，统一显示格式。

#### 主要方法

| 方法名 | 功能描述 | 参数 | 返回值 |
| ----- | ------- | ---- | ----- |
| `format_result_message` | 格式化执行结果消息 | `result`: 包含step、success和message字段的结果字典 | 格式化后的消息 |

## 5. 报告模块接口

### report.py

#### 主要功能
生成执行报告，记录操作步骤和执行结果。

#### 主要方法

| 方法名 | 功能描述 | 参数 | 返回值 |
| ----- | ------- | ---- | ----- |
| `generate_report` | 生成执行报告 | `results`: 执行结果列表<br>`output_path`: 输出路径 | 报告文件路径 |

## 6. 模型模块接口

### models.py

#### 主要功能
定义了程序中使用的数据模型，包括操作步骤的数据结构和相关配置信息的模型定义。

#### 主要数据结构

| 数据结构 | 描述 | 字段 |
| ------- | ---- | ---- |
| `Step` | 操作步骤 | `type`: 操作类型<br>`params`: 操作参数<br>`description`: 步骤描述 |
| `StepResult` | 步骤执行结果 | `step`: 步骤索引<br>`success`: 是否成功<br>`message`: 结果消息 |

## 7. 前端界面模块

### ui/main_window.py

#### 主要功能
实现了应用程序的主窗口界面，包含文件选择面板、操作步骤列表和功能选项卡。

#### 主要方法

| 方法名 | 功能描述 | 参数 | 返回值 |
| ----- | ------- | ---- | ----- |
| `add_files` | 添加文件到处理列表 | - | - |
| `add_folder` | 添加文件夹中的Excel文件 | - | - |
| `remove_selected_files` | 移除选中的文件 | - | - |
| `clear_files` | 清空文件列表 | - | - |
| `add_step` | 添加操作步骤 | `step`: 操作步骤 | - |
| `remove_step` | 移除操作步骤 | `index`: 步骤索引 | - |
| `move_step_up` | 上移操作步骤 | `index`: 步骤索引 | - |
| `move_step_down` | 下移操作步骤 | `index`: 步骤索引 | - |
| `execute_steps` | 执行所有操作步骤 | - | - |
| `update_progress` | 更新进度显示 | `current`: 当前进度<br>`total`: 总进度<br>`message`: 进度消息 | - |

### ui/file_operations.py

#### 主要功能
实现文件操作界面，包括文件选择和管理。

### ui/step_operations.py

#### 主要功能
实现步骤操作界面，包括步骤添加、编辑、删除和调整顺序。

### ui/worksheet_operations.py

#### 主要功能
实现工作表操作界面，包括添加和删除工作表。

### ui/row_col_operations.py

#### 主要功能
实现行列操作界面，包括插入和删除行列。

## 8. 前后端交互流程

1. **用户界面操作流程**：
   - 用户通过UI界面添加Excel文件
   - 用户配置操作步骤并添加到步骤列表
   - 用户点击执行按钮开始处理

2. **后台处理流程**：
   - UI调用processing.py中的process_files方法
   - processing.py遍历文件列表，对每个文件调用process_single_file方法
   - process_single_file方法加载工作簿，然后遍历步骤列表
   - 对每个步骤，调用execution.py中的execute_step方法
   - execute_step根据步骤类型调用相应的执行方法
   - 执行方法调用core.py中的相应功能实现
   - 执行结果通过message_utils.py格式化后返回
   - 所有文件处理完成后，调用report.py生成执行报告

3. **进度更新流程**：
   - processing.py在处理过程中通过progress_callback回调函数更新进度
   - UI接收进度更新并显示在界面上

4. **错误处理流程**：
   - 执行过程中的错误被捕获并记录在执行结果中
   - UI显示错误信息并允许用户查看详细错误报告

## 9. 数据流向图

```
用户界面 (UI) → 步骤配置 → 处理模块 (processing.py) → 执行模块 (execution.py) → 核心模块 (core.py) → Excel文件
     ↑                                      |                                               |
     |                                      v                                               |
     └─────────── 进度更新 ← 消息处理模块 (message_utils.py) ← 执行结果 ←─────────────────┘
     |                                                                                     |
     └─────────── 执行报告 ← 报告模块 (report.py) ←───────────────────────────────────────┘
```

## 10. 注意事项

1. **错误处理**：所有模块都应该捕获并处理可能的异常，确保程序不会因为单个操作失败而崩溃。

2. **参数验证**：在调用核心功能前，应该验证参数的有效性，避免无效操作。

3. **进度更新**：长时间操作应该提供进度更新，让用户了解处理状态。

4. **资源释放**：确保在操作完成后释放所有资源，特别是Excel工作簿对象。

5. **兼容性**：考虑不同Excel版本的兼容性问题，特别是处理特殊格式和功能时。