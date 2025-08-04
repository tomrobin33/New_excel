# Excel MCP 分批读取指南

## 问题背景

当处理大型Excel文件时，一次性读取所有数据可能导致"Chunk too big"错误。这是因为：

1. **平台限制**：大模型平台对单次响应大小有限制
2. **内存限制**：大模型处理大量数据时可能超出内存限制
3. **Token限制**：响应数据过大可能超出token限制

## 解决方案：分批读取

### 1. 新增工具函数

#### `get_excel_file_info`
- **用途**：获取Excel文件基本信息，制定读取策略
- **返回**：文件大小、工作表信息、建议批次大小
- **使用时机**：在开始读取前先了解文件结构

#### `read_excel_data_in_batches`
- **用途**：分批读取Excel数据
- **参数**：
  - `batch_size`: 每批读取行数（默认50）
  - `start_row`: 开始行号
  - `end_row`: 结束行号
- **返回**：当前批次数据 + 下一批次信息

#### `preview_excel_data`
- **用途**：快速预览文件内容
- **参数**：`max_preview_rows`, `max_preview_cols`
- **返回**：小规模数据样本

### 2. 推荐使用流程

#### 步骤1：获取文件信息
```python
# 先了解文件结构
file_info = get_excel_file_info(filepath="https://example.com/large_file.xlsx")
```

#### 步骤2：制定读取策略
根据返回的建议批次大小，制定读取计划：
- 小文件（<1000单元格）：批次大小200行
- 中等文件（1000-5000单元格）：批次大小100行
- 大文件（5000-10000单元格）：批次大小50行
- 超大文件（>10000单元格）：批次大小20行

#### 步骤3：分批读取
```python
# 第一批
batch1 = read_excel_data_in_batches(
    filepath="https://example.com/large_file.xlsx",
    batch_size=50,
    start_row=1
)

# 根据返回的next_batch_info继续读取
batch2 = read_excel_data_in_batches(
    filepath="https://example.com/large_file.xlsx",
    batch_size=50,
    start_row=51  # 从上一批结束位置开始
)
```

### 3. 大模型使用建议

#### 对于大模型：
1. **先预览**：使用`preview_excel_data`快速了解文件结构
2. **获取信息**：使用`get_excel_file_info`了解文件规模
3. **制定策略**：根据文件大小选择合适的批次大小
4. **分批处理**：使用`read_excel_data_in_batches`逐批读取
5. **累积结果**：将每批数据合并处理

#### 示例对话流程：
```
用户：请分析这个Excel文件 https://example.com/data.xlsx

大模型：
1. 先调用 get_excel_file_info 了解文件结构
2. 根据文件大小决定批次策略
3. 使用 read_excel_data_in_batches 分批读取
4. 对每批数据进行处理
5. 最后汇总分析结果
```

### 4. 错误处理

#### 如果仍然遇到"Chunk too big"：
1. **减小批次大小**：将batch_size从50改为20或10
2. **使用预览模式**：先用`preview_excel_data`查看结构
3. **分列处理**：只读取必要的列
4. **使用stdio传输**：对于超大文件，建议使用stdio而不是SSE/HTTP

### 5. 性能优化建议

#### 批次大小选择：
- **小文件**：200行/批
- **中等文件**：100行/批  
- **大文件**：50行/批
- **超大文件**：20行/批

#### 内存优化：
- 使用`read_only=True`模式
- 及时关闭工作簿
- 避免同时处理多个大文件

### 6. 最佳实践

1. **总是先获取文件信息**
2. **根据文件大小调整批次大小**
3. **保持批次大小一致**（便于处理）
4. **记录读取进度**（避免重复读取）
5. **处理完一批后立即处理**（避免内存累积）

通过这种分批读取的方式，可以有效避免"Chunk too big"错误，同时保持对大Excel文件的处理能力。 