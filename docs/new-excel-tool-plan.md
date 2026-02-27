
# Excel Agent excel tool 实现计划

## 一、整体架构

```
libs/
  excel_com/
    __init__.py
    instance_manager.py      # 单例，管理 Excel.Application 与 workbook 注册表
    workbook_handle.py       # WorkbookHandle，封装单个工作簿的状态与 COM 引用
    range_ops.py             # 单元格 / 区域 读写
    formula_ops.py           # 公式读写
    table_ops.py             # ListObject（表格）CRUD
    query_ops.py             # PowerQuery / M Code CRUD + 刷新
    sheet_ops.py             # 工作表 CRUD
    errors.py                # 自定义异常体系
    constants.py             # Excel 枚举常量（xlUp, xlDown, xlCalculationManual…）
    utils.py                 # 路径规范化、COM 安全调用、重试装饰器等

tools/
  excel_tools/
    __init__.py              # ExcelToolProvider(ToolProvider) 入口
    _common.py               # 工具层公共：参数校验、错误格式化、COM 错误映射
    workbook_tool.py         # 打开/保存/关闭/列出工作簿
    read_tool.py             # 批量读取
    write_tool.py            # 批量写入
    formula_tool.py          # 批量公式读写
    table_tool.py            # 表格批量操作
    query_tool.py            # PowerQuery 批量操作
    sheet_tool.py            # 工作表批量操作
```

---

## 二、底层 COM 层：`libs/excel_com`

### 2.1 `instance_manager.py` — ExcelInstanceManager（核心）

这是整个方案的关键组件，采用单例模式，负责所有 Excel COM 生命周期管理。

**内部数据结构：**

```python
@dataclass
class WorkbookEntry:
    file_path: str              # 规范化后的绝对路径（小写）
    workbook: CDispatch         # COM Workbook 对象引用
    was_already_open: bool      # attach 时该文件是否已被用户打开
    is_editing: bool            # agent 是否正在编辑中
    is_saved: bool              # 上次操作后是否已保存
    opened_at: datetime         # 打开时间
    last_operation_at: datetime # 最后操作时间
    undo_checkpoint: bool       # 是否设置了 undo 检查点（save 点）

class ExcelInstanceManager:
    _instance: ClassVar[Optional['ExcelInstanceManager']] = None
    _app: Optional[CDispatch] = None            # Excel.Application
    _registry: Dict[str, WorkbookEntry] = {}    # key = normalized path
```

**关键职责与方法：**

**2.1.1 获取 / 创建 Excel.Application**

`_ensure_app(self) -> CDispatch`

流程如下。首先检查 `_app` 是否存在且仍然有效（通过尝试访问 `_app.Version` 来验证 COM 引用未失效）。如果已失效或为 None，先尝试 `win32com.client.GetActiveObject("Excel.Application")` 来附着到用户已打开的 Excel 进程。如果没有正在运行的 Excel，再用 `win32com.client.DispatchEx("Excel.Application")` 创建新实例。

注意事项：使用 `DispatchEx` 而非 `Dispatch`，因为 `DispatchEx` 总是创建新的进程实例，避免和用户实例冲突——但在这个场景下我们实际上**希望**附着到用户实例，所以优先用 `GetActiveObject`。只有在用户没有打开 Excel 的时候才 `DispatchEx`。获取到 Application 后，记录这个 Application 是附着的还是新建的（`_app_is_attached: bool`），这决定了最终清理时是否应该退出 Application。

**2.1.2 打开工作簿**

`open_workbook(self, file_path: str) -> WorkbookEntry`

流程：先将 `file_path` 规范化（`os.path.realpath` + `os.path.normcase`），然后查 `_registry`。如果已经在注册表中，验证 COM 引用是否仍有效（访问 `wb.Name`），有效则直接返回。如果不在注册表中，遍历 `_app.Workbooks` 集合，逐一比对规范化路径（通过 `wb.FullName`），如果找到则说明用户已经打开了这个文件，创建 `WorkbookEntry` 并标记 `was_already_open=True`。如果遍历后未找到，调用 `_app.Workbooks.Open(file_path)`，创建 `WorkbookEntry` 并标记 `was_already_open=False`。

注意事项：`Workbooks.Open` 时需要传入关键参数 `UpdateLinks=False`（避免弹更新链接对话框）、`ReadOnly=False`。如果打开时出现只读情况（`wb.ReadOnly == True`），需要检查原因——可能是文件被锁。此时尝试关闭再以写入模式重新打开，或向上层报告错误。对于 `was_already_open=True` 的工作簿，如果它处于只读模式，不要尝试强行切换，而是返回错误提示用户先关闭其他占用。

**2.1.3 关闭工作簿**

`close_workbook(self, file_path: str, save: bool = True)`

流程：从 `_registry` 中查找。如果 `was_already_open=True`，只保存不关闭——用户打开的文件不应该被 agent 关闭。如果 `was_already_open=False`，根据 `save` 参数决定是否保存，然后关闭，从 `_registry` 移除。

**2.1.4 保存工作簿**

`save_workbook(self, file_path: str)`

调用 `wb.Save()`，更新 `entry.is_saved = True`。

**2.1.5 注册表健康检查**

`_validate_registry(self)`

遍历 `_registry`，对每个 entry 尝试访问 `wb.Name`，如果抛出 COM 异常说明引用已失效（用户手动关闭了文件），则从注册表中移除。这个方法在每次 `open_workbook` 和 `get_workbook` 调用前执行。

**2.1.6 全局清理**

`cleanup(self)`

关闭所有 `was_already_open=False` 的工作簿（保存），对 `was_already_open=True` 的只保存不关闭。如果 `_app_is_attached=False`（Application 是我们创建的），且关闭完所有我们打开的工作簿后 `_app.Workbooks.Count == 0`，则调用 `_app.Quit()`。恢复所有被我们修改的 Application 级设置。

**2.1.7 状态查询（供状态栏使用）**

`get_all_status(self) -> List[WorkbookStatus]`

返回所有注册工作簿的状态信息，供 UI 状态栏展示。

**2.1.8 性能优化上下文管理器**

`batch_operation(self, file_path: str)` — 作为 context manager

进入时：`ScreenUpdating = False`，`EnableEvents = False`，`Calculation = xlCalculationManual`，`DisplayAlerts = False`。退出时恢复所有设置，并且 `Calculation = xlCalculationAutomatic` 会触发重新计算。特别注意：这些都是 Application 级别的设置，如果用户自己也在看 Excel，`ScreenUpdating=False` 会影响用户体验。因此对于 `was_already_open=True` 的工作簿，可以考虑只关闭 `Calculation` 和 `EnableEvents`，保留 `ScreenUpdating=True`，确保用户无感。

---

### 2.2 `workbook_handle.py` — WorkbookHandle

对外暴露的操作句柄，封装单个工作簿的常用快捷访问。

提供 `get_sheet(name_or_index)` 方法，带验证。提供 `active_sheet` 属性。提供 `sheet_names` 属性。代理 `save()`、`close()` 到 InstanceManager。

---

### 2.3 `range_ops.py` — 区域读写

**读取：**

`read_range(wb, sheet, range_address) -> List[List[Any]]`

对于单个单元格返回 `[[value]]` 保持统一格式。对于大区域（超过 `10000` 单元格），COM 的 `Range.Value` 返回的是嵌套 tuple，直接转 list 即可，性能没有问题。注意处理 `None`（空单元格）、`datetime`（COM 返回 `pywintypes.datetime`，需转换）、`ERROR` 值（COM 返回的错误整数如 `-2146826281` 对应 `#REF!`，需映射为可读字符串）。

`read_used_range(wb, sheet) -> Tuple[str, List[List[Any]]]`

返回 UsedRange 的地址和全部值。

**写入：**

`write_range(wb, sheet, start_cell, values: List[List[Any]])`

通过计算 values 的行列数确定目标区域大小，然后用 `Range.Value = values` 一次性写入（而非逐单元格，性能差异巨大）。注意 `values` 必须是二维列表，即使只有一行也要 `[[v1, v2, ...]]`。写入后更新 `entry.is_saved = False`。

**自动调整列宽：**

`auto_fit_columns(wb, sheet, range_address=None)`

如果不传 range，则对 `UsedRange.Columns.AutoFit()` 进行调用。

---

### 2.4 `formula_ops.py` — 公式读写

`get_formula(wb, sheet, range_address) -> List[List[str]]`

通过 `Range.Formula` 读取。如果单元格没有公式则返回空字符串。注意 `Range.Formula` 返回的是英文公式（`=SUM(A1:A10)`），`Range.FormulaLocal` 返回本地化公式。统一使用 `Formula`（英文版），因为 LLM 生成的公式通常是英文的。

`set_formula(wb, sheet, range_address, formula: str)`

通过 `Range.Formula = formula` 设置。如果是要对一整列设置同一公式，使用 `Range.Formula = formula` 设置首个单元格后，可用 `Range.AutoFill` 或直接对整个区域赋值。

注意事项：LLM 可能生成语法错误的公式，COM 会抛出异常。需要捕获并返回有意义的错误信息给 LLM，例如"公式语法错误"，而不是原始的 COM 错误码。

---

### 2.5 `table_ops.py` — 表格（ListObject）操作

`create_table(wb, sheet, range_address, name, has_headers=True)`

通过 `sheet.ListObjects.Add(SourceType=xlSrcRange, Source=range, XlListObjectHasHeaders=xlYes)`。创建后设置 `table.Name = name`。注意：表名在工作簿内必须唯一，需要先检查是否重名。创建表后需要等待 COM 完成操作——从你的日志来看，`table_create` 耗时 71 秒，说明可能存在性能问题或者 COM 在做大量内部处理。建议在创建表后加一个 `DoEvents` 等效操作（可通过 `pythoncom.PumpWaitingMessages()` 实现）。

`delete_table(wb, sheet, table_name)`

`add_column(wb, sheet, table_name, column_name, formula=None, position=None)`

通过 `table.ListColumns.Add(Position=pos)` 添加列。如果需要设置公式，对 `col.DataBodyRange.Formula = formula` 赋值。从你的日志来看，这一步报错了。常见原因：公式引用不正确（表格结构引用 `[@Column1]` 的列名不匹配）、表格还在更新中。解决方案：添加列和设置公式分成两步，中间加 `PumpWaitingMessages()`；如果公式设置失败，改用逐单元格设置。

`list_tables(wb, sheet=None) -> List[TableInfo]`

`resize_table(wb, sheet, table_name, new_range)`

`set_table_style(wb, sheet, table_name, style_name)`

---

### 2.6 `query_ops.py` — PowerQuery / M Code 操作

这是最复杂的部分。Excel COM 通过 `Workbook.Queries` 集合操作 PowerQuery。

`list_queries(wb) -> List[QueryInfo]`

遍历 `wb.Queries`，返回每个查询的 Name、Formula（M Code）、Description。

`create_query(wb, name, m_code, description="")`

通过 `wb.Queries.Add(Name=name, Formula=m_code, Description=description)`。注意：这只是创建了查询定义，还没有加载到工作表。如果需要加载结果到 sheet，还需要创建 Connection 和 QueryTable。

`update_query(wb, name, m_code)`

通过找到对应 Query 对象，设置 `query.Formula = m_code`。

`delete_query(wb, name)`

通过 `query.Delete()`。注意同时需要清理关联的 Connection（`wb.Connections`）和 QueryTable。

`refresh_query(wb, name, wait=True)`

找到对应的 Connection，调用 `conn.Refresh()`。如果 `wait=True`，需要设置 `conn.OLEDBConnection.BackgroundQuery = False` 以同步刷新，否则 COM 会立刻返回而刷新在后台进行。同步刷新可能耗时很长（取决于数据源），需要设置合理的超时。

`load_query_to_sheet(wb, query_name, sheet_name, start_cell)`

这是关键功能。创建查询后需要将结果加载到工作表。实现方式：通过 `wb.Connections` 找到或创建 Power Query 关联的连接，然后在目标 sheet 上创建 `QueryTable` 并绑定到该连接。或者使用更简单的方式：通过 `ListObjects.Add(SourceType=xlSrcQuery, Source=connection_string)` 来创建一个绑定了 PQ 的表。

注意事项：PowerQuery 的 COM 支持在 Excel 2016+ 才比较完整。`Workbook.Queries` 集合在较老版本中不存在。M Code 中的换行符处理需要注意——LLM 生成的 M Code 可能用 `\n`，COM 接口需要真实的换行符。M Code 语法错误会在刷新时才报错（而非设置 Formula 时），所以需要在刷新后检查连接状态。

---

### 2.7 `sheet_ops.py` — 工作表操作

`list_sheets(wb) -> List[str]`

`add_sheet(wb, name, position=None) -> str`

`delete_sheet(wb, name)`

注意删除时需要 `DisplayAlerts=False`，否则会弹确认对话框。

`rename_sheet(wb, old_name, new_name)`

`copy_sheet(wb, source_name, new_name, position=None)`

`get_sheet_info(wb, name) -> SheetInfo`

返回 UsedRange 地址、行数、列数等概要信息。

---

### 2.8 `errors.py` — 异常体系

```
ExcelComError (base)
├── ExcelInstanceError          # Application 获取/创建失败
├── WorkbookNotFoundError       # 注册表中找不到该文件
├── WorkbookReadOnlyError       # 文件以只读模式打开
├── WorkbookLockedError         # 文件被其他进程锁定
├── SheetNotFoundError          # 工作表不存在
├── RangeError                  # 区域地址无效
├── FormulaError                # 公式语法错误
├── TableError                  # 表格操作错误（重名、不存在等）
├── QueryError                  # PowerQuery 操作错误
├── ComCallError                # 底层 COM 调用失败（包装原始 com_error）
└── StaleReferenceError         # COM 引用已失效
```

每个异常都包含 `user_message`（给 LLM/用户看的清晰描述）和 `raw_error`（原始 COM 错误，用于调试日志）。

---

### 2.9 `utils.py` — 工具函数

**路径规范化：** `normalize_path(path: str) -> str` — `os.path.realpath` + `os.path.normcase` + 处理中文路径。

**COM 安全调用装饰器：** `com_retry(max_retries=2, delay=0.5)` — 捕获瞬态 COM 错误（如 RPC 超时、`-2147418111 Call was rejected by callee`），自动重试。你的日志里已经有类似的重试逻辑，需要继续保留。区分瞬态错误（可重试）和逻辑错误（不可重试）。

**COM 错误码映射：** 维护一个 dict 将常见的 HRESULT 错误码映射为可读的中文错误信息。比如 `-2146827284` 是 `TYPE_MISMATCH`（类型不匹配），`-2147024809` 是 `E_INVALIDARG`（参数无效）。

**消息泵：** `pump_messages()` — 调用 `pythoncom.PumpWaitingMessages()`，在长操作之间调用，防止 COM 通道阻塞。

---

## 三、工具层：`tools/excel_tools`

### 3.1 设计原则

每个 tool 接受一个操作列表，实现单次调用和批量调用的兼顾。统一的参数模式为 `file_path`（必填）+ `operations: List[OperationSchema]`。每个 tool 的 description 需要对 LLM 友好，包含明确的参数说明、示例和限制条件。

所有 tool 内部统一使用以下流程：获取或打开工作簿 → 进入 batch context → 遍历 operations 逐一执行 → 如果某一步失败则中断并触发 undo → 退出 batch context → 返回结果摘要。

### 3.2 Undo / 回滚策略

你要求使用 Excel Undo 机制，但需要注意 COM 操作的 Undo 可靠性有限。以下是实际可行的混合策略。

方案：在每次 tool 调用开始前，先调用 `wb.Save()` 作为检查点（checkpoint）。如果工作簿有未保存的用户更改且 `was_already_open=True`，则先保存一次（保护用户数据）。如果操作过程中失败，首先尝试 `Application.Undo()` 来回退（循环调用，直到回退到 checkpoint 状态或 Undo 不再有效）。如果 Undo 不够可靠（某些操作不生成 undo 记录），则关闭工作簿不保存（`wb.Close(SaveChanges=False)`），然后重新打开——这样文件就恢复到了 checkpoint 时的状态。

`WorkbookEntry` 中用一个 `undo_checkpoint_saved: bool` 字段标记是否已经做过 checkpoint save。

对于用户已打开的文件（`was_already_open=True`），关闭重新打开的回退方案是不可接受的（会影响用户），所以对这类文件只能依赖 Undo，如果 Undo 也无法恢复就保留当前状态并告知 LLM 哪些操作成功了、哪些失败了。

### 3.3 各 Tool 定义

**3.3.1 `workbook_tool.py` — WorkbookTool**

名称：`excel_workbook`

功能：打开、保存、关闭工作簿，获取工作簿信息。

operations 类型：`open`（参数：file_path）、`save`（参数：file_path）、`close`（参数：file_path, save=True）、`info`（参数：file_path，返回 sheet 列表、已有表格、已有查询等概要）、`list`（返回当前所有已打开工作簿状态）。

注意：这是所有其他 tool 的前置。但其他 tool 应该也能自动 open（如果文件未打开则自动打开），不强制用户先调用 open。

---

**3.3.2 `read_tool.py` — ReadTool**

名称：`excel_read`

operations schema：

```
{
  "file_path": "...",
  "operations": [
    {
      "type": "range",          // range | used_range | sheet_info
      "sheet": "Sheet1",        // sheet 名或索引
      "range": "A1:D10"         // type=range 时必填
    }
  ]
}
```

返回格式：每个 operation 返回 `{"status": "ok", "data": [[...]], "range": "A1:D10"}` 或 `{"status": "error", "message": "..."}`。

注意：返回数据如果过大（比如几千行），需要截断并告知 LLM "返回了前 100 行，共 5000 行"，避免 token 爆炸。设置一个 `MAX_RETURN_ROWS` 常量（建议 50-100 行），超过的告知总数和建议分批读取。

---

**3.3.3 `write_tool.py` — WriteTool**

名称：`excel_write`

operations schema：

```
{
  "file_path": "...",
  "auto_fit": true,          // 写完后是否自动调整列宽
  "save": true,              // 写完后是否保存
  "operations": [
    {
      "sheet": "Sheet1",
      "start_cell": "A1",
      "values": [["Name", "Age"], ["Alice", 30]]
    }
  ]
}
```

---

**3.3.4 `formula_tool.py` — FormulaTool**

名称：`excel_formula`

operations schema：

```
{
  "file_path": "...",
  "operations": [
    {
      "type": "set",             // set | get
      "sheet": "Sheet1",
      "range": "F2:F100",
      "formula": "=D2*E2"       // type=set 时必填
    }
  ]
}
```

注意：当 `range` 是一个区域（如 `F2:F100`）且 `formula` 只给了一个公式时，Excel 会自动调整行引用。告诉 LLM 这一行为。如果 LLM 设置的公式无效，捕获 COM 异常后给出清晰提示，如"公式 '=SUMX(...)' 在单元格 F2 设置失败：未知函数名"。

---

**3.3.5 `table_tool.py` — TableTool**

名称：`excel_table`

operations schema：

```
{
  "file_path": "...",
  "operations": [
    {
      "type": "create",          // create | delete | add_column | remove_column | list | set_style | resize | info
      "sheet": "Sheet1",
      "range": "A1:D10",        // create 时必填
      "table_name": "SalesData",
      "has_headers": true,
      // add_column 时：
      "column_name": "Total",
      "formula": "=[@Price]*[@Qty]",   // 可选，结构化引用
      "position": null
    }
  ]
}
```

注意：批量操作 table 时，`create` 必须在 `add_column` 之前。工具层需要验证 operations 的顺序合理性，或者在执行每一步前重新获取 table 引用。`add_column` 设置公式失败时（你日志中的问题），改为分步执行：先 `ListColumns.Add` 创建空列，`PumpWaitingMessages()`，然后再设置 `Formula`，如果仍失败则降级为逐单元格设置。

---

**3.3.6 `query_tool.py` — QueryTool**

名称：`excel_query`

operations schema：

```
{
  "file_path": "...",
  "operations": [
    {
      "type": "create",         // create | update | delete | refresh | list | load_to_sheet | get_mcode
      "query_name": "CleanData",
      "m_code": "let\n  Source = ...\nin\n  Result",
      "description": "",
      // load_to_sheet 时：
      "target_sheet": "Results",
      "target_cell": "A1",
      // refresh 时：
      "wait": true,
      "timeout_seconds": 60
    }
  ]
}
```

注意：M Code 中的换行、缩进需要正确传递，LLM 在 JSON 中需要用 `\n`。`create` 和 `load_to_sheet` 通常配合使用，建议在 tool description 中说明这一点。`refresh` 同步等待时需要超时机制——在循环中检查连接状态，超时后抛出异常。刷新失败的错误信息在 `conn.OLEDBConnection.LastRefreshInfo` 中，取出后返回给 LLM。

---

**3.3.7 `sheet_tool.py` — SheetTool**

名称：`excel_sheet`

operations schema：

```
{
  "file_path": "...",
  "operations": [
    {
      "type": "add",            // add | delete | rename | copy | list
      "sheet_name": "NewSheet",
      "new_name": "...",         // rename 时用
      "position": 1              // add/copy 时可选
    }
  ]
}
```

---

### 3.4 `_common.py` — 工具层公共逻辑

**统一错误处理装饰器**

所有 tool 的 `_run` 方法包裹一个装饰器，负责：将 `ExcelComError` 子类映射为结构化错误字符串返回给 LLM（而非抛出异常中断 agent）。在错误信息中包含建议操作，比如"表 'SalesData' 不存在，可用表有：Table1, Table2"。捕获未预期的异常，记录完整 traceback 到日志，给 LLM 返回简化的错误信息。

**参数预处理**

`file_path` 自动规范化。`sheet` 参数如果未传，默认使用 active sheet。`range` 参数支持 `"A1"` 和 `"A1:D10"` 两种格式。

**返回值格式统一**

每个 tool 返回 JSON 字符串，格式为：

```json
{
  "success": true,
  "results": [
    {"operation_index": 0, "status": "ok", "message": "..."},
    {"operation_index": 1, "status": "error", "message": "..."}
  ],
  "file_saved": true
}
```

如果某个 operation 失败导致回滚，后续 operations 标记为 `"status": "skipped"`。

---

### 3.5 `__init__.py` — ExcelToolProvider

```python
class ExcelToolProvider:
    def __init__(self, instance_manager: ExcelInstanceManager):
        self._manager = instance_manager

    def get_tools(self) -> list[BaseTool]:
        return [
            WorkbookTool(self._manager),
            ReadTool(self._manager),
            WriteTool(self._manager),
            FormulaTool(self._manager),
            TableTool(self._manager),
            QueryTool(self._manager),
            SheetTool(self._manager),
        ]
```

---

## 四、需要特别注意的细节

### 4.1 COM 引用失效问题

用户可能在 agent 操作期间手动关闭了 Excel 或关闭了某个工作簿。任何 COM 调用都可能因此抛出异常。所有底层方法的入口都应该先验证引用有效性。如果检测到 Application 被关闭，需要清空整个 `_registry` 并重新获取/创建 Application。

### 4.2 文件路径一致性

Windows 路径大小写不敏感，但 Python 字符串比较是大小写敏感的。LLM 每次传入的路径格式可能不同（正斜杠/反斜杠、有无尾部斜杠、相对路径/绝对路径）。必须在入口处统一规范化，`_registry` 的 key 使用规范化后的路径。

### 4.3 Excel DisplayAlerts 和 ScreenUpdating 的恢复

这两个是 Application 级设置，不是 Workbook 级的。如果 agent 操作中途崩溃（Python 异常），可能导致这些设置没有恢复。必须用 `try/finally` 确保恢复，并且 `ExcelInstanceManager.cleanup()` 中也要兜底恢复。

### 4.4 中文兼容性

Sheet 名、Table 名、文件路径都可能包含中文。COM 接口本身支持 Unicode，但日志输出、JSON 序列化时需要确保编码正确。M Code 中的中文字段名也需要正确处理。

### 4.5 Tool Description 编写

这是影响 agent 效果的关键。每个 tool 的 description 应该包含：工具功能的一句话总结、parameters 的 JSON schema 及每个字段的含义、至少一个完整的调用示例（包括输入和输出）、常见错误和处理方式、和其他 tool 的关系（如"使用此 tool 前需要确保文件已打开，或传入 file_path 自动打开"）。

description 不宜过长（会消耗 context），重点突出 LLM 容易犯错的地方，比如 values 必须是二维数组、公式使用英文函数名、M Code 中的换行用 `\n`。

### 4.6 Large Data 处理

LLM 的 context window 有限。读取大区域时必须截断。写入大量数据时，LLM 不可能在一次 tool call 中写入几千行，所以大数据写入场景应该引导使用 PowerQuery 从外部数据源导入，而非通过 `excel_write`。

### 4.7 `table_add_column` 的公式设置时机

从你的日志分析，`table_add_column` 中 `col.DataBodyRange.Formula = formula` 失败的根本原因大概率是：添加列后 table 的内部结构尚未完全更新，立即设置公式导致 COM 错误。解决方案是在 `ListColumns.Add` 之后插入 `PumpWaitingMessages()` 和短暂 `time.sleep(0.3)`，然后重新获取 column 引用（不要使用 `Add` 返回的引用，重新从 `table.ListColumns(name)` 获取），再设置公式。如果仍然失败，降级为对 `DataBodyRange` 的各个单元格逐一设置。

### 4.8 Slicer 创建问题

你的日志中 `SlicerCaches.Add2` 失败（`E_INVALIDARG`），可能原因是 `source_name` 参数格式不正确。`Add2` 的 `Source` 参数应该是 table 的 `Name` 属性而非 sheet 上的 range 地址，`SourceField` 应该是列名字符串或索引号。需要在工具层做参数校验，在调用前确认 table 和 column 都存在。

### 4.9 Agent 操作期间的互斥

虽然你说会在 agent 操作时禁用状态栏按钮，但还是建议在 `ExcelInstanceManager` 中加一个简单的 `_is_agent_operating: bool` 标志。状态栏的保存/关闭按钮在调用前检查这个标志，双重保险。同时 `WorkbookEntry.is_editing` 也应该在 tool 执行开始时设为 True，结束时设为 False，供状态栏展示。

### 4.10 进程残留清理

如果 Python 进程异常退出（crash），可能留下幽灵 Excel 进程。建议在应用启动时检查是否有上次残留的 Excel 进程（可以通过在某个临时文件中记录 PID），并尝试清理。但这个要谨慎——不要误杀用户自己的 Excel。