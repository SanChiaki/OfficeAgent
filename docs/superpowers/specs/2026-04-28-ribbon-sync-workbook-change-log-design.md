# Ribbon Sync Workbook Change Log Design

日期：2026-04-28

## 背景

Ribbon Sync 目前只在同步结果弹框中反馈上传 / 下载结果，不在工作簿内留下可追溯记录。用户需要一个 workbook 内的 `xISDP_Log` 工作表，用来查看最近由同步动作造成的业务单元格改动。

本设计只覆盖 Ribbon Sync 的同步动作：

- 部分下载、全量下载执行成功后写入 Excel 的业务单元格
- 部分上传、全量上传执行成功后提交给业务系统的业务单元格

普通手工编辑、初始化当前表、模板操作、`ISDP_Setting` 改动、任务窗格 Agent 写入不进入本日志。

## 用户可见行为

插件会在当前 workbook 中维护一个可见工作表 `xISDP_Log`。如果不存在，则第一次写日志时自动创建。

`xISDP_Log` 固定表头为：

| key | 表头 | 修改模式 | 修改值 | 原始值 | 修改时间 |
| --- | --- | --- | --- | --- | --- |

列含义：

- `key`：业务行 ID，即当前 Ribbon Sync 约定的 `row_id`
- `表头`：当前 Excel 表头显示文本；双层表头使用 `父表头/子表头`
- `修改模式`：`上传` 或 `下载`
- `修改值`：同步动作写入 Excel 或提交给业务系统的新值
- `原始值`：下载时为覆盖前 Excel 值；上传时为用户编辑前 Excel 值
- `修改时间`：本机时间，格式为 `yyyy-MM-dd HH:mm:ss`

日志最多保留最近 2000 条记录。追加新记录后，如果超过 2000 条，删除最旧记录，只保留最新 2000 条。

日志不记录：

- ID 列本身
- 无 `row_id` 的行
- 未匹配到受管字段的单元格
- 原始值与修改值完全相同的单元格

## 架构

新增 ExcelAddIn 层日志组件，避免把 workbook / sheet 写入细节泄漏到 Core：

- `WorksheetChangeLogEntry`
  - 表示一条待写入 `xISDP_Log` 的日志记录。
- `IWorksheetChangeLogStore`
  - ExcelAddIn 内部接口，提供 `Append(entries)`。
- `WorksheetChangeLogStore`
  - 使用 ExcelAddIn 层 worksheet adapter 写入 `xISDP_Log`，负责创建表头、追加记录和裁剪到 2000 条。
  - 可通过扩展现有 `IWorksheetGridAdapter`，或新增专用 log sheet adapter 实现。实现必须支持确保工作表存在、读取现有日志区域、重写最新 2000 条。
- `WorksheetPendingEditTracker`
  - ExcelAddIn 内存缓存，记录用户在业务 sheet 上首次改动某个单元格前的 Excel 值。
- `WorksheetSyncExecutionService`
  - 在同步执行成功后生成日志记录并调用日志 store。
- `ThisAddIn.Application_SheetSelectionChange`
  - 在用户编辑前缓存当前选区的 Excel 文本值，作为后续 `SheetChange` 的 before snapshot。
- `ThisAddIn.Application_SheetChange`
  - 对非 `ISDP_Setting`、非 `xISDP_Log` 的业务 sheet，把已变更坐标标记为 pending edit。

Core 仍保持业务系统合同、字段映射、上传下载编排，不直接知道 `xISDP_Log`。

## 数据流

### 下载

1. 用户触发部分下载或全量下载。
2. 执行服务解析当前运行时列、调用业务系统 `/find`。
3. 覆盖 Excel 单元格前，执行服务读取目标单元格当前文本作为 `原始值`。
4. 写入下载值。
5. 对实际变化的非 ID 单元格生成 `下载` 日志。
6. 将日志追加到 `xISDP_Log`，并裁剪到 2000 条。

下载不依赖 pending edit tracker，因为覆盖前的 Excel 值就是本次下载的原始值。

### 上传

1. 用户选中业务 sheet 单元格时，`Application.SheetSelectionChange` 缓存当前选区的 Excel 文本值。
2. 用户在业务 sheet 手工编辑单元格。
3. `Application.SheetChange` 触发，pending edit tracker 对变更坐标记录第一次缓存到的旧 Excel 值。
4. 用户触发部分上传或全量上传。
5. 执行服务按现有逻辑读取当前 Excel 值并生成 `CellChange`。
6. `BatchSave` 成功返回后，执行服务为实际提交且非 ID 的单元格生成 `上传` 日志。
7. `原始值` 来自 pending edit tracker；如果没有可用的修改前值，则该单元格不写上传日志，因为插件无法可靠还原用户编辑前的 Excel 值。
8. 已成功上传的缓存项被清除，避免后续上传重复使用旧原始值。

上传日志只在 `BatchSave` 成功后写入。若上传失败，不写 `上传` 日志，也不清除 pending edit 缓存。

上传日志只记录 pending edit tracker 中确认为实际变化的单元格。全量上传即使会把所有非 ID 字段提交给业务系统，也只为本次会话内已捕获到旧值且当前值确实不同的单元格写日志。

## 表头显示规则

执行服务已经在运行时拥有 `WorksheetColumnBinding`：

- 单层字段：使用 `ChildHeaderText`，为空时回退 `ParentHeaderText`，再回退 `ApiFieldKey`
- 双层字段：使用 `ParentHeaderText/ChildHeaderText`

日志显示当前 Excel 匹配后的表头文本，而不是 `SheetFieldMappings` 的默认 ISDP 表头。

## 错误处理

日志写入失败不阻断上传 / 下载主流程。失败时写入 `%LocalAppData%\OfficeAgent\logs\officeagent.log`，同步动作仍按原行为显示成功或失败。

如果 `xISDP_Log` 被用户手工删除，下一次有日志时重新创建。

如果 `xISDP_Log` 表头被用户改坏，下一次写日志时重写表头行，不尝试迁移异常列。

## 测试

单元测试覆盖：

- 下载日志记录覆盖前 Excel 值、下载后新值、`下载` 模式和双层表头显示。
- 上传日志在 `BatchSave` 成功后记录 pending edit tracker 中的原始值。
- 上传失败不写日志、不清缓存。
- ID 列、无变化单元格、无 ID 行不写日志。
- `xISDP_Log` 超过 2000 条时裁剪最旧记录。
- `Application.SheetSelectionChange` 为后续手工编辑捕获修改前值。
- `Application.SheetChange` 不跟踪 `ISDP_Setting` 和 `xISDP_Log`。

手工验证覆盖：

- 部分下载一个单元格后查看 `xISDP_Log`。
- 修改一个业务单元格并部分上传后查看 `xISDP_Log`。
- 连续写入超过 2000 条后确认只保留最近 2000 条。
