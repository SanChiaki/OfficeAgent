# Upload Filter Preview Design

日期：2026-04-29

## 背景

Ribbon Sync 当前上传会先从 Excel 解析出 `CellChange[]`，再生成确认预览，用户确认后把 `preview.Changes` 交给当前 sheet 绑定的业务连接器上传。真实业务系统接入时需要在上传前按业务规则过滤部分单元格，并让用户在确认前看到实际上传数量、跳过数量和跳过原因。

## 设计

新增一个可选上传过滤扩展点，由具体业务系统连接器实现。上传准备阶段生成原始 `CellChange[]` 后，`WorksheetSyncExecutionService` 通过 `WorksheetSyncService` 调用当前连接器的过滤能力：

- 未实现过滤扩展点的连接器：所有 `CellChange` 继续按现有行为上传。
- 实现过滤扩展点的连接器：返回实际上传项和跳过项，跳过项必须携带原因。

`SyncOperationPreview.Changes` 只保存实际上传项，确保确认后提交给 `BatchSave()` 的内容和确认弹窗一致。预览摘要显示“将上传 X 个单元格，跳过 Y 个单元格”，明细显示部分实际上传项和跳过原因。第一版继续使用现有原生确认弹窗，因此跳过明细做数量限制，避免 MessageBox 过长。

## 数据合同

- `IUploadChangeFilter`：业务连接器可选实现，用于过滤上传变更。
- `UploadChangeFilterResult`：包含 `IncludedChanges` 和 `SkippedChanges`。
- `SkippedCellChange`：包含被跳过的 `CellChange` 和 `Reason`。
- `SyncOperationPreview.SkippedChanges`：供确认弹窗、测试和后续自定义明细对话框使用。

## 边界

过滤发生在上传预览前，不在 `BatchSave()` 内静默处理。这样用户看到的实际上传数量和最终提交内容一致。

第一版不做通用规则配置 UI，也不做规则脚本引擎。真实系统可以在连接器内部写代码、读取本地配置，或调用业务接口获取规则。

## 测试

- Core 预览工厂测试：同时展示实际上传数量、跳过数量和跳过原因。
- Excel 上传执行服务测试：过滤后 `preview.Changes` 只包含实际上传项，`ExecuteUpload()` 只提交实际上传项。
- 现有无过滤连接器测试保持现有上传行为。
