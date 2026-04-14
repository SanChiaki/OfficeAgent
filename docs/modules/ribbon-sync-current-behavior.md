# Ribbon Sync Current Behavior

日期：2026-04-14

状态：已实现并可联调。当前只接入 `current-business-system`，但内部继续保留 `systemKey` 作为未来多系统扩展边界。

## 1. 模块范围

Ribbon Sync 是独立于 Agent / task pane 的 Excel Ribbon 数据同步能力。

当前只包含四个动作：

- 全量下载
- 部分下载
- 全量上传
- 部分上传

当前不包含：

- 增量上传
- 本地快照差异比对
- `SheetSnapshots` 元数据表

所有确认、告警、结果反馈都通过 Office / WinForms 原生弹框完成，不走任务窗格。

## 2. Ribbon 入口

当前 Ribbon 分为三组：

- 项目
  - 项目下拉框
  - `初始化当前表`
- 下载
  - `全量下载`
  - `部分下载`
- 上传
  - `全量上传`
  - `部分上传`

主入口代码：

- [src/OfficeAgent.ExcelAddIn/AgentRibbon.cs](../../src/OfficeAgent.ExcelAddIn/AgentRibbon.cs)
- [src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs](../../src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs)

## 3. 元数据模型

运行时元数据都保存在可见工作表 `_OfficeAgentMetadata` 中。

当前使用两张表：

- `SheetBindings`
- `SheetFieldMappings`

当前 `_OfficeAgentMetadata` 的展示布局是单个 sheet 内上下两个可读区域：

- 上半区是 `SheetBindings`
- 下半区是 `SheetFieldMappings`
- 每个区域都包含：
  - 一行标题
  - 一行表头
  - 多行数据
- 两个区域之间固定留两行空白分隔

当前不会再使用旧的“首列表名 + 每行一条压平记录”格式；一旦发生 metadata 写入，插件会按上述可读布局整表重写 `_OfficeAgentMetadata`。

### 3.1 SheetBindings

当前列固定为：

- `SheetName`
- `SystemKey`
- `ProjectId`
- `ProjectName`
- `HeaderStartRow`
- `HeaderRowCount`
- `DataStartRow`

含义：

- `HeaderStartRow`
  - 表头起始行
  - 默认 `1`
- `HeaderRowCount`
  - 表头行数
  - 默认 `2`
- `DataStartRow`
  - 数据区起始行
  - 默认 `3`

### 3.2 SheetFieldMappings

`SheetFieldMappings` 的列结构不写死在 Excel 层，实际列由连接器返回的 `FieldMappingTableDefinition` 决定。

当前系统的典型结构示意：

| SheetName | HeaderId | HeaderType | ApiFieldKey | IsIdColumn | DefaultSingleDisplayName | CurrentSingleDisplayName | DefaultParentDisplayName | CurrentParentDisplayName | DefaultChildDisplayName | CurrentChildDisplayName | ActivityId | PropertyId |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Sheet1 | row_id | single | row_id | true | ID | ID |  |  |  |  |  |  |
| Sheet1 | owner_name | single | owner_name | false | 负责人 | 负责人 |  |  |  |  |  |  |
| Sheet1 | start_12345678 | activityProperty | start_12345678 | false |  |  | 测试活动111 | 测试活动111 | 开始时间 | 开始时间 | 12345678 | start |

说明：

- 第一列固定是 `SheetName`
- 其余列来自业务系统连接器
- Excel 运行时按“语义角色”读取映射，不依赖写死的列顺序
- 当前不会持久化 Excel 列号；每次上传/下载都会重新按当前表头文本识别列

元数据读写代码：

- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs)
- [src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs](../../src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs)
- [src/OfficeAgent.ExcelAddIn/Excel/MetadataSheetLayoutSerializer.cs](../../src/OfficeAgent.ExcelAddIn/Excel/MetadataSheetLayoutSerializer.cs)

## 4. 项目选择与初始化

### 4.1 项目选择

当前行为：

- 用户先通过项目下拉框选择项目
- 选择结果会写入 `SheetBindings`
- 如果当前 sheet 已有绑定，切换回来时下拉框会自动回填
- 如果当前 sheet 没有绑定，下拉框显示 `先选择项目`

一个重要细节：

- 重新选择项目时会保留当前 sheet 已维护的 `HeaderStartRow`、`HeaderRowCount`、`DataStartRow`
- 选择项目不会激活 `_OfficeAgentMetadata`

### 4.2 自动尝试初始化 + 显式初始化兜底

选择项目后，插件会自动尝试初始化当前 sheet：

- 如果当前 sheet 没有绑定
- 或绑定项目变了
- 或 `SheetFieldMappings` 为空

则自动重新初始化。

如果自动初始化失败，控制器会弹出警告，引导用户使用 `初始化当前表`。

`初始化当前表` 的职责只有两件事：

- 写入 / 刷新 `SheetBindings`
- 写入 / 刷新 `SheetFieldMappings`

它不会改动业务单元格。

执行入口：

- [src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs](../../src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs)
- [src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs](../../src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs)
- [src/OfficeAgent.Core/Sync/WorksheetSyncService.cs](../../src/OfficeAgent.Core/Sync/WorksheetSyncService.cs)

## 5. 表头布局与列识别

当前支持同一项目内同时存在：

- 单层表头列
- 双层活动表头列

### 5.1 HeaderRowCount = 1

当 `HeaderRowCount = 1` 时：

- 所有列都只写一行表头
- 活动属性列只显示当前子表头名

### 5.2 HeaderRowCount = 2

当 `HeaderRowCount = 2` 时：

- 单层列会占两行并做纵向合并
- 活动列按活动名在第一行横向合并
- 第二行写活动属性名

### 5.3 运行时匹配规则

上传和下载都会基于当前工作表文本重新识别列：

- 不依赖持久化列号
- 允许用户手工增删改列
- 允许用户手工修改显示列名，只要同步维护 `SheetFieldMappings`

当前匹配规则：

- ID 列允许不在用户选区内
- 表头行允许不在用户选区内
- 运行时会根据 `HeaderStartRow` 和 `HeaderRowCount` 去当前表头区识别列
- 双层表头只在前两层里识别：顶层主表头 + 第二层子表头

关键代码：

- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs)
- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs)

## 6. 下载行为

### 6.1 全量下载

当前流程：

1. 读取 `SheetBindings` 和 `SheetFieldMappings`
2. 尝试按当前表头文本识别运行时列
3. 如果识别成功，只刷新受管数据列的数据区
4. 如果表头区为空，则按 `SheetFieldMappings` 渲染表头，再写数据
5. 如果表头区已有文本但无法匹配映射，则报错，要求先修正表头或元数据

当前不会重写“已识别成功”的现有表头。

这允许用户在表头上方或表头与数据区之间插入统计行，只要 `SheetBindings` 配置正确即可。

### 6.2 部分下载

当前流程：

1. 读取当前可见选区
2. 结合运行时匹配到的列，解析出目标 `rowId + fieldKey`
3. 调用 `/find`
4. 仅把查回值回写到原目标单元格

当前选区规则：

- 仅可见单元格优先
- 支持非连续选区
- 选区可不包含 ID 列
- 选区可不包含表头行

关键代码：

- [src/OfficeAgent.ExcelAddIn/Excel/ExcelVisibleSelectionReader.cs](../../src/OfficeAgent.ExcelAddIn/Excel/ExcelVisibleSelectionReader.cs)
- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetSelectionResolver.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetSelectionResolver.cs)
- [src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs](../../src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs)

## 7. 上传行为

### 7.1 全量上传

当前流程：

1. 从 `DataStartRow` 开始扫描工作表
2. 只处理有 ID 的行
3. 对每个非 ID 列生成一个 `CellChange`
4. 把所有 `CellChange` 发给 `BatchSave`

### 7.2 部分上传

当前流程：

1. 解析当前可见选区
2. 自动回找每个目标单元格所在行的 ID
3. 自动回找该列对应的 `ApiFieldKey`
4. 每个单元格生成一个 `CellChange`
5. 调用 `BatchSave`

### 7.3 当前边界

当前不支持：

- 增量上传
- 本地脏数据检测
- 无 ID 新增行上传
- 删除行同步
- 服务端并发冲突判断

## 8. 当前业务系统合同

当前系统通过 `ISystemConnector` 抽象接入：

- [src/OfficeAgent.Core/Services/ISystemConnector.cs](../../src/OfficeAgent.Core/Services/ISystemConnector.cs)

当前实现：

- [src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs](../../src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs)
- [src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs](../../src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs)

当前 mock 契约：

- `POST /head`
  - 返回 `headList`
  - 包含所有非活动列
  - 活动只返回活动头，不返回活动属性列
- `POST /find`
  - 全量下载和部分下载共用
  - `ids` 为空表示全量
  - `fieldKeys` 为空表示返回整行
  - 行数据是平铺 JSON
- `POST /batchSave`
  - 全量上传和部分上传共用
  - 请求体是按单元格变更组成的 list

当前约定的唯一 ID 字段是 `row_id`。

当前 mock 文档：

- [tests/mock-server/README.md](../../tests/mock-server/README.md)

## 9. 主要代码入口

如果后续继续迭代 Ribbon Sync，建议优先看：

- 入口与交互
  - [src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs](../../src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs)
- 执行编排
  - [src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs](../../src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs)
- 初始化与连接器编排
  - [src/OfficeAgent.Core/Sync/WorksheetSyncService.cs](../../src/OfficeAgent.Core/Sync/WorksheetSyncService.cs)
- 元数据持久化
  - [src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs)
- 表头匹配与布局
  - [src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs)
  - [src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs)
- 当前系统接入
  - [src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs](../../src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs)
  - [src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs](../../src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs)

## 10. 主要测试入口

- 元数据存储
  - [tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs](../../tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs)
- 表头匹配
  - [tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderMatcherTests.cs](../../tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderMatcherTests.cs)
- 执行链路
  - [tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs](../../tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs)
- Ribbon 控制器
  - [tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs](../../tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs)
- 当前系统连接器
  - [tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs](../../tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs)
  - [tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessFieldMappingSeedBuilderTests.cs](../../tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessFieldMappingSeedBuilderTests.cs)
- mock 集成链路
  - [tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs](../../tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs)

## 11. 相关文档

- 设计说明
  - [docs/superpowers/specs/2026-04-14-office-agent-ribbon-sync-configurability-design.md](../superpowers/specs/2026-04-14-office-agent-ribbon-sync-configurability-design.md)
- 实施计划
  - [docs/superpowers/plans/2026-04-14-office-agent-metadata-layout-implementation-plan.md](../superpowers/plans/2026-04-14-office-agent-metadata-layout-implementation-plan.md)
- 真实系统接入
  - [docs/ribbon-sync-real-system-integration-guide.md](../ribbon-sync-real-system-integration-guide.md)
- 手工测试
  - [docs/vsto-manual-test-checklist.md](../vsto-manual-test-checklist.md)

## 12. 文档维护约定

如果 Ribbon Sync 的用户可见行为发生变化，至少同步更新：

- 本文第 2 节到第 8 节
- [docs/module-index.md](../module-index.md)
- [docs/vsto-manual-test-checklist.md](../vsto-manual-test-checklist.md)
