# OfficeAgent Ribbon Sync 可配置性增强设计说明

日期：2026-04-14

状态：设计已确认，待进入实施计划

## 1. 目标

在现有 Ribbon Sync 基础上增强工作表配置能力，使插件可以同时支持：

- 表头起始行可配置
- 表头行数可配置
- 数据起始行可配置
- 以可见元数据表维护当前 sheet 的字段映射
- 直接接管用户已有的 Excel 表格，而不是只支持插件首次下载生成的表
- 保持当前只接入一个业务系统，同时保留未来多系统扩展能力

本次设计同时降低初版复杂度：

- 移除 `增量上传` Ribbon 按钮和相关功能
- 不新建 `SheetSnapshots` 元数据表
- 不做本地快照差异、脏数据检测和自动冲突判断

## 2. 范围

### 2.1 本次范围

- 调整 Ribbon 入口，增加 `初始化当前表`
- 将 sheet 绑定信息和布局配置合并到同一张元数据表
- 新增按当前 sheet 维护的字段映射表
- 支持 `HeaderStartRow`、`HeaderRowCount`、`DataStartRow`
- 支持单行表头和双行混合表头
- 支持已有 Excel 的自动识别接管
- 支持失败后通过显式初始化动作兜底
- 下载和上传时每次按当前表头文本重新识别列

### 2.2 本次明确不做

- 不做增量上传
- 不做基于快照的本地改动检测
- 不记录列号、列地址、固定 Excel 列位置
- 不自动追踪用户后续对 Excel 表头的改名
- 不自动同步用户后续对列增删改的动作

## 3. Ribbon 交互

Ribbon 保持三个功能区，但调整为：

- `项目`
  - 项目下拉框
  - `初始化当前表`
- `下载`
  - `全量下载`
  - `部分下载`
- `上传`
  - `全量上传`
  - `部分上传`

交互规则：

- 用户先通过 Ribbon 下拉框选择项目
- 下拉框下方放置 `初始化当前表` 按钮
- 切换到已有绑定信息的 sheet 时，项目下拉框自动回填
- 切换到无绑定信息的 sheet 时，下拉框显示 `先选择项目`
- 所有确认、提示、错误反馈都继续使用 Office/WinForms 原生弹框

## 4. 总体方案选择

本次采用：

- `元数据优先`
- `连接器补种子`
- `自动尝试 + 初始化兜底`

含义如下：

- 运行时以 `_OfficeAgentMetadata` 中当前 sheet 的显式配置为准
- 连接器负责提供字段定义、映射表列定义和首次初始化种子数据
- 用户选择项目后，插件会自动尝试识别当前 sheet
- 如果识别不完整或当前 sheet 尚未完成初始化，用户可显式点击 `初始化当前表`

未来多系统扩展时，插件核心继续只处理“语义角色”和“工作表操作”，具体字段结构由 `systemKey` 对应的连接器提供。

## 5. 元数据模型

`_OfficeAgentMetadata` 在本次设计中保持可见，仅维护两张表：

- `SheetBindings`
- `SheetFieldMappings`

不创建 `SheetSnapshots`。

元数据展示方式采用“同一个 sheet 内上下两个标准表格区域”：

- 第一块区域是 `SheetBindings`
- 第二块区域是 `SheetFieldMappings`
- 每个区域都采用：
  - 一行区域标题
  - 一行表头
  - 多行数据
- 两个区域之间固定留空行

这次明确以“更适合人调试和手工维护”为优先，不再沿用旧的机器压平格式。

同时，本次不兼容旧格式：

- 不做旧格式解析
- 不做旧格式迁移
- 首次发生 metadata 写入时，直接按新布局重建整个 `_OfficeAgentMetadata`

### 5.1 SheetBindings

`SheetBindings` 同时承担：

- 当前 sheet 的项目绑定
- 当前 sheet 的布局配置

展示结构：

| SheetName | SystemKey | ProjectId | ProjectName | HeaderStartRow | HeaderRowCount | DataStartRow |
| --- | --- | --- | --- | --- | --- | --- |
| Sheet1 | current-business-system | project-1 | 绩效项目 | 3 | 2 | 6 |

区域布局示意：

```text
SheetBindings
SheetName | SystemKey | ProjectId | ProjectName | HeaderStartRow | HeaderRowCount | DataStartRow
Sheet1    | ...       | ...       | ...         | 3              | 2              | 6
```

规则：

- 一行对应一张业务 sheet
- 不做项目默认值回退
- 不做系统默认值继承
- 运行时只认这张表里的显式值
- 用户允许直接修改这张表

种子值规则：

- `HeaderStartRow` 默认写入 `1`
- `HeaderRowCount` 默认写入 `2`
- `DataStartRow` 默认写入 `HeaderStartRow + HeaderRowCount`

这些默认值只用于首次创建绑定行，不代表运行时的隐式回退逻辑。

### 5.2 SheetFieldMappings

`SheetFieldMappings` 用于维护：

- 字段身份
- Excel 表头显示名
- 双层表头的父子关系
- 当前 sheet 的字段映射工作副本

规则：

- 一张总表承载多个 sheet 的映射
- 每条映射行必须带 `SheetName`
- `SheetName` 是插件内部固定作用域列
- 除 `SheetName` 外，其余业务列不在插件内部写死
- 这些业务列由 `systemKey` 对应连接器提供列定义和语义角色定义
- 在 `_OfficeAgentMetadata` 中以标准表头 + 数据区展示，而不是按“首列表名 + 行记录”压平存储

当前业务系统的示例形态如下：

| SheetName | HeaderId | HeaderType | ApiFieldKey | IsIdColumn | DefaultSingleDisplayName | CurrentSingleDisplayName | DefaultParentDisplayName | CurrentParentDisplayName | DefaultChildDisplayName | CurrentChildDisplayName | ActivityId | PropertyId |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Sheet1 | row_id | single | row_id | true | ID | ID |  |  |  |  |  |  |
| Sheet1 | owner_name | single | owner_name | false | 负责人 | 项目负责人 |  |  |  |  |  |  |
| Sheet1 | progress_status | single | progress_status | false | 进展状态 | 进展状态 |  |  |  |  |  |  |
| Sheet1 | start_12345678 | activityProperty | start_12345678 | false |  |  | 测试活动111 | 测试活动111 | 开始时间 | 开始时间 | 12345678 | start |
| Sheet1 | end_12345678 | activityProperty | end_12345678 | false |  |  | 测试活动111 | 测试活动111 | 结束时间 | 结束时间 | 12345678 | end |

区域布局示意：

```text
SheetFieldMappings
SheetName | HeaderId | HeaderType | ApiFieldKey | IsIdColumn | ...
Sheet1    | row_id   | single     | row_id      | true       | ...
Sheet1    | owner... | single     | owner...    | false      | ...
```

插件内部只认以下语义角色，不认固定列名：

- 字段稳定标识
- 字段类型
- 接口字段键
- 是否 ID 列
- 单层默认显示名
- 单层当前显示名
- 双层父表头默认显示名
- 双层父表头当前显示名
- 双层子表头默认显示名
- 双层子表头当前显示名
- 活动标识
- 属性标识
- 其他辅助标识

因此未来接别的系统时，可以更换：

- 映射表列名
- 映射表附加字段
- 语义角色与列之间的绑定关系

而不必重写下载上传主流程。

## 6. 连接器扩展边界

虽然当前只接入一个业务系统，但连接器需要继续承担可扩展职责。

当前连接器至少负责提供：

- 项目列表
- `/head` 返回的字段定义
- 当前系统的映射表列定义
- 当前系统的映射表语义角色定义
- 当前项目的映射表种子数据
- `/find` 下载能力
- `/batchSave` 上传能力

建议在当前 `ISystemConnector` 之上补充两个面向配置的返回模型：

- `SheetBindingSeed`
- `FieldMappingTableDefinition`

其中：

- `SheetBindingSeed`
  - 给出首次绑定时建议写入的 `HeaderStartRow`、`HeaderRowCount`、`DataStartRow`
- `FieldMappingTableDefinition`
  - 给出 `SheetFieldMappings` 应有哪些业务列
  - 给出每个语义角色由哪一列承载

## 7. 表头模型与布局规则

### 7.1 HeaderStartRow

表头不再默认固定从第 `1` 行开始。

插件在所有下载、上传、初始化、表头识别过程中，都必须从 `SheetBindings.HeaderStartRow` 读取表头起始行。

### 7.2 HeaderRowCount

表头行数由 `SheetBindings.HeaderRowCount` 决定。

支持值：

- `1`
- `2`

规则：

- 当 `HeaderRowCount = 1` 时，所有表头都在同一行展示
- 当 `HeaderRowCount = 2` 时，才允许双层活动表头
- 当 `HeaderRowCount = 2` 时，单层表头列按纵向合并单元格处理

### 7.3 DataStartRow

数据区不再默认紧接表头之后。

插件在所有上传、下载回写、选区解析、ID 回找过程中，都必须从 `SheetBindings.DataStartRow` 读取数据起始行。

这允许用户在表头和数据区之间插入统计行、汇总行或说明行。

### 7.4 单行与双行表头

当前业务系统允许一个项目内混合存在：

- 单层字段
- 双层活动属性字段

布局规则：

- `HeaderRowCount = 1`
  - 只按单行表头渲染和识别
  - 不支持活动父子表头识别
- `HeaderRowCount = 2`
  - 单层字段使用纵向合并单元格
  - 双层字段使用父表头加子表头

## 8. 初始化与已有 Excel 接管

### 8.1 自动尝试

用户选择项目后，插件立即执行一次轻量自动尝试。

触发条件：

- 当前 sheet 尚未完成初始化，或
- 当前 sheet 还没有 `SheetFieldMappings`

自动尝试流程：

1. 在 `SheetBindings` 中为当前 sheet 建立或更新绑定行
2. 写入首次种子配置
3. 拉取当前项目的字段定义与映射表列定义
4. 判断当前业务 sheet 是否已经存在表头内容
5. 如果已存在，则按当前 `HeaderStartRow` 和 `HeaderRowCount` 识别现有表头
6. 将识别结果和未识别字段一起写入 `SheetFieldMappings`

自动尝试的原则：

- 允许成功
- 允许部分成功
- 不因无法完全识别而强行改写业务 sheet
- 识别不完整时，将控制权交给用户手工修正元数据表

### 8.2 初始化当前表

`初始化当前表` 是显式兜底动作。

触发方式：

- 用户主动点击 Ribbon 按钮
- 或者在上传/下载前发现当前 sheet 未完成初始化时，由弹框提示用户先执行初始化

执行规则：

- 弹框说明此操作会重建当前 sheet 的字段映射配置
- 默认只改 `_OfficeAgentMetadata`
- 若 `_OfficeAgentMetadata` 已存在旧布局或脏布局，直接整张 sheet 按新布局重建
- 不直接改用户当前业务 sheet 的表头或数据

初始化动作适用场景：

- 现有 Excel 首次接入
- 自动尝试未能完整识别
- 用户人工修改了元数据后需要重新生成映射

## 9. 表头识别规则

运行时不依赖固定列号，每次都按当前业务 sheet 的表头文本重新识别列。

### 9.1 单行表头识别

当 `HeaderRowCount = 1` 时：

- 按当前表头文本与映射表中的单层显示名精确匹配
- 优先使用“当前显示名”
- 若当前显示名为空，可回退到默认显示名

### 9.2 双行表头识别

当 `HeaderRowCount = 2` 时：

- 单层字段按纵向合并的单层显示名识别
- 双层字段按“当前父显示名 + 当前子显示名”组合精确匹配
- 若当前显示名为空，可回退到默认显示名

### 9.3 不记录列位置

本设计明确不记录：

- `ColumnIndex`
- 列地址
- 固定列顺序

原因：

- 用户可能自行增删改列
- 插件减少自动化动作，把表结构维护权交给用户
- 每次运行时按表头文本重新识别，更符合已有 Excel 接管场景

后果：

- 用户如果改了 Excel 表头，必须同步维护 `SheetFieldMappings`
- 用户如果增删列，也需要同步维护 `SheetFieldMappings`
- 无法识别的列不参与上传和下载回写

## 10. 下载与上传执行规则

### 10.1 通用前置步骤

所有下载和上传动作都先执行：

1. 读取当前 sheet 的 `SheetBindings`
2. 读取当前 sheet 的 `SheetFieldMappings`
3. 读取当前业务 sheet 的实际表头区域
4. 按当前表头文本重新建立列识别结果

若前置条件不满足：

- 未绑定项目：提示先选择项目
- 未初始化：提示先执行 `初始化当前表`
- 表头无法识别：提示用户修正 `_OfficeAgentMetadata`

### 10.2 全量下载

规则：

- 若当前业务 sheet 已有可识别表头，则只按已识别列写入数据
- 若当前业务 sheet 表头为空，则按 `SheetFieldMappings` 的当前显示名建立表头
- 表头写入从 `HeaderStartRow` 开始
- 数据写入从 `DataStartRow` 开始
- 不自动改动表头上方区域
- 不自动改动表头和数据区之间的统计区域

### 10.3 部分下载

规则：

- 继续采用 `仅可见单元格优先`
- 支持非连续选区
- 选区可以不包含 ID 列
- 选区可以不包含表头行
- 通过当前已识别列和 ID 列回找字段身份
- 只回写目标单元格

### 10.4 全量上传

规则：

- 从 `DataStartRow` 开始扫描数据区
- 只处理有 ID 的行
- 只处理当前可识别到的映射列
- 每个单元格展开为一个 `/batchSave` item

### 10.5 部分上传

规则：

- 只处理当前可见选区中的非 ID 单元格
- 列身份通过当前表头重新识别
- 行身份通过当前 sheet 的 ID 列回找
- 每个目标单元格展开为一个 `/batchSave` item

## 11. 初版确认与错误处理

由于本次不做快照和本地差异追踪，初版确认逻辑保持简单。

### 11.1 下载确认

下载确认框展示：

- 操作名称
- 项目名称
- 记录数
- 字段数
- 是否将覆盖目标区域

不展示：

- 本地脏数据 diff
- 基于快照的覆盖风险

### 11.2 上传确认

上传确认框展示：

- 操作名称
- 项目名称
- 即将提交的单元格数
- 示例改动列表

### 11.3 错误提示

典型错误统一通过原生弹框提示：

- 当前 sheet 未选择项目
- 当前 sheet 未初始化
- `SheetBindings` 缺少必要配置
- `SheetFieldMappings` 缺少 ID 列定义
- 当前表头无法与映射表匹配
- 选区无法回找到字段或 ID
- 接口返回错误

## 12. 当前系统接口约束

当前示例系统继续使用三类接口：

- `/head`
- `/find`
- `/batchSave`

约束如下：

- `/head`
  - 返回字段定义
  - 返回当前系统映射表种子数据所需信息
- `/find`
  - 全量下载和部分下载共用
- `/batchSave`
  - 全量上传和部分上传共用
  - 请求体为单元格改动 list

## 13. 模块调整建议

本次实现重点会集中在以下模块。

### 13.1 Ribbon 与控制器

- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`

调整点：

- 移除 `增量上传`
- 新增 `初始化当前表`
- 将未初始化场景接到弹框兜底

### 13.2 元数据与 Excel 读写

- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
- `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs`
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs`
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetSelectionResolver.cs`

调整点：

- `SheetBindings` 结构重做
- 新增 `SheetFieldMappings` 的读写
- `_OfficeAgentMetadata` 改为同 sheet 上下两个标准表格区域
- 读取逻辑按“区域标题 + 表头行 + 数据区”解析
- 写入逻辑整块重写对应区域，不再使用旧的压平行格式
- 布局与识别逻辑改为读取可配置行号
- 运行时去掉固定列号依赖

### 13.3 执行编排

- `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
- `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`

调整点：

- 所有行号读取改为走 `SheetBindings`
- 新增“自动尝试初始化”和“初始化当前表”
- 移除增量上传执行链路

### 13.4 连接器

- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`
- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSchemaMapper.cs`

调整点：

- 输出 `FieldMappingTableDefinition`
- 输出当前项目的 `SheetFieldMappings` 种子数据
- 保持当前系统实现，同时抽出未来多系统扩展边界

## 14. 测试重点

需要重点覆盖以下场景：

- `SheetBindings` 新结构读写
- `SheetFieldMappings` 动态列定义与语义角色读取
- `HeaderStartRow` 可变
- `HeaderRowCount = 1` 与 `HeaderRowCount = 2`
- `DataStartRow` 可变
- 表头与数据区之间存在统计行
- 现有 Excel 自动识别初始化
- 初始化当前表的显式兜底流程
- 运行时按表头文本重新识别列
- 不依赖 `ColumnIndex` 的部分下载与部分上传
- Ribbon 中移除增量上传后的交互行为

## 15. 结论

本次 Ribbon Sync 可配置性增强采用以下最终方案：

- 只保留全量下载、部分下载、全量上传、部分上传
- 新增 `初始化当前表`，放在项目下拉框下方
- `_OfficeAgentMetadata` 只维护 `SheetBindings` 和 `SheetFieldMappings`
- `SheetBindings` 同时承载项目绑定和行号配置
- `SheetFieldMappings` 是当前 sheet 的字段映射工作副本
- `_OfficeAgentMetadata` 采用更适合人调试的上下双区域表格布局
- 不兼容旧 metadata 压平格式，首次写入直接按新布局重建
- 运行时不记录列位置，每次按表头文本重新识别
- 自动尝试接管已有 Excel，失败后由显式初始化兜底
- 当前系统先落地，内部继续按 `systemKey + 语义角色` 保持可扩展

这套方案优先保证：

- 用户可控
- 元数据可见
- 已有 Excel 可接入
- 初版复杂度可控

同时为后续继续扩展多系统、更多配置项和更强自动化保留了清晰边界。
