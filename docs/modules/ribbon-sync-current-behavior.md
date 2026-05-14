# Ribbon Sync Current Behavior

日期：2026-05-12

状态：已实现并可联调。当前只注册了 `current-business-system`，但内部已经落地 `ISystemConnectorRegistry + systemKey` 路由，可继续扩展到多个业务系统。

## 1. 模块范围

Ribbon Sync 是独立于 Agent / task pane 的 Excel Ribbon 数据同步能力。

当前 Ribbon 入口只暴露两个同步动作：

- 下载（按当前可见选区执行部分下载）
- 上传（按当前可见选区执行部分上传）

说明：

- `全量下载` 的底层执行路径仍然保留在代码中
- `全量上传` 的底层执行路径仍然保留在代码中
- 但当前 Ribbon 已隐藏 `全量下载` 和 `全量上传` 按钮，不再对用户直接显示

当前不包含：

- 增量上传
- 本地快照差异比对
- `SheetSnapshots` 元数据表

所有确认、告警、结果反馈都通过 Office / WinForms 原生弹框完成，不走任务窗格。

当前用户可见宿主文案已支持中英文双语：

- 当 Excel UI 语言属于 `zh-*` 时，Ribbon、登录弹窗、布局对话框、同步确认/结果提示、WebView2 宿主兜底提示统一显示中文
- 其他所有 Excel UI 语言统一显示英文
- 当前版本只支持 `zh` / `en` 两套 UI；不跟随 workbook 内容语言或业务数据语言切换
- 如果未来设置中写入 `uiLanguageOverride = zh | en`，宿主 UI 会优先使用该覆盖值；`system` 才回退到 Excel UI 语言

## 2. Ribbon 入口

当前顶层 Ribbon 选项卡标签为 `xISDP`。Ribbon Sync 相关入口分为五组：

- 项目 / `Project`
  - 项目下拉框 / project dropdown
  - `初始化当前表` / `Initialize sheet`
  - `AI映射列` / `AI map columns`
- 配置 / `Setting`
  - `应用配置` / `Apply Setting`
  - `保存配置` / `Save Setting`
  - `另存配置` / `Save as Setting`
- 数据同步 / `Data sync`
  - `下载` / `Download`
  - `上传` / `Upload`
- 帮助 / `Help`
  - `文档` / `Documentation`
  - `关于` / `About`

Ribbon 分组、按钮和项目下拉框状态文案都会跟随当前宿主 UI 语言切换。例如：

- 中文 Excel：`项目`、`配置`、`数据同步`、`账号`、`帮助`、`先选择项目`、`请先登录`
- 英文 Excel：`Project`、`Setting`、`Data sync`、`Account`、`Help`、`Select project`、`Sign in first`

所有 Ribbon 按钮都使用 Office 内置 `imageMso` 图标，并按按钮语义选择图标。`初始化当前表` / `Initialize sheet`、`AI映射列` / `AI map columns` 以及 `配置` / `Setting` 组内的三个命令按钮使用 Office 常规小按钮布局；数据同步、账号、帮助等其余命令按钮使用 Office 大按钮布局，图标显示在上方、文字显示在下方。项目下拉框仍然只显示当前项目或状态文本。

中文 UI 下，两个汉字的 Ribbon 按钮标签会在文本末尾追加 carriage return no-wrap 提示，等价于 Ribbon XML 中的 `&#13;`，用于避免 Office 把 `下载`、`上传`、`文档`、`关于`、`登录` 拆成逐字换行显示。该处理只发生在 Ribbon 按钮赋值层，不改变同步结果、确认弹窗等普通本地化文案。

如果后续要把 Office 内置图标替换为项目自带图片，按 [Ribbon Button Custom Icons Guide](../ribbon-button-custom-icons-guide.md) 操作。

`xISDP AI` 组中的任务窗格按钮只显示图标，不显示 `Open` 文案。`文档` 会用默认浏览器打开 `https://github.com/SanChiaki/OfficeAgent`；`关于` 会显示当前插件版本号、程序集版本、构建配置和构建时间。

Ribbon Sync 会记录运营埋点事件，用于统计项目选择、初始化、AI 映射列、下载、上传、配置按钮和结果状态。事件包含 `systemKey`、`projectId`、`projectName`、`sheetName`、操作类型、确认/取消/成功/失败状态等低敏维度；不会上报单元格原始业务值、API key、cookie 或业务接口请求/响应正文。

主入口代码：

- [src/OfficeAgent.ExcelAddIn/AgentRibbon.cs](../../src/OfficeAgent.ExcelAddIn/AgentRibbon.cs)
- [src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs](../../src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs)

## 3. 元数据模型

运行时元数据都保存在可见工作表 `xISDP_Setting` 中。

兼容规则：

- 如果打开旧工作簿时只存在 `ISDP_Setting`，插件在读取 metadata 时会自动把该 worksheet 重命名为 `xISDP_Setting`，并继续使用原有内容
- 如果 `xISDP_Setting` 和 `ISDP_Setting` 同时存在，插件优先读取 `xISDP_Setting`，不会自动合并两个 sheet

当前使用三张逻辑表：

- `TemplateBindings`
- `SheetBindings`
- `SheetFieldMappings`

当前 `xISDP_Setting` 的展示布局是单个 sheet 内上下三个可读区域：

- 最上区是 `TemplateBindings`
- 中间区是 `SheetBindings`
- 下半区是 `SheetFieldMappings`
- 每个区域都包含：
  - 一行标题
  - 一行表头
  - 多行数据
- 区域之间固定留两行空白分隔

当前不会再使用旧的“首列表名 + 每行一条压平记录”格式；一旦发生 metadata 写入，插件会按上述可读布局整表重写 `xISDP_Setting`。

其中：

- `TemplateBindings` 只记录当前 sheet 与模板资产的关系
- 真正参与下载、上传、初始化执行的运行时事实来源，仍然是 `SheetBindings + SheetFieldMappings`

### 3.1 TemplateBindings

当前列固定为：

- `SheetName`
- `TemplateName`
- `TemplateRevision`
- `TemplateOrigin`
- `TemplateId`
- `AppliedFingerprint`
- `TemplateLastAppliedAt`
- `DerivedFromTemplateId`
- `DerivedFromTemplateRevision`

当前语义：

- `TemplateOrigin = store-template`
  - 当前表已绑定到本机模板库中的正式模板
- `TemplateOrigin = ad-hoc`
  - 当前表只有展开态工作副本，没有绑定固定模板
- `AppliedFingerprint`
  - 记录最近一次“应用配置”或“保存配置”后，对应的归一化模板指纹
- `DerivedFrom...`
  - 记录当前模板最初从哪个模板分叉而来；只表达派生关系，不改变当前“保存回原模板”的目标

### 3.2 SheetBindings

当前列固定为：

- `SheetName`
- `ProjectId`
- `ProjectName`
- `HeaderStartRow`
- `HeaderRowCount`
- `DataStartRow`
- `SystemKey`

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

### 3.3 SheetFieldMappings

`SheetFieldMappings` 的列结构不写死在 Excel 层，实际列由连接器返回的 `FieldMappingTableDefinition` 决定。

当前系统的典型结构示意：

| SheetName | HeaderType | ISDP L1 | ISDP L2 | Excel L1 | Excel L2 | HeaderId | ApiFieldKey | IsIdColumn | ActivityId | PropertyId |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Sheet1 | single | ID |  | ID |  | row_id | row_id | true |  |  |
| Sheet1 | single | 负责人 |  | 负责人 |  | owner_name | owner_name | false |  |  |
| Sheet1 | activityProperty | 测试活动111 | 开始时间 | 测试活动111 | 开始时间 | start_12345678 | start_12345678 | false | 12345678 | start |

说明：

- 第一列固定是 `SheetName`
- 其余列来自业务系统连接器
- 当前 `current-business-system` 会把所有表头显示字段收敛成四列：`ISDP L1`、`ISDP L2`、`Excel L1`、`Excel L2`
- `L1` 对应单层表头文本或双层表头父文本；`L2` 对应双层表头子文本
- 所有 ID / 接口字段相关列都放在显示列之后，便于手工阅读和修改
- Excel 运行时按“语义角色”读取映射，不依赖写死的列顺序
- 当前不会持久化 Excel 列号；每次上传/下载都会重新按当前表头文本识别列
- `AI映射列` 只会更新 `Excel L1` / `Excel L2` 两类当前 Excel 显示文本；不会改写 `ISDP L1` / `ISDP L2`、HeaderId、ApiFieldKey、ID 标记、活动字段标识或其他业务 metadata
- `AI映射列` 的 L1 / L2 写回不按 `HeaderType` 或 `HeaderRowCount` 做业务限制；只要模型推荐通过目标身份、重复和置信度校验，就可以把任意实际 L1 / L2 写到任意字段行的当前 Excel 显示文本
- 旧的六列表头显示模型不再兼容；需要重新执行一次 `初始化当前表` 才会写成新结构

元数据读写代码：

- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs)
- [src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs](../../src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs)
- [src/OfficeAgent.ExcelAddIn/Excel/MetadataSheetLayoutSerializer.cs](../../src/OfficeAgent.ExcelAddIn/Excel/MetadataSheetLayoutSerializer.cs)

性能约束：

- `xISDP_Setting` 读取当前使用 `UsedRange.Value2` 批量读，不再逐单元格 COM 扫描
- `TemplateBindings`、`SheetBindings` 和 `SheetFieldMappings` 在当前活动工作簿内按表级做内存缓存，写入后同步刷新缓存
- 当用户在同一个 Excel 进程里切换到另一个工作簿时，插件会自动失效上一工作簿的 metadata 缓存，避免把 `TemplateBindings` / `SheetBindings` / `SheetFieldMappings` 串到其他 Excel 文件
- 如果用户手工编辑 `xISDP_Setting` 或旧名 `ISDP_Setting`，插件会在对应 `SheetChange` 事件上自动失效上述缓存；下一次业务 sheet 切换或重新触发同步动作时，会重新读取最新元数据

### 3.4 本机模板资产层

当前模板资产不保存在 workbook 中，而是保存在本机模板库：

- `%LocalAppData%\\OfficeAgent\\templates\\<systemKey>\\<projectId>\\<templateId>.json`

当前约束：

- 模板列表按当前 sheet 的 `SystemKey + ProjectId` 过滤
- 模板内容不依赖 workbook，也不持久化具体 `SheetName`
- 用户在业务表上编辑的仍然是 `xISDP_Setting` 展开态，不是只读模板引用

## 4. 项目选择与初始化

### 4.1 项目选择

当前行为：

- 用户先通过项目下拉框选择项目
- 项目下拉框选项通过连接器项目接口动态加载，不再使用本地硬编码项目列表
- 下拉框条目文本显示为 `ProjectId-DisplayName`
- 如果当前 sheet 已有绑定，切换回来时下拉框会自动回填
- 即使项目列表尚未重新加载，下拉框也会先根据 `SheetBindings.ProjectId + ProjectName` 显示当前绑定项目
- 如果当前 sheet 没有绑定，下拉框显示 `先选择项目`
- 如果项目接口返回 `401 Unauthorized` 或 `403 Forbidden`，下拉框显示 `请先登录`
- 当项目接口因未登录返回 `401/403` 时，会弹出本地化提示框：
  - 中文：`当前未登录，请先登录` + `点我登录`
  - 英文：`You're not signed in. Sign in first.` + `Sign in`
  用户可直接触发 Ribbon 登录；登录成功后会立即重载项目列表
- 如果项目接口返回空数组，下拉框显示 `无可用项目`
- 如果项目接口出现其他异常，下拉框显示 `项目加载失败`
- Ribbon 当前项目状态按“活动 sheet 变化”刷新，不再随同一 sheet 内的每次选区移动重复读取 `xISDP_Setting`
- 当同时打开多个 Excel 工作簿时，Ribbon 当前项目状态、`SheetBindings`、`SheetFieldMappings` 都按当前活动工作簿隔离，不会再因为另一个文件里存在同名 sheet 而互相覆盖或串读

一个重要细节：

- 当前 sheet 首次绑定项目，或切换到不同项目时，会先弹出布局对话框
- 布局对话框默认值优先取当前 sheet 已保存的 `HeaderStartRow`、`HeaderRowCount`、`DataStartRow`；如果当前 sheet 还没有绑定记录，则回退到连接器 `CreateBindingSeed` 默认值
- 布局对话框会提示这三个值将写入当前工作表与 ISDP 实施计划的映射配置表 `xISDP_Setting`
- 只有用户在布局对话框点击确认后，才会把项目和布局值写入 `SheetBindings`
- 布局对话框点击取消会完全中止本次项目切换，并恢复下拉框到切换前项目状态
- 重选与当前绑定相同的项目时不会弹出布局对话框，也不会重写 `SheetBindings`
- 布局对话框会根据当前字体自动扩展，避免中文 UI 字体放大时出现标签裁切或控件重叠
- 选择项目不会激活 `xISDP_Setting`
- Ribbon 下拉框内部使用 `systemKey + projectId` 复合键，避免未来多系统下同名 `projectId` 冲突
- Ribbon 下拉框当前显示的是选中条目文本，不单独显示控件标题
- `401/403` 之外的项目加载异常仍走普通失败提示，不会触发登录引导
- `请先登录` / `Sign in first`、`无可用项目` / `No projects available`、`项目加载失败` / `Failed to load projects` 都属于双语 sticky 状态文案；Ribbon 会按当前语言稳定识别并保持这些状态

### 4.2 显式初始化

选择项目后，插件当前只会更新当前 sheet 的 `SheetBindings`，不会自动初始化当前 sheet，也不会自动刷新 `SheetFieldMappings`。

如果当前 sheet 是首次绑定，或者绑定项目变了，用户需要显式点击 `初始化当前表` 来写入或刷新 `SheetFieldMappings`。

如果切换到了其他项目，插件会先清掉当前 sheet 原有的 `SheetFieldMappings`；在用户重新执行 `初始化当前表` 之前，下载和上传都不会静默自动初始化，而是直接报错要求先初始化。

`初始化当前表` 的职责只有两件事：

- 写入 / 刷新 `SheetBindings`
- 写入 / 刷新 `SheetFieldMappings`

它不会改动业务单元格。

认证失败行为：

- 如果初始化过程中底层业务接口返回 `401/403`，不会继续显示普通错误，而是弹出本地化未登录提示框：
  - 中文：`当前未登录，请先登录` + `点我登录`
  - 英文：`You're not signed in. Sign in first.` + `Sign in`
- 用户点击本地化登录按钮后会触发 Ribbon 登录流程
- 本次初始化不会自动重试；登录成功后需要用户重新点击 `初始化当前表`

执行入口：

- [src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs](../../src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs)
- [src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs](../../src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs)
- [src/OfficeAgent.Core/Sync/WorksheetSyncService.cs](../../src/OfficeAgent.Core/Sync/WorksheetSyncService.cs)

### 4.3 AI 映射列

`AI映射列` / `AI map columns` 是一个独立 Ribbon 按钮，位于项目组中 `初始化当前表` / `Initialize sheet` 下方。

它用于已有项目绑定和 `SheetFieldMappings` 的 sheet：当实际 Excel 表头文本和初始化写入的原始表头不一致时，插件会把当前 sheet 上的实际表头推荐映射回 `SheetFieldMappings` 中的原始字段行，并在确认后写入 `Excel L1` / `Excel L2`。

当前流程：

1. 要求当前 sheet 已选择项目并已初始化。
2. 按当前 `SheetBindings.HeaderStartRow` 和 `HeaderRowCount` 扫描完整表头区域，而不是使用当前选区。
3. 单层表头读取实际单元格文本；双层表头读取父 / 子两行文本，并兼容父表头横向合并或只在第一列填写的常见布局。
4. 读取当前 `SheetFieldMappings` 作为候选字段，包含默认 ISDP 表头、当前 Excel 表头、字段类型、HeaderId、ApiFieldKey、ID 字段和活动属性标识。
5. 在本地先过滤已经与 `SheetFieldMappings.Excel L1 / Excel L2` 完全一致的实际表头，并排除 ID 列候选；如果过滤后没有待映射表头，则不调用模型，直接返回没有可应用推荐。
6. 弹出本地 WinForms 处理中对话框，提示 AI 正在分析当前表头，并提供 `中止` / `Abort` 按钮。用户点击中止后会取消本次 HTTP / SSE 模型调用，关闭处理中对话框，不弹出预览，也不会写入 metadata。
7. 调用 AI 列映射客户端生成推荐。该客户端复用现有设置中的 `Base URL`、`API Key`、`Model` 和 `API Format`。
   - `API Format = OpenAI Compatible` 时，请求使用 OpenAI-compatible `chat/completions`：endpoint 为 `/v1/chat/completions` 或保留已有路径前缀后的 `/chat/completions`，认证使用 `Authorization: Bearer <API Key>`。
   - `API Format = Anthropic Messages` 时，请求使用 Anthropic Messages：endpoint 为 `/v1/messages` 或保留已有路径前缀后的 `/messages`，认证使用 `x-api-key`，并发送 `anthropic-version: 2023-06-01`。
   - 当候选字段数量不超过 30 时，请求会保留当前 sheet 中除 ID 列外的全部候选字段。
   - 当候选字段超过 30 时，插件会先用本地文本相似度给候选排序，只把每个实际表头最相关的一批候选拼进 prompt，减少模型输入规模。
   - 本地相似度只用于候选召回和排序，不会直接决定最终映射；最终是否应用仍由模型推荐、预览确认和本地校验共同决定。
   - prompt 中的映射 JSON 使用紧凑序列化，避免多余缩进增加 token。
   - AI 列映射客户端会优先使用 `stream: true` 按 SSE 增量读取模型文本；OpenAI-compatible 格式读取 `choices[].delta.content`，Anthropic Messages 格式读取 `content_block_delta` / `text_delta` 并在 `message_stop` 时结束。如果服务明确不支持流式请求，会自动回退到对应格式的非流式请求。
   - 当前流式只发生在网络读取层，插件仍会等完整 JSON 解析成功后再生成并弹出预览确认，不会边生成边显示预览行。
8. 弹出 WinForms 预览确认对话框，只展示可应用的 accepted 推荐。表格列固定为“是否修改”、Excel 字母列号、当前实际表头 `L1/L2`、匹配到的 ISDP 表头 `L1/L2`；已经等同于当前 `Excel L1 / Excel L2` 的 accepted no-op 推荐不会显示。
9. AI 映射列的处理中、预览确认、完成和错误提示都会以当前 Excel 主窗口为 owner 显示，避免弹框被 Excel 主界面遮挡。
10. “是否修改”默认全部选中；用户取消勾选的行即使原本是 accepted，也不会写入 `SheetFieldMappings`。用户点击取消则不会保存任何 metadata。
11. 用户确认后，仅把仍被勾选且状态为 accepted 的推荐写入 `SheetFieldMappings.Excel L1 / Excel L2`。
12. 完成提示会显示应用数量和跳过数量；如果没有可应用推荐，则只显示没有可应用的推荐。

安全规则：

- ID 列不会作为模型候选发送，模型返回或被篡改出的 ID 列目标也会被拒绝。
- 低置信度、未匹配、目标字段不存在、目标重复、Excel 列重复等推荐不会进入可应用确认列表，也不会应用。
- 本地校验不会因为候选字段是 `activityProperty`、`single` 或其他类型而限制 `Excel L1 / Excel L2` 写回；模型推荐的实际 L1 / L2 会作为自由显示名保存。
- 应用阶段会重新校验预览内容，防止被篡改的 accepted 项写入 metadata。
- 未登录或模型接口配置错误会显示普通错误或登录提示；不会静默写入。

执行入口：

- [src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs](../../src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs)
- [src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs](../../src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs)
- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderScanner.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderScanner.cs)
- [src/OfficeAgent.Infrastructure/Http/AiColumnMappingClient.cs](../../src/OfficeAgent.Infrastructure/Http/AiColumnMappingClient.cs)

### 4.4 模板工作流

当前模板操作全部通过 Ribbon 的“配置”组触发，不走 task pane。

### 应用配置

- 只显示当前项目下的本机模板
- 应用时会把模板展开写回当前 sheet 的 `SheetBindings + SheetFieldMappings`
- 同时会更新 `TemplateBindings`
- 如果当前表相对已绑定模板存在未保存改动，会先提示用户确认覆盖

### 保存配置

- 只对已绑定 `store-template` 的当前表开放
- 保存前会校验当前项目和字段定义是否仍与模板兼容
- 如果模板在本机模板库中的版本已被其他改动推进，会提示用户：
  - 覆盖原模板
  - 另存为新模板
  - 取消本次操作

### 另存配置

- 只要当前表已有项目绑定即可执行
- 会把当前 `xISDP_Setting` 展开态归一化后保存成新模板
- 保存成功后，当前表会切换绑定到新模板
- 此时 `TemplateOrigin` 会写成 `store-template`

## 5. 表头布局与列识别

当前支持同一项目内同时存在：

- 单层表头列
- 双层活动表头列

### 5.0 `HeaderType = single` 的元数据语义

当前 `single` 字段分两种识别形态：

- `HeaderType = single` 且 `Excel L2` 为空：按普通单层列处理
- `HeaderType = single` 且 `Excel L2` 非空：按 grouped single 处理，但字段类型仍然是 `single`

grouped single 当前支持的运行场景：

- `下载`（部分下载执行路径）
- `上传`（部分上传执行路径）
- 已有 grouped single 表头布局时的 `全量下载`

限制：

- 如果当前 sheet 表头区为空，`全量下载` 仍会按普通单层列生成扁平表头，不会因为 `single + Excel L2` 自动生成 grouped single 父表头
- `HeaderRowCount = 1` 时如果 `SheetFieldMappings` 里出现 grouped single 元数据，这是 `xISDP_Setting` 配置错误

### 5.1 HeaderRowCount = 1

当 `HeaderRowCount = 1` 时：

- 所有列都只写一行表头
- 活动属性列只显示一个当前 Excel 表头名，优先取 `Excel L1`；如果旧配置只维护了 `Excel L2`，则回退使用 `Excel L2`
- `single + Excel L2` 不合法；如果 metadata 把单层字段配成 grouped single，则应视为 `xISDP_Setting` 配置错误

### 5.2 HeaderRowCount = 2

当 `HeaderRowCount = 2` 时：

- 单层列会占两行并做纵向合并
- 活动列按活动名在第一行横向合并
- 第二行写活动属性名
- `single + Excel L2` 会按 grouped single 参与运行时识别，但空表头生成阶段仍不会自动写出 grouped single 父表头

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
- `single + Excel L2` 会进入 grouped single 的双层匹配索引，但匹配成功后仍回到 `single` 字段语义执行上传 / 下载
- 匹配阶段会先把 `SheetFieldMappings` 建成单层 / 双层表头索引，再按当前表头文本查找，避免每列重复全表扫描

关键代码：

- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs)
- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs)

## 6. 下载行为

### 6.1 全量下载（当前 Ribbon 按钮已隐藏）

当前流程：

1. 读取 `SheetBindings` 和 `SheetFieldMappings`
2. 尝试按当前表头文本识别运行时列
3. 如果识别成功，只刷新受管数据列的数据区
4. 如果表头区为空，则按 `SheetFieldMappings` 渲染表头，再写数据
5. 如果表头区已有文本但无法匹配映射，则报错，要求先修正表头或元数据
6. 如果业务接口返回 0 条匹配记录，则只提示“查询结果为空，请确认列名是否正确匹配。”，不再展示包含字段数的下载确认框，也不会清空或写入工作表

当前不会重写“已识别成功”的现有表头。

这意味着：

- 如果当前 sheet 已经有人手工维护好的 grouped single 表头，`全量下载` 可以直接复用这套现有布局
- 如果当前 sheet 表头区为空，即使 metadata 里存在 `single + Excel L2`，`全量下载` 也仍会生成扁平 child-only 单层表头，不会生成 grouped single 父表头

这允许用户在表头上方或表头与数据区之间插入统计行，只要 `SheetBindings` 配置正确即可。

性能细节：

- 全量下载写数据时，会先按“受管列是否连续”切成多个连续列段
- 每个连续受管列段使用一次批量 `Range.Value2` 写入
- 非受管列会被跳过，因此用户插入的备注列等非受管区域不会被批量覆盖

### 6.2 下载（部分下载执行路径）

当前流程：

1. 读取当前可见选区的区域快照；如果无法获取区域快照，再回退到逐可见单元格读取
2. 结合运行时匹配到的列，解析出目标 `rowId + fieldKey`
3. 调用 `/find`
4. 仅把查回值回写到原目标单元格
5. 如果 `/find` 返回 0 条匹配记录，则只提示“查询结果为空，请确认列名是否正确匹配。”，不再展示包含字段数的下载确认框，也不会回写任何单元格

认证失败行为：

- 如果底层业务接口返回 `401/403`，不会继续显示普通错误，而是弹出本地化未登录提示框
- 用户点击本地化登录按钮后会触发 Ribbon 登录流程
- 本次 `下载` 不会自动重试；登录成功后需要用户重新触发

当前选区规则：

- 仅可见单元格优先
- 支持非连续选区
- 选区可不包含 ID 列
- 选区可不包含表头行
- 用户全选 Sheet 时，按当前可识别的所有非 ID 受管字段列处理
- 用户整列选择时，只处理所选列中可识别的非 ID 受管字段
- 非连续区域只回写区域内真实选中的行列交叉单元格，不会把不同区域的行列做笛卡尔积扩散写入

性能细节：

- 对整列、整表或大范围选择，下载路径不会逐单元格枚举选区，而是读取可见区域的行列边界
- 大范围下载只在 `DataStartRow..最后使用行` 内按 5000 行批次读取 ID 列，并只把有 `row_id` 且落在选区区域内的行发给 `/find`
- 在一次部分同步操作内，行号到 `row_id` 的查找结果会做内存缓存
- 同一行内多个目标单元格复用同一次 ID 读取，避免重复逐格回查 Excel
- 下载回写前会按连续矩形批量读取旧值，用于写入 `xISDP_Log`，避免逐单元格读取旧值
- 回写阶段会把选中的连续目标单元格归并成矩形批次，并通过 `Range.Value2` 批量写入；非连续选区会拆成多个批次，但不再按单元格逐个写回

关键代码：

- [src/OfficeAgent.ExcelAddIn/Excel/ExcelVisibleSelectionReader.cs](../../src/OfficeAgent.ExcelAddIn/Excel/ExcelVisibleSelectionReader.cs)
- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetSelectionResolver.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetSelectionResolver.cs)
- [src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs](../../src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs)

## 7. 上传行为

### 7.1 全量上传（当前 Ribbon 按钮已隐藏）

当前流程：

1. 从 `DataStartRow` 开始扫描工作表
2. 只处理有 ID 的行
3. 对每个非 ID 列生成一个 `CellChange`
4. 如果当前业务连接器实现了上传过滤扩展点，则先过滤 `CellChange`
5. 上传确认弹窗显示实际上传单元格数、跳过单元格数，以及部分跳过原因
6. 只把实际上传的 `CellChange` 发给 `BatchSave`

性能细节：

- 全量上传会先按已识别受管区域做一次批量 `Range.Value2` 读取
- 同一受管区域的 number format 也会批量读取，用于判断能否安全归一化
- 只有遇到日期、百分比等不安全格式单元格时，才回退到逐单元格 `Text` 读取

### 7.2 上传（部分上传执行路径）

当前流程：

1. 解析当前可见选区
2. 自动回找每个目标单元格所在行的 ID
3. 自动回找该列对应的 `ApiFieldKey`
4. 每个单元格生成一个 `CellChange`
5. 如果当前业务连接器实现了上传过滤扩展点，则先过滤 `CellChange`
6. 上传确认弹窗显示实际上传单元格数、跳过单元格数，以及部分跳过原因
7. 只把实际上传的 `CellChange` 调用 `BatchSave`

上传过滤说明：

- 过滤发生在确认弹窗之前，因此用户看到的实际上传数量和最终提交内容一致
- `SyncOperationPreview.Changes` 只包含实际上传项
- 被跳过的单元格会保存在 `SyncOperationPreview.SkippedChanges`，每项包含原始 `CellChange` 和跳过原因
- 未实现上传过滤扩展点的业务连接器保持现有行为：所有解析出的 `CellChange` 都进入上传预览和提交

认证失败行为：

- 如果底层业务接口返回 `401/403`，不会继续显示普通错误，而是弹出本地化未登录提示框
- 用户点击本地化登录按钮后会触发 Ribbon 登录流程
- 本次 `上传` 不会自动重试；登录成功后需要用户重新触发

性能细节：

- 在一次部分上传操作内，行号到 `row_id` 的查找结果也会做内存缓存
- 同一行内多个目标单元格复用同一次 ID 读取，避免重复逐格回查 Excel

### 7.3 同步日志

当前 workbook 内会维护一个可见工作表 `xISDP_Log`，用于记录 Ribbon Sync 造成的近期业务单元格改动。

`xISDP_Log` 固定列为：

- `key`
- `表头`
- `修改模式`
- `修改值`
- `原始值`
- `修改时间`

字段含义：

- `key` 是业务行 ID，也就是当前系统约定的 `row_id`
- `表头` 使用当前 Excel 表头显示文本；双层表头显示为 `父表头/子表头`
- `修改模式` 固定为 `上传` 或 `下载`
- `修改值` 是同步动作写入 Excel 或提交给业务系统的新值
- `原始值` 在下载时取覆盖前 Excel 值，在上传时取用户编辑前 Excel 值
- `修改时间` 使用本机时间，格式为 `yyyy-MM-dd HH:mm:ss`

保留策略：

- 最多保留最近 2000 条日志
- 超过 2000 条时删除最旧记录
- 如果 `xISDP_Log` 不存在，下一次写日志时自动创建

当前不会记录：

- ID 列本身
- 无 `row_id` 的行
- 未被当前表头映射识别的非受管单元格
- 同步前后文本值未变化的单元格
- 普通手工编辑、初始化当前表、模板操作、`xISDP_Setting` 改动、任务窗格 Agent 写入

上传日志依赖当前 Excel 会话内捕获到的用户编辑前值。插件会在选区变化时缓存待编辑单元格的旧值，并在 `SheetChange` 后标记为 pending；只有 `BatchSave` 成功后才写 `上传` 日志并清除对应 pending 值。如果上传失败，不写日志也不清除 pending 值。

下载日志不依赖 pending 值；插件在覆盖 Excel 单元格前直接读取当前单元格文本作为 `原始值`，写入成功后追加 `下载` 日志。

日志写入失败不会阻断上传 / 下载主流程；失败细节会写入 `%LocalAppData%\\OfficeAgent\\logs\\officeagent.log`。

关键代码：

- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetChangeLogStore.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetChangeLogStore.cs)
- [src/OfficeAgent.ExcelAddIn/Excel/WorksheetPendingEditTracker.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetPendingEditTracker.cs)
- [src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs](../../src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs)

### 7.4 当前边界

当前不支持：

- 增量上传
- 本地脏数据检测
- 无 ID 新增行上传
- 删除行同步
- 服务端并发冲突判断

## 8. 当前业务系统合同

当前系统通过 `ISystemConnector` 抽象接入，并由 `ISystemConnectorRegistry` 聚合项目列表、按 `systemKey` 路由后续下载/上传：

- [src/OfficeAgent.Core/Services/ISystemConnector.cs](../../src/OfficeAgent.Core/Services/ISystemConnector.cs)
- [src/OfficeAgent.Core/Services/ISystemConnectorRegistry.cs](../../src/OfficeAgent.Core/Services/ISystemConnectorRegistry.cs)
- [src/OfficeAgent.Core/Services/SystemConnectorRegistry.cs](../../src/OfficeAgent.Core/Services/SystemConnectorRegistry.cs)

当前实现：

- [src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs](../../src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs)
- [src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs](../../src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs)

当前 mock 契约：

- `GET /projects`
  - Ribbon 项目下拉框加载入口
  - 返回当前系统可用项目列表
  - 当前 mock server 内置 4 个项目：`performance`、`delivery-tracker`、`customer-onboarding`、`large-activity-benchmark`
- `POST /head`
  - 返回 `headList`
  - 按 `projectId` 返回对应项目的字段头
  - 包含所有非活动列
  - 活动只返回活动头，不返回活动属性列
- `POST /find`
  - 全量下载和部分下载共用
  - 按 `projectId` 返回对应项目的数据集
  - `ids` 为空表示全量
  - `fieldKeys` 为空表示返回整行
  - 行数据是平铺 JSON
- `POST /batchSave`
  - 全量上传和部分上传共用
  - 按每个 item 的 `projectId` 写回对应项目
  - 请求体是按单元格变更组成的 list

当前约定的唯一 ID 字段是 `row_id`。

当前项目列表来源：

- Ribbon 启动时，`RibbonSyncController` 通过 `WorksheetSyncService -> SystemConnectorRegistry -> ISystemConnector.GetProjects()` 获取项目列表
- 运行期绑定信息仍然写入 `SheetBindings.SystemKey + ProjectId`
- 后续下载 / 上传始终以 `SheetBindings.SystemKey` 找回对应连接器

连接器认证失败合同：

- Ribbon 不会直接根据底层 HTTP 状态码判断“未登录”；项目下拉框和同步动作都只认 `AuthenticationRequiredException`
- `SystemConnectorRegistry` 只聚合连接器项目列表并透传异常，不会把普通异常翻译成“未登录”
- 因此任意 `ISystemConnector` 的 `GetProjects()`、`BuildFieldMappingSeed()`、`Find()`、`BatchSave()` 在遇到登录失效或无权限场景时，都应把至少 `401/403` 统一转换成 `AuthenticationRequiredException("当前未登录，请先登录")`
- Ribbon 登录提示弹窗不会直接显示 `AuthenticationRequiredException.Message`；用户可见文案由当前宿主 UI 语言决定。异常消息主要用于连接器合同、诊断日志和普通异常区分。
- 如果连接器改为抛普通 `InvalidOperationException`、`HttpRequestException` 或其他异常，项目列表会显示 `项目加载失败`，同步动作也会显示普通错误，不会弹出 `点我登录`

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
  - [src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderScanner.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderScanner.cs)
  - [src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs](../../src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs)
- 当前系统接入
  - [src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs](../../src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs)
  - [src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs](../../src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs)

## 10. 主要测试入口

- 元数据存储
  - [tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs](../../tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs)
- 表头匹配
  - [tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderMatcherTests.cs](../../tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderMatcherTests.cs)
  - [tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderScannerTests.cs](../../tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderScannerTests.cs)
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
  - [docs/superpowers/specs/2026-05-08-ai-column-mapping-design.md](../superpowers/specs/2026-05-08-ai-column-mapping-design.md)
  - [docs/superpowers/specs/2026-04-14-office-agent-ribbon-sync-configurability-design.md](../superpowers/specs/2026-04-14-office-agent-ribbon-sync-configurability-design.md)
  - [docs/superpowers/specs/2026-04-24-office-agent-bilingual-ui-design.md](../superpowers/specs/2026-04-24-office-agent-bilingual-ui-design.md)
- 实施计划
  - [docs/superpowers/plans/2026-05-08-ai-column-mapping-implementation-plan.md](../superpowers/plans/2026-05-08-ai-column-mapping-implementation-plan.md)
  - [docs/superpowers/plans/2026-04-14-office-agent-metadata-layout-implementation-plan.md](../superpowers/plans/2026-04-14-office-agent-metadata-layout-implementation-plan.md)
- 真实系统接入
  - [docs/ribbon-sync-real-system-integration-guide.md](../ribbon-sync-real-system-integration-guide.md)
- 手工测试
  - [docs/vsto-manual-test-checklist.md](../vsto-manual-test-checklist.md)
- Task Pane 快照
  - [docs/modules/task-pane-current-behavior.md](./task-pane-current-behavior.md)

## 12. 文档维护约定

如果 Ribbon Sync 的用户可见行为发生变化，至少同步更新：

- 本文第 2 节到第 8 节
- [docs/module-index.md](../module-index.md)
- [docs/vsto-manual-test-checklist.md](../vsto-manual-test-checklist.md)
