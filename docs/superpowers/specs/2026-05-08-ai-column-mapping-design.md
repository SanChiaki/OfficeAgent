# AI Column Mapping Design

日期：2026-05-08

状态：设计已确认，待进入实施计划

## 1. 目标

Ribbon Sync 当前依赖 `SheetFieldMappings` 中的 `Excel L1 / Excel L2` 识别当前工作表列。初始化当前表时，连接器会把业务系统原始表头写入 `ISDP L1 / ISDP L2`，并默认复制到 `Excel L1 / Excel L2`。如果用户的实际 Excel 表头与业务原始表头不一致，用户现在必须手工维护 `xISDP_Setting`。

本功能新增一个独立 Ribbon 按钮，用 AI 根据当前绑定 sheet 的完整表头区，自动建议实际 Excel 表头到原始业务表头的映射。用户确认预览后，插件只把确认的结果写回 `SheetFieldMappings.Excel L1 / Excel L2`，让后续上传和下载继续复用现有表头匹配链路。

目标：

- 降低已有 Excel 接入 Ribbon Sync 时的手工配置成本
- 保持 `ISDP L1 / ISDP L2`、`ApiFieldKey` 和身份列不可被 AI 改写
- 对 AI 输出做结构化校验，并在用户确认后才写入 metadata
- 保持业务单元格不变

## 2. 用户入口

在 Ribbon 顶层 `xISDP` 选项卡的 `项目 / Project` 组新增按钮：

- 中文：`AI映射列`
- 英文：`AI map columns`

按钮位置在 `初始化当前表 / Initialize sheet` 下方。该按钮是独立动作，不自动挂到初始化流程中。

启用和执行前置条件：

- 当前 sheet 必须已选择项目
- 当前 sheet 必须已有 `SheetBindings`
- 当前 sheet 必须已有可用 `SheetFieldMappings`
- 当前 sheet 的表头区必须有可扫描文本
- 本机设置必须已有可用 `BASE_URL / API_KEY / MODEL`

如果未选择项目或未初始化当前表，提示用户先选择项目并执行 `初始化当前表`。

## 3. 范围

### 3.1 本次范围

- 扫描当前绑定 sheet 的完整表头区，而不是用户选区
- 复用现有设置中的 `BASE_URL / API_KEY / MODEL` 调用 OpenAI-compatible chat completions
- 新增专用列映射 LLM client，要求模型返回严格 JSON
- 显示映射预览，用户确认后再写入
- 只更新匹配成功且达到置信度门槛的 `Excel L1 / Excel L2`
- 未匹配或低置信度项不写入，保留原有 metadata
- 更新 Ribbon Sync 当前行为文档和真实系统接入说明

### 3.2 本次明确不做

- 不改变初始化当前表的现有行为
- 不自动改写业务 sheet 可见表头
- 不改写 `ISDP L1 / ISDP L2`
- 不改写 `HeaderId`、`ApiFieldKey`、`IsIdColumn`、`ActivityId`、`PropertyId`
- 不新增业务系统接口，例如 `/autoMapHeaders`
- 不把该能力放进 task pane 对话流
- 不支持用户在预览窗口里手工编辑映射；第一版只确认或取消

## 4. 数据流

1. Ribbon 用户点击 `AI映射列 / AI map columns`。
2. `RibbonSyncController` 校验当前项目状态，并调用 `WorksheetSyncExecutionService` 准备映射预览。
3. 执行服务读取 `SheetBinding`，确定 `HeaderStartRow` 和 `HeaderRowCount`。
4. 执行服务读取完整表头区：
   - 扫描到 `grid.GetLastUsedColumn(sheetName)`
   - `HeaderRowCount = 1` 时，每列提取一个实际 L1
   - `HeaderRowCount > 1` 时，每列提取实际 L1/L2，并沿用当前父表头延续规则处理横向合并场景
5. 执行服务读取当前 `SheetFieldMappings` 和连接器 `FieldMappingTableDefinition`。
6. 构建 LLM 请求：
   - 候选字段：`HeaderId`、`HeaderType`、`ApiFieldKey`、`ISDP L1`、`ISDP L2`、当前 `Excel L1`、当前 `Excel L2`、ID 标记
   - 实际表头：列号、实际 L1、实际 L2、规范化显示文本
7. LLM 返回 JSON 映射建议。
8. 执行服务校验建议：
   - 每个实际 Excel 表头最多匹配一个候选字段
   - 每个候选字段最多被一个实际 Excel 表头占用
   - 返回的 `HeaderId` 或 `ApiFieldKey` 必须存在于候选字段中
   - 置信度必须是可解析数值
9. 宿主显示预览确认窗口。
10. 用户确认后，执行服务只更新通过校验且达到置信度门槛的 mapping 行。
11. `WorksheetMetadataStore.SaveFieldMappings()` 保存更新后的 `SheetFieldMappings`，后续上传和下载继续走 `WorksheetHeaderMatcher`。

## 5. LLM 合同

新增列映射专用 client，放在 Infrastructure 层，复用 `FileSettingsStore` 加载的设置。

请求使用 OpenAI-compatible chat completions：

- endpoint 复用现有 `LlmPlannerClient` 的 base URL 规则
- `model` 使用 `settings.Model`
- `Authorization: Bearer <API_KEY>`
- 要求 JSON object 输出

系统提示约束：

- 只返回 JSON，不返回 markdown
- 只在语义足够明确时给出映射
- 不允许创造不存在的字段
- 不允许把多个 Excel 表头映射到同一个字段
- 保留 ID 字段的特殊性，只有实际表头明显是 ID / row id 时才映射到 ID 字段
- 中文、英文、缩写、同义词可以作为判断依据

建议响应结构：

```json
{
  "mappings": [
    {
      "excelColumn": 3,
      "actualL1": "项目负责人",
      "actualL2": "",
      "targetHeaderId": "owner_name",
      "targetApiFieldKey": "owner_name",
      "confidence": 0.91,
      "reason": "项目负责人与原始字段负责人语义一致"
    }
  ],
  "unmatched": [
    {
      "excelColumn": 5,
      "actualL1": "用户备注",
      "actualL2": "",
      "reason": "没有足够接近的业务字段"
    }
  ]
}
```

第一版建议置信度门槛为 `0.75`。低于门槛的项在预览中显示为低置信度，但确认写入时不更新 metadata。

## 6. 预览确认

新增原生 WinForms 预览对话框，保持 Ribbon Sync 当前“不走任务窗格”的交互模式。

预览至少展示：

- Excel 列号
- 实际表头 L1/L2
- 建议匹配到的 `ISDP L1 / ISDP L2`
- 将写入的 `Excel L1 / Excel L2`
- 置信度
- AI 原因
- 未匹配或低置信度状态

用户操作：

- 确认：写入匹配成功且达到门槛的项
- 取消：不写入任何 metadata

第一版不提供逐项勾选或手工改值，避免把预览窗口做成复杂配置编辑器。若后续需要，可以在同一数据模型上扩展为可编辑 grid。

## 7. Metadata 写入规则

只允许更新当前 sheet 的 `SheetFieldMappings` 行。

对普通 `single`：

- 写入 `CurrentSingleHeaderText`
- 写入同一 RoleKey 对应的 `CurrentParentHeaderText`
- `CurrentChildHeaderText` 保持空，除非该 mapping 本来就是 grouped single

对 grouped single：

- 写入 `CurrentParentHeaderText = actualL1`
- 写入 `CurrentChildHeaderText = actualL2`
- 保持 `HeaderType = single`

对 `activityProperty`：

- 写入 `CurrentParentHeaderText = actualL1`
- 写入 `CurrentChildHeaderText = actualL2`

未匹配、低置信度、重复冲突、无法校验身份的建议全部不写入。保存前再次确认：

- 所有被更新行仍保留原有 `HeaderId / ApiFieldKey`
- 没有两个字段产生相同的运行时匹配键
- `HeaderRowCount = 1` 时不写入需要 `Excel L2` 才能识别的结果

## 8. 错误处理

- 未配置 API key：提示先在设置中配置 API Key
- base URL 非法：复用现有设置错误提示风格
- LLM 请求失败：显示失败原因，不写入 metadata
- LLM 返回非 JSON：提示无法解析 AI 映射结果，不写入
- LLM 返回重复映射或不存在字段：拒绝对应项，必要时拒绝整次写入
- 当前 sheet 未绑定：提示先选择项目
- 当前 sheet 未初始化：提示先执行 `初始化当前表`
- 表头区为空：提示检查 `HeaderStartRow / HeaderRowCount`

业务系统 SSO 不参与该功能。该按钮只调用 LLM 设置，不新增业务系统 API。

## 9. 架构变更

Core：

- 新增 AI 列映射请求、响应、预览、应用结果模型
- 新增纯逻辑服务，用于校验 LLM 建议并应用到 `SheetFieldMappingRow`
- 继续通过 `FieldMappingSemanticRole` 访问动态列，不写死 `Excel L1 / Excel L2` 列顺序

Infrastructure：

- 新增列映射 LLM client
- 尽量复用 `LlmPlannerClient` 中 base URL、chat completions 响应解析的模式

ExcelAddIn：

- `WorksheetSyncExecutionService` 新增准备预览和应用确认结果的方法
- 新增表头扫描 helper，复用 `IWorksheetGridAdapter`
- `RibbonSyncController` 新增执行入口和错误处理
- `AgentRibbon` 新增按钮并本地化
- 新增预览确认 WinForms 对话框

Frontend：

- 不变

## 10. 测试

Core tests：

- 应用器只更新 `Current*HeaderText` 角色
- 未匹配和低置信度不更新
- 重复字段映射被拒绝
- `HeaderRowCount = 1` 下拒绝需要 L2 的结果

Infrastructure tests：

- LLM client 使用现有设置构造 chat completions 请求
- 成功解析 JSON object
- 非 JSON、空响应、HTTP 错误返回明确异常

ExcelAddIn tests：

- 扫描完整单层表头区
- 扫描双层表头区并处理父表头延续
- 预览取消不写入
- 预览确认只写入通过校验的映射
- Ribbon 按钮位于初始化当前表下方
- 中文 / 英文按钮文案覆盖

Manual tests：

- 已绑定并初始化的 sheet，实际表头与 `ISDP L1/L2` 不一致，点击按钮生成预览并确认，`xISDP_Setting` 更新 `Excel L1/L2`
- 未初始化 sheet 点击按钮，提示先初始化
- AI 返回未匹配列时，确认后原有 mapping 保持不变
- 确认写入后，部分上传 / 部分下载可以按新表头识别列

## 11. 文档更新

用户可见行为变化后，同步更新：

- `docs/modules/ribbon-sync-current-behavior.md`
- `docs/ribbon-sync-real-system-integration-guide.md`
- `docs/vsto-manual-test-checklist.md`

如果后续把 AI 映射扩展到 task pane 或模板工作流，再单独补充对应模块快照。
