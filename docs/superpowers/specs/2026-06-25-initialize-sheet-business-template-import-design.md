# Initialize Sheet Business Template Import Design

日期：2026-06-25

状态：设计已确认，待实施计划

## 1. 目标

当前 `初始化当前表` 只写入 `SheetBindings + SheetFieldMappings`，不改动业务单元格。空白 Excel 的新用户即使完成初始化，也仍然面对一张空白 sheet，不知道下一步怎么开始。

本设计把业务系统的模板导出能力并入 `初始化当前表`：当用户已选择项目并点击初始化时，插件弹出初始化对话框。用户可以选择从业务系统模板创建当前作业表，也可以选择仅刷新同步配置。空白 sheet 默认从模板创建作业表；非空 sheet 默认仅初始化配置，避免误覆盖用户内容。

目标：

- 让空白 Excel 用户通过现有 `初始化当前表` 入口快速得到可编辑、可上传/下载的作业表
- 继续保留原来的“仅初始化配置”能力
- 让业务系统导出的 Excel 成为新手首表内容的权威来源
- 不改变下载/上传的显式操作边界

## 2. 领域语言

本设计使用以下领域术语，完整词汇见根目录 `CONTEXT.md`：

- `Project`：业务系统中的项目，是 worksheet 同步目标
- `Business Export Template`：业务系统按项目提供的导出模板
- `Business Data Sheet`：业务系统导出 workbook 中名为 `Business Data` 的数据 sheet
- `Work Sheet`：用户当前 workbook 中被初始化、编辑、上传/下载的业务 sheet
- `Sync Configuration Template`：插件已有的本机同步配置模板，和 `Business Export Template` 不是同一个概念

界面文案可以继续使用“模板”，但代码、接口和文档中必须区分业务导出模板与现有同步配置模板。

## 3. 用户入口

入口仍是 Ribbon `xISDP` 选项卡项目组中的：

- 中文：`初始化当前表`
- 英文：`Initialize sheet`

前置条件保持简单：

- 用户必须先通过 Ribbon 项目下拉框选择项目
- 项目选择时继续使用现有布局对话框确认 `HeaderStartRow / HeaderRowCount / DataStartRow`
- 如果未选择项目，沿用现有提示要求先选择项目
- `xISDP_Setting` 和 `xISDP_Log` 禁止初始化和模板导入

初始化按钮在项目已选且当前 sheet 允许初始化时，总是打开新的初始化对话框，不再直接执行初始化。

## 4. 初始化对话框

对话框始终展示两个模式：

- `从模板创建作业表`
- `仅初始化配置`

打开对话框时先加载当前项目的业务模板列表。加载完成或失败后才允许点击确认按钮。

默认选择规则：

- 空白 Work Sheet 且模板可用：默认 `从模板创建作业表`，并选中业务系统返回的第一个模板
- 空白 Work Sheet 但模板为空、加载失败或连接器不支持业务模板导出：默认 `仅初始化配置`
- 非空 Work Sheet：默认 `仅初始化配置`
- 非空 Work Sheet 用户手动选择模板导入时，显示覆盖风险提示，主按钮改为 `覆盖并初始化`

空白判断：

- 只要 sheet 中没有用户输入或导入的单元格内容，就视为空白
- 仅有格式、列宽、行高、冻结窗格或 UsedRange 放大不算非空

模板列表第一版不做搜索，按业务系统返回顺序展示 `templateName`。选择后使用稳定 `templateId` 调用导出接口。

## 5. 业务系统接口合同

新增业务导出模板能力应作为可选连接器扩展实现，不直接加入 `ISystemConnector` 主接口。当前业务系统连接器实现该扩展；未来其他系统不支持时，初始化对话框禁用模板导入并保留仅初始化配置。

建议领域模型：

```csharp
public sealed class BusinessExportTemplateOption
{
    public string TemplateId { get; set; } = string.Empty;
    public string TemplateName { get; set; } = string.Empty;
}
```

建议扩展接口：

```csharp
public interface IBusinessExportTemplateConnector
{
    IReadOnlyList<BusinessExportTemplateOption> GetBusinessExportTemplates(string projectId);

    BusinessExportWorkbook ExportBusinessWorkbook(
        string projectId,
        string templateId,
        CancellationToken cancellationToken);
}
```

接口合同：

- 模板列表最少返回 `templateId` 和 `templateName`
- `templateId` 是稳定 ID，不能用展示名代替
- 导出接口按 `projectId + templateId` 返回 `.xlsx` 二进制文件
- 第一版只支持标准 `.xlsx`
- 不支持 `.xls`、`.csv`、加密 workbook 或宏工作簿作为导入格式
- 导出 workbook 必须包含名为 `Business Data` 的 sheet
- 业务系统保证 `Business Data` 自包含，并保证其表头布局与项目绑定布局、初始化写入的 `SheetFieldMappings` 天然一致

插件不要求导出接口额外返回字段映射 metadata。初始化 Setting 仍由现有项目初始化链路生成。

## 6. 模板导入数据流

从模板创建作业表的成功路径：

1. 校验当前 sheet 不是 `xISDP_Setting` 或 `xISDP_Log`
2. 校验当前 sheet 已选择项目并存在有效 `SheetBinding`
3. 校验 workbook/sheet 未被保护到无法写入
4. 加载当前项目业务模板列表
5. 用户选择模板导入模式和模板
6. 显示进度对话框
7. 调用业务系统导出接口下载 `.xlsx`
8. 打开临时 workbook
9. 查找 `Business Data` sheet
10. 预先准备初始化所需的 `SheetBinding`、`FieldMappingTableDefinition`、`SheetFieldMappings`
11. 确认 metadata sheet 可写
12. 将 `Business Data` 内容导入到当前 Work Sheet
13. 保留当前 Work Sheet 名称
14. 清理旧 field mappings，并写入当前项目的 `SheetBindings + SheetFieldMappings`
15. 激活当前 Work Sheet，选中 `A1`
16. 删除临时导出文件
17. 显示成功提示

导入不记录业务模板 `templateId/templateName` 到 `xISDP_Setting`。业务模板只是初始化时的一次性输入，不成为 workbook 的长期绑定事实。

## 7. Sheet 导入策略

空白 Work Sheet 默认导入到当前 sheet 并保留当前 sheet 名称。非空 Work Sheet 只有用户明确选择模板导入路径时才允许覆盖当前 sheet。

导入应尽量完整保留 `Business Data` 的作业体验：

- 单元格值
- 公式
- 格式
- 合并单元格
- 列宽、行高
- 隐藏行列
- 冻结窗格/视图状态，尽量保留

第一版不承诺：

- VBA
- 外部数据连接
- 工作簿级名称
- 保护密码
- 导出 workbook 中的其他 sheet
- 跨 sheet 公式引用的完整可用性

业务系统负责保证 `Business Data` 是适合单 sheet 导入的作业表。

## 8. 仅初始化配置数据流

用户选择 `仅初始化配置` 时，执行现有初始化语义：

- 写入/刷新当前项目 `SheetBindings`
- 清理并重写当前项目 `SheetFieldMappings`
- 不改动业务单元格内容
- 不导出业务系统 Excel
- 不记录模板信息

成功提示改为强调当前表内容未修改。

## 9. 进度与取消

模板导入路径需要进度对话框。

阶段建议：

- `正在下载模板 Excel...`：显示取消按钮
- `正在导入到当前工作表...`：取消按钮禁用或隐藏
- `正在写入同步配置...`：取消按钮禁用或隐藏

第一版只保证取消业务系统 Excel 下载 I/O。用户在下载阶段点击取消：

- 中止 HTTP 请求
- 关闭进度对话框
- 当前 sheet 不变
- 不写 Setting
- 不显示错误
- 记录 info 日志和取消埋点

COM 复制阶段暂不支持安全中止。

## 10. 错误处理

模板列表为空：

- 禁用 `从模板创建作业表`
- 提示当前项目暂无可用模板
- 默认并允许 `仅初始化配置`

模板列表加载失败：

- 禁用 `从模板创建作业表`
- 提示模板加载失败，可先仅初始化配置
- 默认并允许 `仅初始化配置`

连接器不支持业务模板导出：

- 禁用 `从模板创建作业表`
- 提示当前业务系统不支持从模板创建作业表
- 默认并允许 `仅初始化配置`

导出或导入失败：

- 不自动降级到仅初始化配置
- 不写 Setting
- 当前 sheet 在下载和预检失败时保持不变
- 若内容已导入但最后 metadata 写入失败，提示严重错误：表内容已导入但同步配置未完成，请重新初始化当前表

失败场景包括：

- 导出接口失败/超时
- 未登录或无权限
- 下载内容不是有效 `.xlsx`
- 找不到 `Business Data` sheet
- workbook/sheet 保护导致不能写入
- 导入或 metadata 写入失败

导出的临时 Excel 文件默认删除，失败时也不保留业务数据文件。日志只记录错误类型、阶段和低敏维度，不记录业务单元格内容。

## 11. 成功反馈

模板导入成功：

> 初始化完成，已从模板创建当前作业表。你可以编辑数据后上传，或全选需要刷新的区域后点击下载。

仅初始化配置成功：

> 初始化完成，当前表内容未修改。你可以继续上传或下载数据。

中英文文案必须通过 `HostLocalizedStrings` 管理，不在对话框或控制器中硬编码。

初始化完成后：

- 不自动执行下载
- 不自动打开任务窗格
- 激活当前 Work Sheet
- 选中 `A1`
- 不写 `xISDP_Log`

## 12. Analytics

新增模板导入路径埋点，只上报低敏维度，不包含业务单元格内容。

建议事件：

- `ribbon.initialize_template_import.started`
- `ribbon.initialize_template_import.canceled`
- `ribbon.initialize_template_import.completed`
- `ribbon.initialize_template_import.failed`

建议维度：

- `projectId`
- `projectName`
- `templateId`
- `sheetName`
- `isBlankSheet`
- `durationMs`
- `failedStage`
- `exceptionType`

模板列表为空、加载失败、连接器不支持等降级状态也应记录，便于评估新手路径覆盖率。

## 13. 不做事项

本设计明确不做：

- 不改造项目选择入口
- 不把项目选择放进初始化对话框
- 不移除项目选择时的布局对话框
- 不改变下载/上传未初始化时的现有报错行为
- 不把全量下载替换为业务系统 Excel 导出
- 不自动下载刷新导入后的数据
- 不自动打开任务窗格
- 不记录业务模板 ID 到 `xISDP_Setting`
- 不复用现有本机模板目录或 `ITemplateCatalog`

## 14. 测试

Core tests：

- 可选连接器扩展不影响未实现业务导出模板的系统
- 业务导出模板模型只依赖 `templateId/templateName`

Infrastructure tests：

- 当前业务连接器请求模板列表
- 当前业务连接器按 `projectId/templateId` 下载 `.xlsx` 二进制
- 401/403 转换为认证要求异常
- 下载 cancellation 不作为失败处理

ExcelAddIn tests：

- 空白 sheet 默认选择模板导入
- 非空 sheet 默认选择仅初始化配置
- 模板为空时禁用模板导入并允许仅初始化配置
- 模板加载失败时禁用模板导入并允许仅初始化配置
- 连接器不支持业务模板导出时禁用模板导入
- `xISDP_Setting` / `xISDP_Log` 上禁止初始化
- 下载取消不写 Setting、不改 sheet
- 找不到 `Business Data` 时不写 Setting
- 模板导入成功后写入当前 sheet 名对应的 `SheetBindings + SheetFieldMappings`
- 模板导入成功后不写 `xISDP_Log`
- 成功提示按路径区分
- 新对话框文案均来自 `HostLocalizedStrings`

Manual tests：

- 新建空白 Excel，选择项目，确认布局，点击初始化，默认从模板创建作业表，导入完成后当前 sheet 有业务数据且名称保持不变
- 非空 sheet 点击初始化，默认仅初始化配置，不改动业务单元格
- 非空 sheet 手动选择模板导入，看到覆盖风险提示，确认后当前 sheet 被模板内容覆盖
- 模板下载阶段点击取消，当前 sheet 不变
- 导出的 workbook 缺少 `Business Data`，显示错误且当前 sheet 不变
- 导入后手动全选数据区域点击下载，仍复用现有部分下载路径

## 15. 文档更新

用户可见行为变化后，同步更新：

- `docs/modules/ribbon-sync-current-behavior.md`
- `docs/ribbon-sync-real-system-integration-guide.md`
- `docs/vsto-manual-test-checklist.md`
- 如 mock server 增加模板列表和导出接口，更新 `tests/mock-server/README.md`

本设计还配套了：

- `CONTEXT.md`
- `docs/adr/0001-import-business-data-into-current-sheet.md`
- `docs/adr/0002-require-explicit-confirmation-for-nonblank-template-import.md`
- `docs/adr/0003-keep-project-selection-out-of-initialization-dialog.md`
- `docs/adr/0004-use-business-export-templates-for-sheet-initialization.md`
- `docs/adr/0005-limit-template-import-cancellation-to-download.md`
