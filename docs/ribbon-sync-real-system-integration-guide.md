# Ribbon Sync 真实业务系统接入指南

本文说明后续把 Ribbon Sync 从当前 mock / 示例系统切换到真实业务系统时，建议如何改造。

目标有两个：

1. 先把当前插件稳定接入一个真实系统
2. 在架构上保留未来接入多个系统的扩展空间

## 1. 先理解当前架构

当前 Ribbon Sync 的核心思路已经从“固定列号 + 快照差异”切换为：

- `xISDP_Setting` 是每个受管 sheet 的运行时事实来源
- `xISDP_Setting.TemplateBindings` 记录当前 sheet 与本机模板库的关系
- `SheetBindings` 记录项目绑定和表格行位置信息
- `SheetFieldMappings` 记录字段映射和当前 Excel 显示名
- 上传 / 下载时总是按当前表头文本重新识别列
- `AI映射列` / `AI map columns` 可以扫描当前 sheet 的完整配置表头区域，把实际 Excel 表头推荐映射到 `SheetFieldMappings.Excel L1 / Excel L2`
- 本机模板资产保存在 `%LocalAppData%\OfficeAgent\templates\...`，不直接替代运行时 metadata
- 当前 Ribbon 入口只做部分下载、部分上传
- `全量下载` 和 `全量上传` 的执行路径仍保留在代码中，但当前按钮已隐藏

当前 `xISDP_Setting` 的具体形态也已经固定：

- 它是一个可见 worksheet，便于调试和人工维护
- 它当前承载三个 section：
  - `TemplateBindings`
  - `SheetBindings`
  - `SheetFieldMappings`
- 三个 section 都采用同样的可读布局：
  - 一行标题
  - 一行表头
  - 多行数据
- `TemplateBindings` 永远在最上
- `SheetBindings` 永远在中间
- `SheetFieldMappings` 永远在最下
- 相邻 section 中间固定保留两行空白
- 当前不再使用旧的“首列表名 + 每行一条压平记录”格式
- 一旦发生 metadata 写入，插件会按这个标准布局整表重写 `xISDP_Setting`
- 旧工作簿如果只存在 `ISDP_Setting`，插件读取 metadata 时会自动把该 worksheet 重命名为 `xISDP_Setting`

当前不做：

- 增量上传
- 本地快照差异
- `SheetSnapshots` 元数据表

## 2. 当前插件对业务系统的最小依赖面

核心抽象在 [src/OfficeAgent.Core/Services/ISystemConnector.cs](../src/OfficeAgent.Core/Services/ISystemConnector.cs)：

```csharp
public interface ISystemConnector
{
    string SystemKey { get; }

    IReadOnlyList<ProjectOption> GetProjects();

    SheetBinding CreateBindingSeed(string sheetName, ProjectOption project);

    FieldMappingTableDefinition GetFieldMappingDefinition(string projectId);

    IReadOnlyList<SheetFieldMappingRow> BuildFieldMappingSeed(string sheetName, string projectId);

    WorksheetSchema GetSchema(string projectId);

    IReadOnlyList<IDictionary<string, object>> Find(
        string projectId,
        IReadOnlyList<string> rowIds,
        IReadOnlyList<string> fieldKeys);

    void BatchSave(string projectId, IReadOnlyList<CellChange> changes);
}
```

项目聚合和运行时路由在：

- [src/OfficeAgent.Core/Services/ISystemConnectorRegistry.cs](../src/OfficeAgent.Core/Services/ISystemConnectorRegistry.cs)
- [src/OfficeAgent.Core/Services/SystemConnectorRegistry.cs](../src/OfficeAgent.Core/Services/SystemConnectorRegistry.cs)

其中真正参与当前主链路的能力是：

- `SystemKey`
- `GetProjects`
- `CreateBindingSeed`
- `GetFieldMappingDefinition`
- `BuildFieldMappingSeed`
- `Find`
- `BatchSave`

`GetSchema` 目前更多保留给连接器测试和辅助逻辑，不是当前 Excel 主执行链路的核心入口。

## 3. Excel 侧现在如何工作

Ribbon 点击链路：

1. [src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs](../src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs)
2. [src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs](../src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs)
3. [src/OfficeAgent.Core/Sync/WorksheetSyncService.cs](../src/OfficeAgent.Core/Sync/WorksheetSyncService.cs)
4. `ISystemConnectorRegistry`
5. `ISystemConnector`

说明：

- 如果只是接入新的业务接口，优先新增或替换连接器，再把它注册到 `SystemConnectorRegistry`
- 只有当真实系统的表头模型、选区解释规则或上传粒度不同，才需要继续改 Excel 层

## 4. 真实系统至少要提供什么

### 4.1 项目列表

项目下拉框需要：

- `projectId`
- `displayName`

其中：

- `systemKey` 由连接器自身提供，不要求项目接口返回
- 如果项目接口返回了 `systemKey`，当前注册表仍会以连接器自己的 `SystemKey` 为准
- 项目列表统一由 `ISystemConnector.GetProjects()` 提供，Ribbon 本身不关心底层是真实接口、聚合接口还是静态配置

当前 Ribbon 对项目列表的用户可见行为是：

- 当前 sheet 没有绑定时，下拉框显示 `先选择项目`
- 项目接口返回有效列表时，下拉框显示 `ProjectId-displayName` 形式的项目条目
- 重选当前已绑定的同一个 `systemKey + projectId` 时是 no-op：不会弹出布局对话框，也不会重写 `SheetBindings`
- 项目接口返回 `401 Unauthorized` 或 `403 Forbidden` 时，Ribbon 会显示 `请先登录`
- 当项目接口因未登录返回 `401/403` 时，会弹出本地化未登录提示框；中文 UI 显示 `当前未登录，请先登录` + `点我登录`，英文 UI 显示 `You're not signed in. Sign in first.` + `Sign in`。用户可直接触发 Ribbon 登录；登录成功后会立即重载项目列表
- 项目接口返回空数组时，Ribbon 会显示 `无可用项目`
- 项目接口发生其他异常时，Ribbon 会显示 `项目加载失败`

因此接入真实系统时，项目接口至少要明确：

- 未登录时的返回状态码
- 空项目列表是不是合法业务状态
- 是否必须先经过 SSO 登录才能访问项目列表

认证失败接入合同：

- Ribbon 项目下拉框不会直接根据 HTTP 状态码判断“未登录”；它只会在连接器最终抛出 `AuthenticationRequiredException` 时弹出登录提示
- `SystemConnectorRegistry` 只负责聚合各个连接器的 `GetProjects()` 返回值和异常，不会把其他异常类型翻译成“未登录”
- 因此真实连接器的 `GetProjects()` 在遇到未登录或无权限场景时，至少应把 `401/403` 统一转换成 `AuthenticationRequiredException("当前未登录，请先登录")`
- Ribbon 登录提示弹窗不会直接显示 `AuthenticationRequiredException.Message`；用户可见文案由当前宿主 UI 语言决定。异常消息主要用于连接器合同、诊断日志和普通异常区分。
- 如果连接器抛出的是普通 `InvalidOperationException`、`HttpRequestException` 或其他异常，Ribbon 会把它当成普通项目加载失败处理，不会出现 `点我登录`

相关代码入口：

- [src/OfficeAgent.ExcelAddIn/AgentRibbon.cs](../src/OfficeAgent.ExcelAddIn/AgentRibbon.cs)
- [src/OfficeAgent.Core/Services/SystemConnectorRegistry.cs](../src/OfficeAgent.Core/Services/SystemConnectorRegistry.cs)
- [src/OfficeAgent.Core/AuthenticationRequiredException.cs](../src/OfficeAgent.Core/AuthenticationRequiredException.cs)

对应模型：

- [src/OfficeAgent.Core/Models/ProjectOption.cs](../src/OfficeAgent.Core/Models/ProjectOption.cs)

### 4.2 绑定默认值

连接器要为新绑定 sheet 提供布局对话框默认值：

- `HeaderStartRow`
- `HeaderRowCount`
- `DataStartRow`

当前默认值是：

- `HeaderStartRow = 1`
- `HeaderRowCount = 2`
- `DataStartRow = 3`

如果你的真实系统有别的默认布局，可以在 `CreateBindingSeed` 里改。

注意：

- `CreateBindingSeed` 现在用于“项目切换时布局对话框的默认值”，不是无提示直接落盘的持久化值
- 当前 sheet 首次绑定项目时，布局对话框会展示连接器 seed 值（例如 `1 / 2 / 3`）
- 切换到其他项目时，布局对话框会优先使用当前 sheet 已保存的布局值；只有不存在时才回退到 `CreateBindingSeed`
- 布局对话框取消会回滚到切换前绑定与下拉框状态，并中止本次项目切换
- 只有用户在布局对话框点击确认后，才会把最终值写入 `SheetBindings`

### 4.3 字段映射定义

连接器必须定义 `SheetFieldMappings` 的动态列结构，也就是：

- 这张元数据表有哪些列
- 每一列承担什么语义角色

这里有两个实现约束：

- Excel 层只固定 `SheetName` 是第一列作用域列
- 除 `SheetName` 外，其余业务列都由连接器定义，并最终落到 `xISDP_Setting` 里的 `SheetFieldMappings` section 中

当前 `current-business-system` 的展示列顺序是：

- `HeaderType`
- `ISDP L1`
- `ISDP L2`
- `Excel L1`
- `Excel L2`
- `HeaderId`
- `ApiFieldKey`
- `IsIdColumn`
- `ActivityId`
- `PropertyId`

当前显示语义是：

- `L1` 表示单层表头文本，或双层表头的父文本
- `L2` 表示双层表头的子文本
- `ISDP` 表示默认值
- `Excel` 表示当前 Excel 中希望匹配 / 展示的值

如果你的真实系统也希望把“显示字段”和“标识字段”分区展示，建议沿用这种顺序，把所有 ID / 接口字段相关列放在显示列之后。

当前系统的例子在：

- [src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs](../src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs)
- [src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs](../src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs)

### 4.4 映射种子数据

初始化当前表时，连接器要生成 `SheetFieldMappings` 的首批数据。

当前推荐做法：

1. 调用真实系统的字段头接口
2. 拿到所有非活动字段 + 活动头
3. 再通过一次样本查询拿到平铺行数据
4. 从平铺字段里识别活动属性列
5. 生成 `SheetFieldMappings`

### 4.5 AI 自动映射列

真实系统接入后，如果用户已有的 Excel 表头和初始化种子里的 `ISDP L1 / ISDP L2` 不完全一致，可以使用 Ribbon 项目组里的 `AI映射列` / `AI map columns` 辅助维护 `SheetFieldMappings.Excel L1 / Excel L2`。

接入约束：

- 真实连接器必须给 `SheetFieldMappings` 列定义提供清晰的语义角色，包括默认表头、当前 Excel 表头、HeaderId、ApiFieldKey、ID 标记，以及活动字段的 ActivityId / PropertyId。
- AI 映射请求会使用这些语义角色构造候选字段，不依赖固定列顺序。
- AI 映射请求会先在本地过滤已经与 `SheetFieldMappings.Excel L1 / Excel L2` 完全一致的实际表头，并排除 ID 列候选；如果过滤后没有待映射表头，则不会调用模型。
- 当当前 sheet 的候选字段超过 30 时，AI 映射会先按实际表头和候选表头的本地文本相似度做候选召回，只把相关候选拼进 prompt；该本地排序不决定最终映射。
- AI 映射客户端复用插件设置里的 `Base URL`、`API Key`、`Model` 和 `API Format`，所以不需要为真实系统新增单独的模型配置；只要现有 LLM 设置可用即可。
- `API Format` 支持 `OpenAI Compatible` 和 `Anthropic Messages`。OpenAI-compatible 使用 `/v1/chat/completions` 或保留路径前缀后的 `/chat/completions`，并发送 `Authorization: Bearer <API Key>`；Anthropic Messages 使用 `/v1/messages` 或保留路径前缀后的 `/messages`，并发送 `x-api-key` 与 `anthropic-version: 2023-06-01`。
- AI 映射客户端优先使用 `stream: true` 读取模型输出；OpenAI-compatible 读取 `choices[].delta.content`，Anthropic Messages 读取 `content_block_delta` / `text_delta` 并在 `message_stop` 时结束。服务不支持流式请求时会回退到对应格式的非流式调用。当前预览仍等完整 JSON 返回后再显示。
- 模型调用期间会显示本地处理中对话框；用户点击 `中止` / `Abort` 会取消本次 HTTP / SSE 调用，关闭处理中对话框，不弹出预览，也不会写入 `xISDP_Setting`。
- 表头扫描按 `SheetBindings.HeaderStartRow + HeaderRowCount` 扫描完整表头区域，不依赖当前选区。
- 用户必须先确认预览；取消预览不会写入 `xISDP_Setting`。
- AI 映射列的处理中、预览确认、完成和错误提示会以 Excel 主窗口为 owner 显示，避免弹框被 Excel 主界面遮挡。
- 预览只展示可应用的 accepted 建议；确认表格包含“是否修改”、Excel 字母列号、当前实际表头 `L1/L2`、匹配到的 ISDP 表头 `L1/L2`。已经等同于当前 `Excel L1 / Excel L2` 的 accepted no-op 建议不会显示。
- “是否修改”默认全部选中；用户取消勾选的建议不会写入 `xISDP_Setting`。
- 确认后只更新 `Excel L1 / Excel L2`，不会覆盖 `ISDP L1 / ISDP L2`、HeaderId、ApiFieldKey、ID 列标记或活动字段标识。
- 低置信度、未匹配、重复目标、重复 Excel 列和 ID 列建议不会进入可应用确认列表，也不会应用。
- `Excel L1 / Excel L2` 是可自由分配的当前显示名；AI 映射确认写回时不会因为候选行是 `activityProperty`、`single` 或其他类型而限制 L1 / L2 组合，也不会因为当前 `HeaderRowCount` 限制模型推荐的 L2。

这意味着真实系统不应把 AI 映射当成字段定义来源。字段定义仍然来自连接器和业务接口；AI 只帮助把用户 Excel 中的实际显示名填回当前显示文本列。

### 4.6 查询接口

插件对 `Find` 的要求是：

- `rowIds` 为空时能返回全量数据
- `fieldKeys` 为空时能返回整行字段
- 返回结果是“平铺后的行数据 list”

示意：

```json
[
  {
    "row_id": "row-1",
    "owner_name": "张三",
    "start_12345678": "2026-01-02",
    "end_12345678": "2026-01-05"
  }
]
```

### 4.7 更新接口

当前上传不是按整行提交，而是按单元格提交 `CellChange`。

也就是说，真实系统如果只有“整行更新”接口，需要在连接器内部把这些单元格改动聚合成目标系统所需 payload，不要把这个复杂度上推到 Excel 层。

### 4.8 上传过滤

如果真实业务系统需要按业务规则跳过部分单元格，不要只在 `BatchSave()` 里静默过滤。当前推荐做法是让真实连接器额外实现 `IUploadChangeFilter`：

- `FilterUploadChanges(projectId, changes)` 返回实际上传项和跳过项
- 实际上传项进入 `SyncOperationPreview.Changes`，确认后才会传给 `BatchSave()`
- 跳过项进入 `SyncOperationPreview.SkippedChanges`，每项包含原始 `CellChange` 和跳过原因
- 上传确认弹窗会显示实际上传数量、跳过数量和部分跳过原因

这保证了用户确认时看到的内容和最终提交给业务系统的内容一致。

接口位置：

- [src/OfficeAgent.Core/Services/IUploadChangeFilter.cs](../src/OfficeAgent.Core/Services/IUploadChangeFilter.cs)
- [src/OfficeAgent.Core/Models/UploadChangeFilterResult.cs](../src/OfficeAgent.Core/Models/UploadChangeFilterResult.cs)
- [src/OfficeAgent.Core/Models/SkippedCellChange.cs](../src/OfficeAgent.Core/Models/SkippedCellChange.cs)

调用时机：

1. Excel 层先把全量上传或部分上传解析成 `CellChange[]`
2. `WorksheetSyncService.FilterUploadChanges(systemKey, projectId, changes)` 找到当前 `systemKey` 对应连接器
3. 如果连接器实现了 `IUploadChangeFilter`，调用连接器的 `FilterUploadChanges(projectId, changes)`
4. 过滤结果生成上传预览
5. 用户确认后，只把 `IncludedChanges` 传给 `BatchSave(projectId, changes)`

如果连接器没有实现 `IUploadChangeFilter`，当前默认行为是所有 `CellChange` 都进入上传。也就是说，上传过滤是可选扩展点，不会影响没有过滤需求的系统。

`CellChange` 已经包含过滤常用字段：

- `SheetName`
- `RowId`
- `ApiFieldKey`
- `OldValue`
- `NewValue`

建议把过滤器只用于“包含 / 跳过”的业务判断，不要在过滤器里改写 `NewValue` 或把一个 `CellChange` 拆成多个变更。如果真实接口需要格式归一化、字段映射或按行聚合，优先放在 `BatchSave()` 内部处理；这样 Excel 预览、上传日志和实际提交内容的对应关系更稳定。

返回约定：

- `IncludedChanges`：实际允许上传的原始 `CellChange`
- `SkippedChanges`：被跳过的项，每项都应保留原始 `CellChange`
- `SkippedCellChange.Reason`：面向最终用户展示的原因，应简短、可读，不要包含接口 token、内部异常栈或敏感字段
- 如果确实没有过滤需求，可以不实现 `IUploadChangeFilter`
- 如果已经实现过滤器，不要返回 `null` 数组；用 `Array.Empty<CellChange>()` 或 `Array.Empty<SkippedCellChange>()` 表达空集合

示例只展示过滤器核心逻辑，`ISystemConnector` 其他成员和 `LoadRowStates` 里的真实接口调用按实际系统补齐：

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

public sealed class RealBusinessSystemConnector : ISystemConnector, IUploadChangeFilter
{
    private static readonly HashSet<string> ReadOnlyFields = new HashSet<string>(
        StringComparer.Ordinal)
    {
        "created_at",
        "approved_at",
    };

    public UploadChangeFilterResult FilterUploadChanges(
        string projectId,
        IReadOnlyList<CellChange> changes)
    {
        var included = new List<CellChange>();
        var skipped = new List<SkippedCellChange>();
        var changeList = changes ?? Array.Empty<CellChange>();
        var rowStates = LoadRowStates(
            projectId,
            changeList.Select(change => change.RowId).Distinct().ToArray());

        foreach (var change in changeList)
        {
            if (ReadOnlyFields.Contains(change.ApiFieldKey))
            {
                skipped.Add(new SkippedCellChange
                {
                    Change = change,
                    Reason = "字段只读，禁止上传",
                });
                continue;
            }

            if (rowStates.TryGetValue(change.RowId, out var rowState) && rowState.IsArchived)
            {
                skipped.Add(new SkippedCellChange
                {
                    Change = change,
                    Reason = "单据已归档，禁止上传",
                });
                continue;
            }

            included.Add(change);
        }

        return new UploadChangeFilterResult
        {
            IncludedChanges = included.ToArray(),
            SkippedChanges = skipped.ToArray(),
        };
    }

    // ISystemConnector 其他成员省略。
}
```

过滤规则可以是连接器内置逻辑、本地配置、`xISDP_Setting` 派生规则，或业务接口下发的字段 / 行状态规则。典型规则包括：

- 按 `ApiFieldKey` 跳过只读字段
- 按 `RowId` 对应的单据状态跳过已归档、已锁定、流程结束的数据
- 按 `NewValue` 跳过空值、默认值或业务系统不接受的值
- 调业务接口校验权限后跳过不可编辑字段

如果过滤规则需要调用真实业务接口，建议注意：

- 尽量按本次上传涉及的 `RowId` / `ApiFieldKey` 批量查询权限或状态，避免每个单元格一次 HTTP 请求
- `401/403` 仍应转换成 `AuthenticationRequiredException`，这样 Ribbon 会走统一登录引导
- 普通业务校验失败如果对应某些单元格不可上传，优先转成 `SkippedChanges + Reason`
- 真正无法继续判断的系统异常可以直接抛出，Ribbon 会中止本次上传并显示错误

用户可见行为：

- 全部允许上传：确认弹窗显示实际上传数量，`BatchSave()` 收到全部变更
- 部分跳过：确认弹窗显示实际上传数量和跳过数量，确认后只提交允许上传的变更
- 全部跳过：不会弹上传确认，也不会调用 `BatchSave()`；用户会看到跳过数量和部分跳过原因
- 上传成功后的完成提示会显示实际上传数量；如果存在跳过项，会额外显示跳过数量
- 上传日志只记录实际提交且 `BatchSave()` 成功的单元格，不记录被过滤跳过的单元格

### 4.9 认证失败异常约定

当前 Ribbon Sync 的登录引导是“异常类型驱动”的，而不是“HTTP 状态码驱动”的。

对真实系统连接器，至少应统一下面几条：

- `GetProjects()` 里的 `401/403` 要转换成 `AuthenticationRequiredException("当前未登录，请先登录")`
- `BuildFieldMappingSeed()` 里的 `401/403` 也要转换成同样的异常；这样 `初始化当前表` 才会弹登录提示，而不是普通错误
- `Find()` 和 `BatchSave()` 里的 `401/403` 也要转换成同样的异常；这样 `部分下载`、`部分上传` 才会走统一的登录引导
- 异常消息不会直接作为最终用户文案展示；中文 / 英文提示由宿主 UI 语言决定
- 登录成功后，项目列表场景会自动重载项目；初始化、下载、上传不会自动重试，用户需要重新触发一次

建议直接参考当前示例连接器的处理方式：

- [src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs](../src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs)

## 5. 当前业务系统的接入模式

当前 mock / 示例系统的合同是：

- 唯一 ID 字段固定为 `row_id`
- `/head` 返回所有非活动字段和活动头
- 活动属性列通过 `/find` 返回的样本平铺字段推导
- `/batchSave` 每个 item 对应一个单元格更新

这套模式非常适合你的当前业务描述：

- ID 列名固定
- Excel 表头和接口字段的映射通过接口拉取
- 双层活动表头由活动头 + 属性字段组合出来

另外，当前项目列表也已经完全接口化：

- `GET /projects` 由连接器拉取
- Ribbon 只消费 `GetProjects()` 的返回值
- 如果真实系统后续改成别的项目聚合逻辑，只需要调整连接器，不需要改 Ribbon 控件层

## 6. 当前推荐改造路线

当前注册中心已经存在，所以接入真实系统时的推荐做法是：

1. 新增一个实现 `ISystemConnector` 的真实连接器
2. 在连接器内部封装项目列表、字段头、查询、更新的真实接口差异
3. 在 `ThisAddIn` 里把该连接器注册到 `SystemConnectorRegistry`
4. 如果要替换当前系统，就只注册新的连接器
5. 如果要并存多个系统，就同时注册多个连接器

建议新增：

- `RealBusinessSystemConnector`
- `RealBusinessFieldMappingSeedBuilder`
- 必要的 DTO / mapper

这条路线下：

- Ribbon 项目下拉框会自动聚合所有已注册连接器的项目
- 绑定到 sheet 上的是 `SystemKey + ProjectId`
- 后续下载 / 上传会自动按 `SystemKey` 找回正确连接器

## 7. 对真实系统最重要的几个约束

### 7.1 表头文本会被当作运行时事实

插件不会持久化列号。

当前上传 / 下载都依赖：

- 当前 sheet 上的表头文本
- `SheetFieldMappings` 里的 `Excel L1 / Excel L2`

这里要特别区分“识别 metadata”和“默认生成布局”：

- `single + Excel L2` 只表示这个字段在运行时要按 grouped single 的双层文本参与识别
- 它不是一个“空表头时默认生成 grouped single 布局”的信号
- 如果用户只改了 Excel 可见表头，没有同步改 `SheetFieldMappings.Excel L1 / Excel L2`，上传 / 下载不会自动识别这是同一列。用户可以手工维护这些列，也可以使用 `AI映射列` 在确认预览后批量写回。

所以如果用户改了 Excel 列名，就要同步维护 `SheetFieldMappings`；插件不会在上传 / 下载中静默探测并回写这种改动。

### 7.2 布局行号由用户控制

当前这三个值都可能被用户手工改：

- `HeaderStartRow`
- `HeaderRowCount`
- `DataStartRow`

真实系统接入时不要把它们重新写死回默认值。

同时要注意，用户现在也可能直接手工维护 `xISDP_Setting`：

- 修改 `SheetBindings` 的配置值
- 修改 `SheetFieldMappings` 的 `Excel L1 / Excel L2`
- 补充或修正映射行

对于 grouped single，真实系统要把下面两部分一起维护：

- 可见工作表上的 grouped header 文本
- `SheetFieldMappings` 里的 `Excel L1 / Excel L2`

只改 Excel 可见表头、不改 metadata，不会被自动识别成同一列。

因此真实系统接入时，不要假设 metadata 一定只会由程序生成；连接器生成的是初始化种子，不是运行时唯一写入来源。

### 7.3 已有 Excel 也要能工作

用户可能已经有带表头和数据的 Excel。

当前策略是：

- 先尝试按当前表头文本自动识别列
- 如果识别成功，就直接上传 / 下载
- 如果识别失败，再要求用户执行 `初始化当前表`

所以真实系统接入时，要确保你的映射定义足够支持“按当前表头文本反查列”。

### 7.4 `HeaderRowCount = 1` 和 `HeaderRowCount = 2` 含义不同

- `HeaderRowCount = 1`
  - 所有列只显示一行表头
  - 活动属性列只显示子属性名
- `HeaderRowCount = 2`
  - 单层列上下合并
  - 活动列第一行显示活动名，第二行显示属性名

对于 `single + Excel L2` 还要补充两点：

- 它是“按 grouped single 识别现有 Excel”的 metadata，不是“空表时自动生成 grouped single”的布局指令
- 所以空表头场景下，`全量下载` 仍会生成扁平单层表头，而不是自动补父表头

如果真实系统的表头层级更多，当前 Excel 布局服务还需要继续扩展。

### 7.5 不兼容旧 metadata 压平格式

当前初版已经明确不兼容旧 metadata 存储格式。

这意味着：

- 不需要为真实系统接入额外设计“旧 metadata 迁移逻辑”
- 初始化或后续 metadata 写入时，可以直接按当前标准 section 布局覆盖 `xISDP_Setting`
- 如果你从别的历史分支带来旧格式数据，应先清理，再按当前版本重新初始化

### 7.6 本机模板库与运行时 metadata 的边界

当前 Ribbon Sync 新增了一层本机模板资产：

- 模板资产保存在 `%LocalAppData%\OfficeAgent\templates\...`
- `xISDP_Setting.TemplateBindings` 只记录“当前 sheet 绑定到哪个模板”
- 真正参与下载、上传、初始化执行的，仍然是 `xISDP_Setting` 中展开后的 `SheetBindings + SheetFieldMappings`

因此接入真实系统时要注意：

- 不能把本机模板库当成运行时执行的唯一事实来源
- 不能跳过 `xISDP_Setting`，直接让下载上传只依赖模板引用
- 如果真实系统要扩展模板能力，应优先保证模板应用结果最终仍然回写到 `xISDP_Setting`

## 8. 真实系统落地步骤

建议按下面顺序做：

1. 明确真实系统的项目接口、表头接口、查询接口、更新接口
2. 确认唯一 ID 字段
3. 新建真实连接器和 DTO
4. 新建真实系统的 `FieldMappingSeedBuilder`
5. 让连接器先跑通 `GetProjects -> BuildFieldMappingSeed -> Find -> BatchSave`
6. 再在 `ThisAddIn` 中注册或切换连接器实例
7. 在 Excel 中执行一次 `初始化当前表`，确认 `xISDP_Setting` 被按当前标准布局写出
8. 额外验证未登录场景下，项目下拉框、`初始化当前表`、`部分下载`、`部分上传` 都能弹出当前宿主语言对应的未登录提示
9. 最后做 Excel 联调和手工回归

当前注册位置：

- [src/OfficeAgent.ExcelAddIn/ThisAddIn.cs](../src/OfficeAgent.ExcelAddIn/ThisAddIn.cs)

## 9. 最容易踩坑的点

### 9.1 日期与显示值格式

因为当前没有快照比对，日期格式问题主要影响的是：

- 下载后写到 Excel 的显示值
- 上传时读取回来的字符串值

如果真实系统要求严格格式，建议在连接器层统一做归一化。

### 9.2 活动属性列不在 `/head` 中直接返回

如果真实系统和当前一样，只返回活动头而不直接返回属性列，就必须保证：

- 样本查询能带回完整平铺字段
- 连接器能从字段名拆出 `propertyId + activityId`

### 9.3 更新接口不是按单元格设计

如果真实接口更偏“按行更新”，就在连接器里做聚合。

不要为了适配某个系统去改 Ribbon 控制器或 Excel 选区解析。

## 10. 推荐测试方案

### 单元测试

至少补：

- 连接器请求体 / 响应体映射测试
- `FieldMappingTableDefinition` 定义测试
- `BuildFieldMappingSeed` 测试
- `BatchSave` payload 测试
- 如果连接器实现 `IUploadChangeFilter`，补过滤规则测试，覆盖全部允许、部分跳过、全部跳过和跳过原因

可参考：

- [tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs](../tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs)
- [tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessFieldMappingSeedBuilderTests.cs](../tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessFieldMappingSeedBuilderTests.cs)
- [tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs](../tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs)
- [tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs](../tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs)
- [tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs](../tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs)

### 集成测试

至少补：

- `BuildFieldMappingSeed -> Find -> BatchSave` roundtrip
- 活动列 schema / mapping 生成正确

可参考：

- [tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs](../tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs)

### Excel 手工测试

至少确认：

- 选择项目后不会自动初始化；只有显式点击 `初始化当前表` 才会写入或刷新 `SheetFieldMappings`
- 对已有实际表头的 sheet，点击 `AI映射列` / `AI map columns` 后应先看到预览确认；确认后只更新 `SheetFieldMappings.Excel L1 / Excel L2`，取消后不写 metadata
- 未登录时项目下拉框显示本地化状态（`请先登录` / `Sign in first`），并弹出本地化未登录提示；点击登录按钮并登录成功后能够自动重载项目列表
- 未登录时执行 `初始化当前表`、`部分下载`、`部分上传`，都会弹出本地化未登录提示
- 项目接口返回空列表时，下拉框显示 `无可用项目`
- 显式初始化不会破坏业务单元格
- `xISDP_Setting` 会以单 sheet、上下三个 section 的可读布局写出
- 全量下载能按配置行号落位
- 已有表头场景下，全量下载不会重写已识别表头
- 部分上传 / 部分下载在不包含 ID / 表头的选区里仍能正确定位
- 如果启用了上传过滤，确认弹窗、完成提示和实际 `BatchSave()` 请求里的上传数量一致
- 如果本次上传全部被过滤跳过，应确认不会调用 `BatchSave()`，并且提示里能看到跳过原因

## 11. 最小交付标准

在你宣布“真实系统已接入”之前，建议至少满足：

1. 能选择真实项目并写入 `SheetBindings`
2. 能初始化并生成 `SheetFieldMappings`
3. 全量下载可用
4. 部分下载可用
5. 全量上传可用
6. 部分上传可用
7. 至少有一套连接器级集成测试

如果真实系统有上传过滤规则，还应额外满足：

1. 过滤规则能返回实际上传项和跳过项
2. 跳过原因能在 Ribbon 提示中被用户看懂
3. `BatchSave()` 只收到实际允许上传的单元格

## 12. 当前最建议的结论

如果你下一步只接一个真实系统，最实际的做法是：

1. 保持 `ISystemConnector` 的主边界不变
2. 新建真实系统连接器和映射种子构建器
3. 在连接器层消化真实接口差异
4. 在 `ThisAddIn` 里把它注册进 `SystemConnectorRegistry`
5. 如果只保留一个系统，就只注册这一个连接器

如果后续要并存多个系统，就继续新增连接器并一起注册，不需要重做 Ribbon Sync 主链路。
