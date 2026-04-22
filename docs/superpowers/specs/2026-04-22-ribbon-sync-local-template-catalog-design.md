# Ribbon Sync 本机模板库与 AI_Setting 展开态联动设计说明

日期：2026-04-22

状态：设计已确认，待进入实施计划

## 1. 目标

在当前 Ribbon Sync 基础上新增“按项目管理多个 `AI_Setting` 模板”的能力，使用户可以：

- 在同一个 `SystemKey + ProjectId` 下维护多套不同过滤条件或字段映射的模板
- 继续直接在 Excel 的 `AI_Setting` 中查看和编辑当前 sheet 的展开态配置
- 将当前 sheet 上已经编辑过的 `AI_Setting` 保存回原模板
- 将当前 sheet 上已经编辑过的 `AI_Setting` 另存为新模板
- 在未来保留从“本机模板库”平滑演进到“共享模板目录”的扩展边界

本次设计要同时解决两个现有问题：

- 仅靠 Excel 文件隔离 `AI_Setting` 不稳定，容易被覆盖或误改
- 纯粹只存模板引用不满足当前用户需要直接编辑 `AI_Setting` 的工作方式

## 2. 范围

### 2.1 本次范围

- 引入本机模板资产层，用于按项目保存多个模板
- 保留 `AI_Setting` 作为当前 sheet 的运行时展开态事实源
- 在 `AI_Setting` 中新增模板来源元信息 section
- 支持模板应用、保存回原模板、另存为新模板、取消模板关联
- 支持按当前项目过滤模板列表
- 为模板保存和应用增加版本冲突、兼容性校验和脏状态判断
- 为上述行为补 Core / Infrastructure / ExcelAddIn 测试和手工验证清单

### 2.2 本次明确不做

- 不做团队共享模板
- 不做服务端模板接口
- 不做模板历史浏览或版本回滚 UI
- 不做模板差异对比界面
- 不做用户编辑 `AI_Setting` 后的自动即时回写模板
- 不新增 task pane 交互

## 3. 总体方案选择

本次采用：

- `本机模板库`
- `AI_Setting 展开态继续可见可编辑`
- `TemplateBindings 记录模板来源元信息`

含义如下：

- 当前 sheet 真正参与下载、上传、初始化的仍然是 `AI_Setting` 中展开后的 `SheetBindings + SheetFieldMappings`
- 模板库独立保存在本机，不再依赖 Excel 文件承载全部模板资产
- `AI_Setting` 中额外保存“当前展开态来自哪个模板”的元信息，以支持保存回原模板和另存为新模板

本次明确不采用以下两条路线：

- 不把模板库继续直接塞进 `AI_Setting`，避免把“运行时状态”和“模板资产”混成一层
- 不只在 workbook 中保留模板引用，避免用户无法方便地直接编辑当前生效配置

## 4. 核心术语

- `模板资产`
  - 保存在本机模板库中的一套可复用配置
  - 与具体 `SheetName` 脱钩
- `展开态`
  - 当前 sheet 在 `AI_Setting` 中实际生效的 `SheetBindings + SheetFieldMappings`
- `模板来源元信息`
  - 当前展开态和哪个模板有关、是否已偏离模板、来自哪个分叉
- `归一化模板快照`
  - 从当前 sheet 的展开态中提取出来、去掉 `SheetName` 后得到的可持久化模板结构

## 5. 存储边界

本次设计把模板和当前 sheet 配置拆成三层。

### 5.1 `AI_Setting` 仍然是运行时事实源

当前 `AI_Setting` 的定位不变：

- `SheetBindings` 继续描述当前 sheet 的项目绑定和布局参数
- `SheetFieldMappings` 继续描述当前 sheet 的字段映射工作副本
- 下载、上传、初始化继续只认当前 workbook 中这份展开态

这意味着：

- 用户仍然可以直接手工维护 `AI_Setting`
- 业务执行层不直接读取本机模板库作为运行时事实源
- 模板应用的结果最终一定会落回 `AI_Setting`

### 5.2 本机模板库承载模板资产

新增 `TemplateStore` 作为模板资产持久化层。

推荐本地存储路径：

```text
%LocalAppData%\OfficeAgent\templates\<systemKey>\<projectId>\<templateId>.json
```

存储原则：

- 一条模板资产对应一个独立文件
- 模板以 `SystemKey + ProjectId` 为主过滤维度
- 模板内容不依赖具体 workbook
- 模板内容不持久化具体 `SheetName`

### 5.3 `AI_Setting` 新增 `TemplateBindings` section

`AI_Setting` 从当前的两个 section 扩展为三个可读区域：

- `TemplateBindings`
- `SheetBindings`
- `SheetFieldMappings`

展示布局保持当前模式：

- 每个区域包含一行标题、一行表头、多行数据
- 区域之间固定留两行空白分隔

推荐顺序：

1. `TemplateBindings`
2. `SheetBindings`
3. `SheetFieldMappings`

这样用户先看到“当前 sheet 与模板的关系”，再看到实际展开后的运行时配置。

## 6. 数据模型

### 6.1 `TemplateDefinition`

模板资产模型建议至少包含：

- `TemplateId`
- `TemplateName`
- `SystemKey`
- `ProjectId`
- `ProjectName`
- `HeaderStartRow`
- `HeaderRowCount`
- `DataStartRow`
- `FieldMappingDefinition`
- `FieldMappingDefinitionFingerprint`
- `FieldMappings`
- `Revision`
- `CreatedAt`
- `UpdatedAt`

规则：

- `FieldMappings` 保存模板级结构，不持久化具体 `SheetName`
- `FieldMappingDefinition` 保存列定义快照，避免未来 connector 定义变化时失去兼容判断依据

### 6.2 `TemplateBinding`

`AI_Setting.TemplateBindings` 一行对应一个业务 sheet，建议字段至少包括：

- `SheetName`
- `TemplateId`
- `TemplateName`
- `TemplateRevision`
- `TemplateOrigin`
- `AppliedFingerprint`
- `TemplateLastAppliedAt`
- `DerivedFromTemplateId`
- `DerivedFromTemplateRevision`

字段语义：

- `TemplateOrigin`
  - `store-template`
    - 当前 sheet 已绑定到本机模板库中的正式模板
  - `ad-hoc`
    - 当前 sheet 是没有固定模板来源的临时工作副本
- `AppliedFingerprint`
  - 当前 sheet 最近一次“应用模板”或“保存模板”后对应的归一化指纹
- `DerivedFrom...`
  - 记录当前正式模板最初是从哪个模板分叉出来
  - 只表达派生关系，不影响当前“保存回原模板”的目标

### 6.3 `TemplateOrigin` 规则

`TemplateOrigin` 的推荐规则如下：

- 首次从模板应用到当前 sheet 后，写为 `store-template`
- 用户把当前 sheet 另存为新模板成功后，写为 `store-template`
- 用户执行取消模板关联后，写为 `ad-hoc`
- 旧 workbook 没有 `TemplateBindings` 时，运行时视为 `ad-hoc`

说明：

- `另存为新模板` 后不能保留为 `ad-hoc`
- 因为此时当前 sheet 已明确绑定到模板库中真实存在的新模板

### 6.4 归一化模板快照

从当前 sheet 提取模板时，必须执行归一化：

- 去掉具体 `SheetName`
- 保留 `SystemKey + ProjectId + ProjectName`
- 保留布局参数
- 保留字段映射定义和字段映射行
- 不把 `TemplateBindings` 自身的元信息再写回模板内容

归一化后的结果用于：

- 保存回原模板
- 另存为新模板
- 生成 `AppliedFingerprint`
- 进行“当前展开态是否已偏离模板”的判断

## 7. 接口边界

### 7.1 `ITemplateStore`

`TemplateStore` 只负责持久化，不直接承担 UI 语义。

建议接口能力：

- `ListByProject(systemKey, projectId)`
- `Get(templateId)`
- `SaveNew(templateDefinition)`
- `Update(templateDefinition, expectedRevision)`
- `Delete(templateId)`  
  本次可以先不开放 UI，但保留接口边界

### 7.2 `ITemplateCatalog`

`TemplateCatalog` 负责模板流程编排。

建议接口能力：

- `ListTemplates(systemKey, projectId)`
- `ApplyTemplateToSheet(sheetName, templateId)`
- `SaveSheetToExistingTemplate(sheetName, templateId, expectedRevision)`
- `SaveSheetAsNewTemplate(sheetName, templateName)`
- `DetachTemplate(sheetName)`
- `GetSheetTemplateState(sheetName)`

`TemplateCatalog` 负责：

- 从 `AI_Setting` 提取归一化模板快照
- 把模板资产展开写回当前 sheet
- 管理 revision 检查
- 管理模板兼容性校验
- 计算 dirty 状态

### 7.3 与现有 `WorksheetMetadataStore` 的边界

`WorksheetMetadataStore` 继续只负责 workbook 内 `AI_Setting`。

需要扩展的能力：

- `SaveTemplateBinding(...)`
- `LoadTemplateBinding(sheetName)`
- `ClearTemplateBinding(sheetName)`

不建议把模板生命周期字段直接并入现有 `SheetBinding` 模型，原因如下：

- `SheetBinding` 当前只表达运行时布局和项目绑定
- 把模板来源、修订号、派生关系直接混入 `SheetBinding` 会让同步执行层承担不相关职责

## 8. `AI_Setting` 布局规则

### 8.1 `TemplateBindings`

区域布局示意：

```text
TemplateBindings
SheetName | TemplateId | TemplateName | TemplateRevision | TemplateOrigin | AppliedFingerprint | TemplateLastAppliedAt | DerivedFromTemplateId | DerivedFromTemplateRevision
Sheet1    | tpl-a      | 条件A        | 3                | store-template | abc123             | 2026-04-22T10:00:00   |                        |
Sheet2    |            |              |                  | ad-hoc         |                    |                       |                        |
```

规则：

- 一行对应一张业务 sheet
- 当前 sheet 无模板来源时，允许显式保存为 `ad-hoc`
- `TemplateId` 为空但 `TemplateOrigin = ad-hoc` 是合法状态

### 8.2 `SheetBindings`

`SheetBindings` 仍然保持当前列定义：

- `SheetName`
- `SystemKey`
- `ProjectId`
- `ProjectName`
- `HeaderStartRow`
- `HeaderRowCount`
- `DataStartRow`

本次不在 `SheetBindings` 中新增模板字段。

### 8.3 `SheetFieldMappings`

`SheetFieldMappings` 的列结构仍由 connector 提供的 `FieldMappingTableDefinition` 决定。

模板化保存时必须同时保存：

- 当前字段映射行
- 对应的列定义快照

## 9. 用户流程

### 9.1 从模板应用到当前表

流程如下：

1. 用户先选择项目，或当前 sheet 已存在项目绑定
2. 用户点击 `应用模板`
3. 系统按当前 `SystemKey + ProjectId` 过滤模板列表
4. 用户选择模板并确认
5. 系统校验模板与当前 connector 定义兼容
6. 系统把模板展开写入当前 sheet 的：
   - `TemplateBindings`
   - `SheetBindings`
   - `SheetFieldMappings`

应用完成后：

- 当前 sheet 进入“来自正式模板的工作副本”状态
- `TemplateOrigin = store-template`
- `TemplateRevision` 写入被应用模板的当前 revision
- `AppliedFingerprint` 更新为本次展开态的归一化指纹

### 9.2 直接编辑 `AI_Setting`

当前工作方式保持不变：

- 用户可以直接改 `SheetBindings`
- 用户可以直接改 `SheetFieldMappings`
- 这些修改先只作用于当前 workbook 中的展开态
- 不自动即时回写模板库

### 9.3 保存回原模板

按钮 `保存模板` 只在以下条件满足时启用：

- 当前 sheet 的 `TemplateOrigin = store-template`
- `TemplateId` 存在
- 当前 sheet 的 `SystemKey + ProjectId` 与目标模板一致

流程如下：

1. 从当前 sheet 提取归一化模板快照
2. 读取本地模板库中对应 `TemplateId`
3. 检查 `expectedRevision == currentRevision`
4. 校验模板所属项目与当前 sheet 一致
5. 覆盖保存模板，模板 `Revision + 1`
6. 回写当前 sheet 的 `TemplateBindings`

保存成功后：

- `TemplateRevision` 更新为新 revision
- `AppliedFingerprint` 更新为当前快照指纹
- `TemplateOrigin` 保持 `store-template`

### 9.4 另存为新模板

按钮 `另存模板` 始终可用。

流程如下：

1. 用户输入新模板名
2. 系统从当前 sheet 提取归一化模板快照
3. 新建新的 `TemplateId`
4. 写入本机模板库
5. 回写当前 sheet 的 `TemplateBindings`

保存成功后：

- `TemplateId` 切换为新模板
- `TemplateName` 切换为新模板名
- `TemplateRevision` 设为新模板当前版本
- `TemplateOrigin = store-template`
- `AppliedFingerprint` 更新为当前快照指纹
- 如果是从已有模板分叉，回写 `DerivedFromTemplateId + DerivedFromTemplateRevision`

### 9.5 取消模板关联

允许用户将当前 sheet 从模板绑定状态切回临时工作副本。

执行后：

- 保留 `SheetBindings + SheetFieldMappings` 不变
- `TemplateOrigin` 置为 `ad-hoc`
- 清空当前目标模板标识字段
- 后续只能“另存为新模板”，不能“保存回原模板”

## 10. Dirty 与指纹规则

本次不要求在用户每次改单元格时即时重写 dirty 标记。

改用“按需重算”的策略：

- 打开模板对话框时
- 执行应用模板前
- 执行保存模板前
- 执行另存模板前

系统从当前 sheet 提取归一化模板快照，生成 `fingerprint`，并与 `AppliedFingerprint` 比较：

- 一致：视为未偏离上次模板应用或保存后的状态
- 不一致：视为当前工作副本已修改

指纹计算必须忽略以下内容：

- `SheetName`
- `TemplateBindings` 元信息字段
- 与当前 workbook 实例相关但不属于模板语义的值

## 11. 冲突与错误处理

### 11.1 应用模板前的覆盖确认

如果当前 sheet 已偏离模板或已存在未保存工作副本：

- 再次应用模板前必须弹确认
- 防止用户误覆盖当前展开态

### 11.2 模板 revision 冲突

如果用户执行“保存回原模板”时，本机模板库里的 revision 已不是当前 sheet 记录的版本：

- 弹框提示模板已被更新
- 提供两个选择：
  - 覆盖原模板
  - 另存为新模板

### 11.3 模板缺失

如果 `TemplateBinding.TemplateId` 指向的模板在本机模板库中已不存在：

- 当前 sheet 的展开态继续可用
- `保存模板` 置灰或给出明确错误
- 允许用户执行“另存为新模板”

### 11.4 项目不一致

如果当前 sheet 的 `SystemKey + ProjectId` 与模板不一致：

- 禁止“保存回原模板”
- 允许“另存模板”

### 11.5 字段定义不兼容

应用模板前必须比较：

- 模板保存时的 `FieldMappingDefinition`
- 当前 connector 返回的定义

处理规则：

- 完全一致：直接应用
- 轻微差异且可按列名对齐：提示后降级应用
- 差异过大：阻止直接应用，并提示用户重新初始化或另存新模板

## 12. Ribbon 与 UI 设计

本次推荐新增单独的模板操作组，而不是新增常驻模板下拉框。

推荐按钮：

- `应用模板`
- `保存模板`
- `另存模板`

原因：

- 当前 Ribbon 已有项目下拉框，模板状态再做成第二个常驻 dropdown，会显著增加状态同步复杂度
- 模板列表天然需要按项目过滤，并可能展示更多元信息，使用按钮 + 对话框更适合

控制器边界建议：

- `RibbonSyncController`
  - 继续只管理项目选择、初始化、下载、上传
- 新增 `RibbonTemplateController`
  - 管理模板应用、保存、另存和状态判断

## 13. 迁移与兼容性

本次设计必须兼容已有 workbook。

迁移规则：

- 旧 workbook 没有 `TemplateBindings` 时，运行时默认视为 `ad-hoc`
- 旧 workbook 的下载、上传、初始化行为保持不变
- 用户第一次执行模板相关操作时，再写入 `TemplateBindings`

模板功能上线后，`AI_Setting` 的最终可见布局会从“双 section”演进到“三 section”。

相关文档在实施时需要同步更新：

- `docs/modules/ribbon-sync-current-behavior.md`
- `docs/ribbon-sync-real-system-integration-guide.md`
- `docs/vsto-manual-test-checklist.md`

## 14. 测试策略

### 14.1 Core

需要覆盖：

- 归一化模板快照提取
- `AppliedFingerprint` 计算和 dirty 判断
- 应用模板、保存回模板、另存模板、取消模板关联
- revision 冲突和项目不一致判断

### 14.2 Infrastructure

需要覆盖：

- 本机 JSON 模板文件读写
- 按 `SystemKey + ProjectId` 过滤模板
- `Revision` 增长与乐观并发检查

### 14.3 ExcelAddIn

需要覆盖：

- `TemplateBindings` 的 metadata roundtrip
- Ribbon 按钮启用状态
- 应用模板后 `AI_Setting` 三个 section 的写入结果
- 模板缺失、项目不一致、revision 冲突时的提示和回退

### 14.4 手工验证

至少覆盖以下场景：

- 从模板应用后立即执行初始化、下载、上传仍然正常
- 手工编辑 `AI_Setting` 后可保存回原模板
- 手工编辑 `AI_Setting` 后可另存为新模板
- 打开旧 workbook 时，不带模板元信息也能正常同步

## 15. 第一阶段收敛范围

为控制首次落地复杂度，第一阶段只实现：

- 本机 JSON 模板库
- `TemplateBindings` section
- `应用模板 / 保存模板 / 另存模板`
- 基础 dirty 判断
- revision 冲突提示

第一阶段不实现：

- 模板删除 UI
- 模板 diff UI
- 模板历史列表
- 团队共享模板

## 16. 结论

本次设计采用“本机模板库 + `AI_Setting` 展开态 + `TemplateBindings` 元信息”的混合模型。

这样可以同时满足：

- 模板资产不再只依附 Excel 文件
- 用户仍然可以方便地直接编辑 `AI_Setting`
- 当前工作副本能够保存回原模板或分叉另存为新模板
- 后续可在不推翻 Ribbon Sync 主链路的前提下演进到共享模板目录
