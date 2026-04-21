# OfficeAgent 架构对比图设计说明

日期：2026-04-21

状态：设计已确认，待用户审阅 spec

## 1. 目标

为当前 OfficeAgent 插件补一组可直接用于文档和评审沟通的技术架构图，并满足以下目标：

- 同时产出 3 张可横向对比的架构图
- 三张图使用统一视觉语言，方便比较信息边界和主次关系
- 明确当前插件存在两条主能力链：
  - `Ribbon Sync`
  - `Task Pane AI`
- 明确 `SystemConnector` 当前只进入 `Ribbon Sync` 主链路
- 明确 `AI_Setting / SheetBindings / SheetFieldMappings` 是连接器与 Excel 表结构之间的桥接层
- 让图中节点名称与代码真实命名保持一致，便于从图回到代码

## 2. 范围

### 2.1 本次范围

- 产出 3 张独立技术架构图
- 每张图导出 `SVG + PNG`
- 图中覆盖以下实际模块边界：
  - `OfficeAgent.ExcelAddIn`
  - `OfficeAgent.Core`
  - `OfficeAgent.Infrastructure`
  - `OfficeAgent.Frontend`
  - 外部业务系统、LLM、SSO
- 重点表现 `SystemConnector` 与插件功能的连接方式
- 在第 3 张图中显式对比 `Ribbon Sync` 与 `Task Pane AI` 的不同业务边界

### 2.2 本次明确不做

- 不修改现有业务代码或同步逻辑
- 不改模块命名或重构实际代码结构
- 不补新的运行时埋点或自动采图逻辑
- 不把 3 张图做成一张拼盘主板
- 不新增与本次架构理解无关的实现细节

## 3. 总体交付方式

本次采用“统一风格三连图”的交付方式。

最终产物为 3 张独立图：

1. `office-agent-architecture-overview`
2. `office-agent-systemconnector-detail`
3. `office-agent-systemconnector-with-taskpane`

这样做的原因是：

- 图 1 负责全景定位
- 图 2 负责把 `SystemConnector` 讲透
- 图 3 负责把 `SystemConnector` 主链路与 `Task Pane AI` 平行链路并列展示

相比“混合图型三连图”或“拼盘式大图”，该方案更适合横向对比，也更适合作为后续文档的稳定资产。

## 4. 三张图的边界定义

### 4.1 图 1：整体插件全景架构图

文件名：

- `office-agent-architecture-overview.svg`
- `office-agent-architecture-overview.png`

目标：

- 用一张图看清 OfficeAgent 的整体分层和两条主能力链

覆盖边界：

- `Excel 用户`
- `Ribbon / Task Pane UI`
- `OfficeAgent.ExcelAddIn`
- `OfficeAgent.Core`
- `OfficeAgent.Infrastructure`
- `外部系统`

必须表达的内容：

- `Ribbon Sync` 和 `Task Pane AI` 是两条平行主链路
- `ThisAddIn` 是组合根
- `SystemConnector` 会被高亮，但不展开到每个接口级别
- `SystemConnector` 当前只进入 `Ribbon Sync`，不直接进入任务窗格 Agent

这一张图优先保证：

- 全景可读
- 层次清晰
- 关键边界一眼可见

### 4.2 图 2：SystemConnector 细节链路图

文件名：

- `office-agent-systemconnector-detail.svg`
- `office-agent-systemconnector-detail.png`

目标：

- 解释 `SystemConnector` 如何真正连接到 Ribbon Sync 的项目选择、初始化、下载、上传

覆盖边界：

- `Ribbon`
- `RibbonSyncController`
- `WorksheetSyncExecutionService`
- `WorksheetSyncService`
- `SystemConnectorRegistry`
- `ISystemConnector`
- `CurrentBusinessSystemConnector`
- 外部业务接口：
  - `/projects`
  - `/head`
  - `/find`
  - `/batchSave`
- 元数据桥接：
  - `AI_Setting`
  - `SheetBindings`
  - `SheetFieldMappings`

必须表达的内容：

- 项目列表通过 `GetProjects()` 进入 Ribbon
- 项目选择通过 `CreateBindingSeed()` 形成 binding seed
- 初始化通过 `GetFieldMappingDefinition()` 和 `BuildFieldMappingSeed()` 写入映射
- 下载通过 `Find()` 完成
- 上传通过 `BatchSave()` 完成
- `AI_Setting` 不是旁支，而是主链路中的桥接事实来源

这一张图优先保证：

- 主链路完整
- `SystemConnector` 职责边界明确
- Excel 元数据桥接位置准确

### 4.3 图 3：SystemConnector 与 Task Pane AI 并行关系图

文件名：

- `office-agent-systemconnector-with-taskpane.svg`
- `office-agent-systemconnector-with-taskpane.png`

目标：

- 在同一张图里比较 `SystemConnector` 主链路和 `Task Pane AI` 平行链路

覆盖边界：

- `Ribbon Sync` 主链路
- `Task Pane AI` 主链路
- 共享宿主与共享基础设施：
  - `ThisAddIn`
  - `SettingsStore / CookieStore / SharedCookies`
  - `ExcelInteropAdapter`

必须表达的内容：

- `Ribbon Sync` 通过 `SystemConnector` 访问业务系统
- `Task Pane AI` 通过 `WebView2 -> WebMessageRouter -> AgentOrchestrator` 进入 Agent 链路
- `Task Pane AI` 的上传能力走 `UploadDataSkill -> BusinessApiClient`
- `Task Pane AI` 不直接依赖 `ISystemConnector`
- 两条链路共享宿主和设置，但业务边界不同

这一张图优先保证：

- 对比关系强
- 主次明确
- 不退化成另一张全景图

## 5. 统一视觉规则

### 5.1 风格

三张图统一采用浅色技术架构图风格：

- 白底
- 轻量阴影
- 圆角矩形为主
- 统一容器框表达层次

适配场景：

- 仓库文档
- 设计说明
- 评审材料
- 汇报页面

### 5.2 容器分层

统一使用分层容器表达以下结构：

- 宿主层
- Core 层
- Infrastructure 层
- 外部系统层

容器样式统一为：

- 浅灰描边
- 轻背景填充
- 标题固定放在左上

### 5.3 节点命名

节点优先使用代码中的真实名称，例如：

- `ThisAddIn`
- `RibbonSyncController`
- `WorksheetSyncExecutionService`
- `WorksheetSyncService`
- `SystemConnectorRegistry`
- `CurrentBusinessSystemConnector`
- `AgentOrchestrator`
- `ExcelInteropAdapter`

规则：

- 主名字不改写
- 必要时只增加小号副标题
- 不把代码名翻译成与源码不一致的自然语言名

## 6. 颜色和箭头语义

### 6.1 颜色语义

- `SystemConnector` 相关节点：橙色强调
- `Ribbon Sync` 主调用链：蓝色
- `Task Pane AI` 链路：绿色
- `AI_Setting / SheetBindings / SheetFieldMappings`：青色
- 外部系统与 API：中性灰或浅紫

### 6.2 箭头语义

- 实线蓝箭头：主调用链
- 绿色或青色辅助箭头：元数据支撑关系
- 橙色强调边：`SystemConnector` 参与的关键边界
- 必要时使用虚线表示支撑、映射或非主调用关系

### 6.3 图例规则

- 图 1 和图 3 带简版图例
- 图 2 带更完整图例，因为其最依赖边语义和角色区分

## 7. 关键架构结论

三张图都必须稳定表达以下结论：

### 7.1 当前插件存在两条主能力链

- `Ribbon Sync`
- `Task Pane AI`

### 7.2 `SystemConnector` 的当前定位

`SystemConnector` 是 `Ribbon Sync` 的业务系统适配边界，而不是整个插件的统一业务访问入口。

### 7.3 元数据桥接层的定位

`AI_Setting` 中的：

- `SheetBindings`
- `SheetFieldMappings`

是连接器能力与 Excel 表头/字段识别之间的桥接事实来源。

### 7.4 Task Pane AI 的边界

任务窗格 Agent 当前走独立链路：

- `React/WebView2`
- `WebMessageRouter`
- `AgentOrchestrator`
- `UploadDataSkill`
- `BusinessApiClient`

该链路不直接依赖 `ISystemConnector`。

## 8. 文件输出约定

建议输出目录：

- `docs/architecture/diagrams/`

最终计划输出：

- `docs/architecture/diagrams/office-agent-architecture-overview.svg`
- `docs/architecture/diagrams/office-agent-architecture-overview.png`
- `docs/architecture/diagrams/office-agent-systemconnector-detail.svg`
- `docs/architecture/diagrams/office-agent-systemconnector-detail.png`
- `docs/architecture/diagrams/office-agent-systemconnector-with-taskpane.svg`
- `docs/architecture/diagrams/office-agent-systemconnector-with-taskpane.png`

规则：

- `SVG` 为源文件和文档嵌入文件
- `PNG` 为直接查看和沟通文件
- 三张图标题统一放左上角

## 9. 验证标准

每张图生成后都必须通过以下检查：

### 9.1 渲染检查

- `SVG` 文件语法合法
- `PNG` 可成功导出

### 9.2 布局检查

- 文本不溢出节点
- 箭头不穿过节点主体
- 同层节点间距稳定
- 关键节点高亮不互相抢焦点

### 9.3 语义检查

- 图 1 能一眼看出插件全景和双链路
- 图 2 能一眼看出 `SystemConnector` 如何接上项目、初始化、下载、上传
- 图 3 能一眼看出 `Ribbon Sync` 与 `Task Pane AI` 的不同边界
- `AI_Setting / SheetBindings / SheetFieldMappings` 在图 2 中位置准确

## 10. 实施顺序

后续正式出图时按以下顺序执行：

1. 先生成图 1，确认整体视觉语言成立
2. 再生成图 2，确认 `SystemConnector` 细节链路表达完整
3. 最后生成图 3，确认并行关系表达清晰
4. 对三张图统一做导出和校验

## 11. 结论

本次最终采用的方案是：

- 生成 3 张同风格对比图
- 分别覆盖：
  - 全景架构
  - `SystemConnector` 细节链路
  - `SystemConnector` 与 `Task Pane AI` 的并行关系
- 在三张图中统一高亮：
  - `SystemConnector`
  - `AI_Setting` 元数据桥接
  - `Ribbon Sync` 与 `Task Pane AI` 的边界差异

这组图的核心价值是：

- 既能解释整体插件结构
- 又能把 `SystemConnector` 的实际位置讲清楚
- 同时避免把任务窗格 Agent 与 Ribbon Sync 混成一条错误主链路
