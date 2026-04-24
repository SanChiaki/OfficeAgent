# Ribbon Sync Mock Server Large Activity Benchmark Design

日期：2026-04-23

状态：设计已确认，待实施

## 1. 目标

为 `tests/mock-server` 新增一个专门用于 Ribbon Sync 性能联调的 mock 项目，满足以下目标：

- 数据量达到万级，默认 `10000` 行
- 保留 activity 列语义，用于覆盖双行表头初始化与写入性能
- 每行包含 `row_id + 10` 个业务字段
- 不改变当前 `GET /projects`、`POST /head`、`POST /find`、`POST /batchSave` 的契约
- 保持 mock server 仍然使用进程内存数据，不引入数据库或额外服务

## 2. 背景

当前 `tests/mock-server/server.js` 只内置了几个小数据量项目：

- `performance`
- `delivery-tracker`
- `customer-onboarding`

这些项目足够覆盖功能联调，但不足以暴露以下性能问题：

- 双行表头初始化时的 Excel 写入耗时
- 大行数下载时的数据区写入耗时
- 大行数下字段映射种子生成、模式识别和表头布局生成耗时

因此需要一个专门的“大数据 + activity 双行表头”项目。

## 3. 范围

### 3.1 本次要做

- 在 `tests/mock-server/server.js` 中新增一个大数据量 Ribbon Sync 项目
- 让该项目出现在 `/projects`
- 让 `/head` 返回普通列 + activity 头
- 让 `/find` 返回 `10000` 行扁平数据
- 让 `/batchSave` 能对该项目按现有逻辑写回内存数据
- 为该项目补集成测试
- 更新 `tests/mock-server/README.md`

### 3.2 本次不做

- 不修改 Ribbon Sync 的连接器协议
- 不修改 Excel Add-in 的表头识别逻辑
- 不引入新的分页、流式下载或懒加载协议
- 不引入落盘持久化

## 4. 方案对比

### 4.1 方案一：静态写死 10000 行 JSON

优点：

- 最直观

缺点：

- `server.js` 或数据文件会迅速膨胀
- 可读性和维护性差
- 很难继续调整字段模式

### 4.2 方案二：启动时程序化生成大数据项目

优点：

- 文件体积可控
- 可稳定生成相同规模和结构的数据
- 最符合当前 mock server “内存种子”模式

缺点：

- 需要额外的生成函数

### 4.3 方案三：外置大数据文件并在启动时加载

优点：

- 服务逻辑和数据文件分离

缺点：

- 仓库里仍会增加大文件
- 对当前目的没有明显收益

### 4.4 结论

采用方案二：启动时程序化生成。

## 5. 推荐设计

### 5.1 新项目结构

新增项目：

- `projectId = large-activity-benchmark`
- `displayName = 大数据活动压测项目`

默认行数：

- `10000`

每行字段共 `11` 个：

- `row_id`
- 普通字段 `4` 个：
  - `owner_name`
  - `region`
  - `priority`
  - `status`
- activity 字段 `6` 个：
  - `name_benchmarka`
  - `start_benchmarka`
  - `end_benchmarka`
  - `name_benchmarkb`
  - `start_benchmarkb`
  - `end_benchmarkb`

这样业务字段总数是 `10`，同时具备两个 activity 分组，能真实覆盖双行表头写入。

### 5.2 `/head` 返回结构

`headList` 需要返回：

- `row_id`
- `owner_name`
- `region`
- `priority`
- `status`
- activity 头 `benchmarka`
- activity 头 `benchmarkb`

其中 activity 头只返回：

- `headType = activity`
- `activityId`
- `activityName`

不直接列出 `name/start/end` 子列，继续沿用当前通过 `/find` 样本行自动推导 activity property 列的机制。

### 5.3 `/find` 返回结构

`/find` 的行对象保持扁平 JSON：

```json
{
  "row_id": "benchmark-row-00001",
  "owner_name": "区域负责人001",
  "region": "华东",
  "priority": "P1",
  "status": "进行中",
  "name_benchmarka": "阶段A-001",
  "start_benchmarka": "2026-01-01",
  "end_benchmarka": "2026-01-03",
  "name_benchmarkb": "阶段B-001",
  "start_benchmarkb": "2026-01-04",
  "end_benchmarkb": "2026-01-06"
}
```

这允许：

- `CurrentBusinessFieldMappingSeedBuilder` 自动推导 6 个 activity property 字段
- `CurrentBusinessSchemaMapper` 自动推导双行表头列

### 5.4 数据生成规则

生成逻辑要求：

- `row_id` 稳定且可读，如 `benchmark-row-00001`
- 普通列值在固定候选集中循环，保证样本稳定
- 日期字段按行号偏移生成，避免所有行都重复同一天
- 两个 activity 组都包含 `name/start/end`

重点不是生成“复杂业务语义”，而是稳定地提供：

- 足够大的行数
- 足够多的列
- 明确的双行表头结构

## 6. 测试策略

### 6.1 集成测试

新增或扩展现有 `CurrentBusinessSystemConnectorIntegrationTests`，至少覆盖：

- `/projects` 能返回新项目
- 新项目 `/find` 默认返回 `10000` 行
- 每行包含 `11` 个字段
- 新项目 schema 中能识别出：
  - `4` 个普通业务列
  - `6` 个 activity property 列

### 6.2 README

更新 `tests/mock-server/README.md`，明确说明：

- 新增的大数据项目名称
- 行数规模
- activity 分组数量
- 字段结构为 `row_id + 10` 个业务字段

## 7. 风险与约束

### 7.1 mock server 启动时间增加

启动时生成 `10000` 行对象会带来少量启动耗时，但在 Node.js 进程内可接受。

### 7.2 文档与测试断言失配

现有集成测试明确断言当前项目数为 `3`，新增项目后必须同步调整，否则测试会失败。

### 7.3 property label 显示名边界

当前连接器内置的 property label 只有：

- `name`
- `start`
- `end`

因此 activity 子列应继续使用这三个 property key，避免需要改连接器实现。

## 8. 结论

本次改动应在不改变 Ribbon Sync 协议的前提下，向 mock server 增加一个“`10000` 行、`2` 个 activity 组、`row_id + 10` 个业务字段”的性能测试项目，作为双行表头初始化和大数据量下载/上传联调基准数据集。
