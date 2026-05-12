# Mock Server

`tests/mock-server` 提供本地联调用的 Node.js mock 服务。

当前会同时启动两类服务：

- SSO 登录服务：`http://localhost:3100`
- 业务 API 服务：`http://localhost:3200`

服务入口脚本：

- [server.js](/D:/Workspace/demos/office-agent/.worktrees/ribbon-sync/tests/mock-server/server.js)

## 手工启动

在仓库根目录执行：

```powershell
cd tests/mock-server
npm install
npm start
```

`npm start` 实际执行：

```powershell
node server.js
```

启动成功后，控制台会输出推荐的插件配置项。

## 插件联调配置

本地联调 Excel 插件时建议使用：

- `Base URL = 你的大模型服务地址`
- `Business Base URL = http://localhost:3200`
- `Analytics Base URL = http://localhost:3200`
- `SSO URL = http://localhost:3100/login`
- `登录成功路径 = /rest/login`
- `API Key = 留空`

说明：

- `Base URL` 只用于大模型 / Agent，不用于 Ribbon Sync 业务接口
- `Business Base URL` 才是 `/head`、`/find`、`/batchSave` 和 `upload_data` 等业务接口的基地址
- `Business Base URL` 也是 `/projects` 项目列表接口的基地址
- `Analytics Base URL` 是 `/insertLog` 埋点接口的基地址
- 当前 mock 服务通过 SSO cookie 鉴权，不走 API Key
- 业务接口在未登录时会返回 `401`

## 当前接口

### SSO 服务

- `GET /login`
  - 返回登录页
- `POST /rest/login`
  - 提交用户名密码
  - 返回登录成功响应并写入 cookie

### 通用业务 API

- `GET /logged-in`
  - 返回登录成功提示页
- `GET /api/performance`
  - 返回绩效示例数据
- `GET /api/performance/:name`
  - 按姓名读取绩效示例数据
- `POST /api/performance`
  - 新增或更新绩效示例数据
- `POST /upload_data`
  - 原有上传演示接口
- `GET /api/download/:projectName`
  - 原有下载演示接口

### Analytics / 埋点接口

#### `POST /insertLog`

模拟内网埋点接口。请求体固定字段为 `frontEndIntent = excelAi`、`clientSource = Excel`、`questionType = 1`，`answer` 必须是可解析 JSON 字符串。

#### `GET /analytics/logs`

返回当前进程内最近 500 条埋点记录，包含原始外层 payload 和解析后的 `answer`。

#### `DELETE /analytics/logs`

清空当前进程内埋点记录。

### Ribbon Sync 相关接口

#### `GET /projects`

用于：

- Ribbon 项目下拉框加载

当前示例返回：

```json
[
  {
    "projectId": "performance",
    "displayName": "绩效项目"
  },
  {
    "projectId": "delivery-tracker",
    "displayName": "交付跟踪项目"
  },
  {
    "projectId": "customer-onboarding",
    "displayName": "客户上线项目"
  },
  {
    "projectId": "large-activity-benchmark",
    "displayName": "大数据活动压测项目"
  }
]
```

#### `POST /head`

返回：

- `headList`

当前约定：

- 按请求体 `projectId` 返回对应项目的字段头
- 返回所有非活动字段头
- 活动列只返回活动头
- 活动属性字段本身不在 `/head` 中单独列出

当前示例：

```json
{
  "headList": [
    { "fieldKey": "row_id", "headerText": "ID", "headType": "single", "isId": true },
    { "fieldKey": "owner_name", "headerText": "负责人", "headType": "single" },
    { "headType": "activity", "activityId": "12345678", "activityName": "测试活动111" }
  ]
}
```

#### `POST /find`

同时用于：

- 全量下载
- 部分下载
- `BuildFieldMappingSeed` 的样本数据获取

请求体支持：

- `projectId`
- `ids`
- `rowIds`
- `fieldKeys`

约定：

- 按请求体 `projectId` 返回对应项目的数据集
- `ids` / `rowIds` 为空时返回全量数据
- `fieldKeys` 为空时返回整行平铺数据
- 每行的唯一 ID 字段是 `row_id`

当前示例返回：

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

#### `POST /batchSave`

用于：

- 全量上传
- 部分上传

请求体是一个 list，每个 item 对应一个单元格改动。

当前字段：

- `projectId`
- `id`
- `fieldKey`
- `value`

当前实现会按每个 item 的 `projectId` 选择目标项目，再把 `id` 映射到行的 `row_id`。

## 当前内置数据

Ribbon Sync mock 数据保存在 [server.js](/D:/Workspace/demos/office-agent/.worktrees/ribbon-sync/tests/mock-server/server.js) 的内存变量中，主要包括：

- `connectorProjectData`
  - 按 `projectId` 组织的项目、字段头和数据行
- `connectorProjects`
  - `/projects` 返回的项目列表

当前内置项目：

- `performance`
  - `displayName = 绩效项目`
  - `2` 条示例数据
  - `activityId = 12345678`
  - `activityName = 测试活动111`
- `delivery-tracker`
  - `displayName = 交付跟踪项目`
  - `10` 条示例数据
  - `activityId = 22334455`
  - `activityName = 交付阶段`
- `customer-onboarding`
  - `displayName = 客户上线项目`
  - `10` 条示例数据
  - `activityId = 99887766`
  - `activityName = 上线流程`
- `large-activity-benchmark`
  - `displayName = 大数据活动压测项目`
  - `10000` 条示例数据
  - `4` 个普通字段：`owner_name`、`region`、`priority`、`status`
  - `2` 个 activity 头：`benchmarka`、`benchmarkb`
  - 每行字段为 `row_id + 10` 个业务字段

## 数据持久化说明

当前 mock 服务不落库，所有数据都保存在内存中。

这意味着：

- 服务运行期间，`/batchSave` 的修改会保留在当前进程内
- 重启 `node server.js` 后，数据会恢复为脚本中的初始值

## 集成测试如何使用它

`tests/OfficeAgent.IntegrationTests` 中的测试不会复用你手工启动的服务。

测试里的 `MockServerFixture` 会自动：

- 启动 `node tests/mock-server/server.js`
- 等待 `http://localhost:3100/login` 可访问
- 在测试结束后关闭该进程

因此：

- 手工联调插件时，需要自己运行 `npm start`
- 跑集成测试时，不需要提前手工启动 mock 服务

## 常见问题

### 1. 端口被占用

当前固定使用：

- `3100`
- `3200`

如果启动失败，先检查本机是否已有其他进程占用了这两个端口。

### 2. 业务接口返回 `401`

说明当前请求没有带上 SSO 登录产生的 cookie。

先完成一次：

- `http://localhost:3100/login`

再访问业务接口，或让插件通过 SSO 登录流程获取 cookie。

### 3. 上传后数据又恢复了

这是预期行为，因为 mock 数据只保存在内存里。重启服务后会回到脚本初始状态。
