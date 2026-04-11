# Mock Server

`tests/mock-server` 提供本地联调用的 Node.js mock 服务。

当前它同时启动两类服务：

- SSO 登录服务：`http://localhost:3100`
- 业务 API 服务：`http://localhost:3200`

服务入口脚本是 [server.js](/D:/Workspace/demos/office-agent/.worktrees/ribbon-sync/tests/mock-server/server.js)。

## 手工启动

在仓库根目录执行：

```powershell
cd tests/mock-server
npm install
npm start
```

`npm start` 会执行：

```powershell
node server.js
```

启动成功后，控制台会输出当前可访问地址和推荐的插件配置项。

## 插件联调配置

本地联调 Excel 插件时，使用以下配置：

- `Base URL = 你的大模型服务地址`
- `Business Base URL = http://localhost:3200`
- `SSO URL = http://localhost:3100/login`
- `登录成功路径 = /rest/login`
- `API Key = 留空`

说明：

- `Base URL` 只用于大模型或 Agent 能力，不绑定业务 mock 服务
- `Business Base URL` 才是 Ribbon Sync、`/find`、`/head`、`/batchSave` 和 `upload_data` 等业务接口的基地址
- 当前 mock 服务通过 SSO cookie 鉴权，不走 API Key
- 业务接口在未登录状态下会返回 `401`

## 当前接口

### SSO 服务

- `GET /login`
  - 返回登录页
- `POST /rest/login`
  - 提交用户名密码
  - 返回登录成功响应并写入 cookie

### 业务 API

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

### Ribbon Sync 相关接口

- `POST /head`
  - 返回表头定义
  - 当前返回体包含 `headList`
- `POST /find`
  - 同时用于全量下载和部分下载
  - 请求体支持：
    - `projectId`
    - `ids`
    - `fieldKeys`
  - `ids` 为空时返回全量行
  - `fieldKeys` 为空时返回整行字段
- `POST /batchSave`
  - 用于全量上传、部分上传、增量上传
  - 请求体是一个 list
  - 每个 item 对应一个单元格改动，字段包括：
    - `id`
    - `fieldKey`
    - `value`
    - `projectId`

## 当前内置数据

当前 Ribbon Sync mock 数据保存在 [server.js](/D:/Workspace/demos/office-agent/.worktrees/ribbon-sync/tests/mock-server/server.js) 内存变量中，主要包括：

- `connectorRows`
  - `/find` 和 `/batchSave` 使用的数据行
- `connectorHeadList`
  - `/head` 返回的表头定义

当前内置了一组活动属性字段示例：

- `start_12345678`
- `end_12345678`

其中活动信息为：

- `activityId = 12345678`
- `activityName = 测试活动111`

## 数据持久化说明

当前 mock 服务不落库，所有数据都在内存中维护。

这意味着：

- 服务运行期间，`/batchSave` 的修改会保留在当前进程内
- 重启 `node server.js` 后，数据会恢复为脚本中的初始值

## 集成测试如何使用它

`tests/OfficeAgent.IntegrationTests` 中的集成测试不会复用你手工启动的服务。

测试里的 `MockServerFixture` 会自动：

- 启动 `node tests/mock-server/server.js`
- 等待 `http://localhost:3100/login` 可访问
- 测试结束后关闭该进程

相关实现位于：

- [BusinessApiIntegrationTests.cs](/D:/Workspace/demos/office-agent/.worktrees/ribbon-sync/tests/OfficeAgent.IntegrationTests/BusinessApiIntegrationTests.cs)

因此：

- 手工联调插件时，请自己运行 `npm start`
- 跑集成测试时，不需要提前手工启动 mock 服务

## 常见问题

### 1. 端口被占用

当前固定使用：

- `3100`
- `3200`

如果启动失败，先检查本机是否已有其他进程占用了这两个端口。

### 2. 业务接口返回 401

说明当前请求没有带上 SSO 登录产生的 cookie。

先完成一次：

- `http://localhost:3100/login`

再访问业务接口，或让插件通过 SSO 登录流程获取 cookie。

### 3. 上传后数据又恢复了

这是当前 mock 服务的预期行为，因为数据只保存在内存里。重启服务后会回到脚本初始状态。
