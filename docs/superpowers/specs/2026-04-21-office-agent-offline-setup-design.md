# OfficeAgent 离线 Setup 设计说明

日期：2026-04-21

状态：设计已确认，待进入实施计划

## 1. 目标

把当前只产出 MSI 的安装流程升级为单文件离线 `setup.exe`，满足以下目标：

- 用户只需要下载并运行一个 `setup.exe`
- `setup.exe` 在本地离线环境下自动检测并按需安装：
  - Microsoft Visual Studio Tools for Office Runtime
  - Microsoft Edge WebView2 Runtime
- 先决条件安装完成后，再安装 OfficeAgent 本体
- 不依赖在线下载，不要求用户手动补装 runtime
- 继续保留现有 MSI 作为企业分发和兜底入口

## 2. 范围

### 2.1 本次范围

- 新增一个 WiX Burn bootstrapper，输出单文件 `OfficeAgent.Setup.exe`
- 把 `vstor_redist.exe`、WebView2 离线安装包、现有 MSI 一起打进 `setup.exe`
- 复用现有注册表检测规则，决定是否跳过对应 runtime 安装
- 保留现有 MSI 的启动条件检查，防止用户绕过 bootstrapper 后装出半残状态
- 更新安装与构建文档，说明离线包的组织方式和发布入口

### 2.2 本次明确不做

- 不做在线下载型安装器
- 不支持 ARM64 发行流程
- 不重写 OfficeAgent 业务逻辑
- 不改 VSTO/WebView2 的运行时依赖模型，只改安装编排
- 不把第三方离线安装包提交进 git 历史

## 3. 架构

安装链分两层：

- 外层：WiX Burn `setup.exe`
- 内层：现有 `OfficeAgent.Setup-x86.msi` 和 `OfficeAgent.Setup-x64.msi`

`setup.exe` 负责先决条件和安装顺序，MSI 负责安装 OfficeAgent 自身文件、注册表和 VSTO 入口。

运行时的 MSI 选择规则固定为：

- 32 位 Windows 安装 `OfficeAgent.Setup-x86.msi`
- 64 位 Windows 安装 `OfficeAgent.Setup-x64.msi`
- 不支持 ARM64 发行流程时直接 fail fast

安装顺序固定为：

1. 检测 VSTO Runtime
2. 缺失时静默安装 `vstor_redist.exe`
3. 检测 WebView2 Runtime
4. 缺失时静默安装对应架构的 WebView2 离线安装包
5. 安装现有 MSI

这样做的好处是：

- 机器级先决条件由外层编排，不污染应用 MSI
- 离线用户拿到一个文件即可完成安装
- 现有 MSI 仍可单独给 IT 或调试场景使用
- 失败点集中在 bootstrapper，行为更可控

## 4. 先决条件检测

### 4.1 VSTO Runtime

沿用现有 MSI 的检测依据：

- `HKLM\\SOFTWARE\\Microsoft\\VSTO Runtime Setup\\v4R\\Version`
- `HKLM\\SOFTWARE\\Microsoft\\VSTO Runtime Setup\\v4\\Install`
- 同时检查 32 位和 64 位注册表视图

判断规则：

- 只要检测到可用版本，就跳过安装
- 否则执行 bundle 内嵌的 `vstor_redist.exe`

### 4.2 WebView2 Runtime

沿用现有 MSI 的检测依据：

- `HKLM\\SOFTWARE\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}`
- `HKCU\\Software\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}`

判断规则：

- 任一受支持位置存在版本号，就跳过安装
- 否则执行 bundle 内嵌的 WebView2 离线安装包
- 64 位 Windows 使用 x64 安装包，32 位 Windows 使用 x86 安装包

### 4.3 失败策略

如果某个先决条件的静默安装失败：

- 立即停止整个安装
- 不继续执行 OfficeAgent MSI
- 输出清晰错误信息和日志入口

## 5. 文件布局

### 5.1 保留现有 MSI

现有 `installer/OfficeAgent.Setup/Product.wxs` 继续负责应用 MSI。

它仍然保留启动条件检查，作为兜底保护：

- 没有 VSTO Runtime 时阻止直接运行 MSI
- 没有 WebView2 Runtime 时阻止直接运行 MSI

### 5.2 新增 bundle 工程

新增一个 sibling bundle 目录，例如：

- `installer/OfficeAgent.SetupBundle/Bundle.wxs`
- `installer/OfficeAgent.SetupBundle/prereqs/`

`prereqs` 目录仅作为本地构建输入，不作为源码依赖提交到 git。

### 5.3 构建产物

最终对外发布物：

- `artifacts/installer/OfficeAgent.Setup.exe`

保留内部中间产物：

- `artifacts/installer/OfficeAgent.Setup-x86.msi`
- `artifacts/installer/OfficeAgent.Setup-x64.msi`

`setup.exe` 内部同时携带两个 MSI，但运行时只会按宿主 Windows 架构选择其一。

## 6. 构建流程

现有 `installer/OfficeAgent.Setup/build.ps1` 继续负责：

1. 构建前端
2. 生成版本号
3. 构建 VSTO add-in
4. 组装 MSI payload
5. 生成 x86/x64 MSI
6. 追加构建 bundle

bundle 构建前必须检查：

- `vstor_redist.exe` 是否存在
- WebView2 x86 离线包是否存在
- WebView2 x64 离线包是否存在
- 现有 MSI 输出是否存在

任何一个缺失都应直接 fail fast。

## 7. 验证策略

### 7.1 代码层验证

补充契约测试，覆盖：

- build 脚本是否明确产出 `setup.exe`
- bundle 源是否引用 VSTO 和 WebView2 离线安装包
- MSI 的 LaunchCondition 是否仍然存在
- 发布文档是否指向 `setup.exe` 作为用户入口

### 7.2 构建层验证

验证 `setup.exe` 构建结果：

- bundle 里确实包含先决条件和 MSI
- 产物路径正确
- 失败时有可定位日志

### 7.3 真实安装验证

由于当前开发机不一定是干净机器，最终安装验证分两类：

- 本机验证已安装依赖和重复运行场景
- Windows Sandbox / 临时 VM 验证全新安装场景

至少覆盖以下状态：

1. 两个依赖都没有
2. 只有 VSTO Runtime
3. 只有 WebView2 Runtime
4. 两个依赖都已安装
5. 已安装旧版 OfficeAgent，再运行新包

## 8. 风险与回退

- 如果 bundle 组装失败，仍可回退到现有 MSI 流程
- 如果用户绕过 `setup.exe`，MSI 的启动条件会继续拦截缺依赖安装
- 如果第三方离线安装包缺失，构建必须失败，不能产出不完整的发布物
