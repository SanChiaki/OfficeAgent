# BatchUploadProgressDialog 使用文档

本文档用于指导内网 AI 或开发人员直接使用“批量上传 5 步进度弹窗壳子”。

弹窗壳子只负责展示，不负责业务判断。业务代码需要把每一步的标题、说明、状态、详情文本整理成 `BatchUploadProgressStep` 列表，然后传给 `BatchUploadProgressDialog`。

## 1. 文件位置

弹窗文件：

```text
src/OfficeAgent.ExcelAddIn/Dialogs/BatchUploadProgressDialog.cs
```

命名空间：

```csharp
OfficeAgent.ExcelAddIn.Dialogs
```

当前项目文件已经包含该弹窗：

```text
src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj
```

## 2. 适用场景

该弹窗适合展示类似“批量上传”的固定流程：

1. 数据准备
2. 字段验证
3. 变更预览
4. 数据上传
5. 上传结果

也可以复用为其他 3 到 8 步左右的流程弹窗。步骤多、详情长时，弹窗中间区域会滚动，底部关闭按钮固定不动。

## 3. 数据结构

业务侧只需要构造 `BatchUploadProgressStep`。

```csharp
new BatchUploadProgressDialog.BatchUploadProgressStep(
    title: "步骤标题",
    description: "步骤说明",
    state: BatchUploadProgressDialog.BatchUploadStepState.Active,
    details: "可选详情文本");
```

字段含义：

| 字段 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `title` | `string` | 是 | 步骤标题，例如 `数据准备` |
| `description` | `string` | 是 | 步骤简短说明，例如 `正在读取 Excel 选区` |
| `state` | `BatchUploadStepState` | 是 | 步骤状态 |
| `details` | `string` | 否 | 详情文本，适合放预览内容、错误原因、上传结果 |

状态枚举：

| 状态 | 用途 |
| --- | --- |
| `Pending` | 尚未开始 |
| `Active` | 当前正在执行 |
| `Completed` | 已完成 |
| `Warning` | 有警告，但流程可继续 |
| `Error` | 失败或不可继续 |

## 4. 最小调用示例

在 Excel Add-in 同一个程序集内，可以直接这样调用：

```csharp
using OfficeAgent.ExcelAddIn.Dialogs;

var steps = new[]
{
    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "数据准备",
        "已读取 Excel 可见选区",
        BatchUploadProgressDialog.BatchUploadStepState.Completed),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "字段验证",
        "字段校验通过",
        BatchUploadProgressDialog.BatchUploadStepState.Completed),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "变更预览",
        "请确认本次将上传的内容",
        BatchUploadProgressDialog.BatchUploadStepState.Active,
        "将上传 48 个单元格\r\n跳过 4 个空白单元格"),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "数据上传",
        "等待上传",
        BatchUploadProgressDialog.BatchUploadStepState.Pending),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "上传结果",
        "等待服务器返回结果",
        BatchUploadProgressDialog.BatchUploadStepState.Pending),
};

using (var dialog = new BatchUploadProgressDialog(steps))
{
    dialog.ShowDialog();
}
```

## 5. 推荐的 Excel 父窗口调用方式

在 VSTO 插件里弹窗时，建议绑定 Excel 主窗口作为父窗口，避免弹窗跑到 Excel 后面。

```csharp
using OfficeAgent.ExcelAddIn.Dialogs;

using (var dialog = new BatchUploadProgressDialog(steps))
{
    var owner = ExcelDialogOwner.FromCurrentApplication();
    if (owner == null)
    {
        dialog.ShowDialog();
    }
    else
    {
        dialog.ShowDialog(owner);
    }
}
```

内网 AI 生成接入代码时，优先使用这一段。

## 6. 批量上传业务映射示例

假设业务侧已经拿到了这些结果：

```csharp
var selectedCellCount = 52;
var uploadCellCount = 48;
var skippedCellCount = 4;
var previewLines = new[]
{
    "0331test / taskFlowNode_13882098334 -> 测试",
    "0331test / taskFlowNode_13892195334 -> 1111",
    "0331test1 / SITEOWNER -> 15012344321",
};
var uploadSucceeded = true;
```

可以映射成：

```csharp
var previewText = string.Join("\r\n", previewLines);

var steps = new[]
{
    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "数据准备",
        "已读取选区，共 " + selectedCellCount + " 个单元格",
        BatchUploadProgressDialog.BatchUploadStepState.Completed),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "字段验证",
        "验证所有者、每日工作及其他字段",
        BatchUploadProgressDialog.BatchUploadStepState.Completed),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "变更预览",
        "本次将上传 " + uploadCellCount + " 个单元格，跳过 " + skippedCellCount + " 个单元格",
        BatchUploadProgressDialog.BatchUploadStepState.Completed,
        previewText),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "数据上传",
        "已提交至服务器",
        BatchUploadProgressDialog.BatchUploadStepState.Completed),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "上传结果",
        uploadSucceeded ? "上传完成" : "上传失败",
        uploadSucceeded
            ? BatchUploadProgressDialog.BatchUploadStepState.Completed
            : BatchUploadProgressDialog.BatchUploadStepState.Error,
        uploadSucceeded ? "成功：" + uploadCellCount + " 项变更" : "请查看日志确认失败原因"),
};
```

## 7. 交互式更新示例

弹窗支持打开后继续更新步骤状态。业务代码可以先用 `Pending` 构造 5 个步骤，再随着上传流程推进调用更新方法。

```csharp
var steps = new[]
{
    new BatchUploadProgressDialog.BatchUploadProgressStep("数据准备", "等待开始", BatchUploadProgressDialog.BatchUploadStepState.Pending),
    new BatchUploadProgressDialog.BatchUploadProgressStep("字段验证", "等待开始", BatchUploadProgressDialog.BatchUploadStepState.Pending),
    new BatchUploadProgressDialog.BatchUploadProgressStep("变更预览", "等待开始", BatchUploadProgressDialog.BatchUploadStepState.Pending),
    new BatchUploadProgressDialog.BatchUploadProgressStep("数据上传", "等待开始", BatchUploadProgressDialog.BatchUploadStepState.Pending),
    new BatchUploadProgressDialog.BatchUploadProgressStep("上传结果", "等待服务器返回结果", BatchUploadProgressDialog.BatchUploadStepState.Pending),
};

using (var dialog = new BatchUploadProgressDialog(steps))
{
    dialog.UploadRequested += (sender, args) =>
    {
        dialog.SetStepActive(4, "数据上传", "正在上传至服务器");
        // 执行上传逻辑...
    };

    dialog.UploadCanceled += (sender, args) =>
    {
        // 做业务清理或取消请求；弹窗会在事件触发后自动关闭。
    };

    dialog.Confirmed += (sender, args) =>
    {
        // 可在这里刷新外部界面；弹窗会在事件触发后自动关闭。
    };

    dialog.Show();

    dialog.SetStepActive(1, "数据准备", "正在读取 Excel 可见选区");
    // 执行业务逻辑...
    dialog.SetStepCompleted(1, "数据准备", "已读取 Excel 可见选区");

    dialog.SetStepActive(2, "字段验证", "正在验证字段");
    dialog.AppendStepDetails(2, "SITEOWNER 校验通过");
    var hasUploadableContent = true; // 来自第 2 步校验结果，例如 validUploadCount > 0。
    dialog.SetPreviewUploadAvailability(hasUploadableContent);
    dialog.SetStepCompleted(2, "字段验证", "验证通过");

    dialog.SetStepActive(3, "变更预览", hasUploadableContent ? "请确认本次上传内容" : "没有可上传内容", "预览详情...");
}
```

交互式 API：

| 方法 | 用途 |
| --- | --- |
| `SetStepPending(stepNumber, title, description, details)` | 把步骤更新为未开始 |
| `SetStepActive(stepNumber, title, description, details)` | 把步骤更新为正在进行，右侧显示动态圆环 |
| `SetStepCompleted(stepNumber, title, description, details)` | 把步骤更新为完成，左侧显示对勾，右侧圆环消失 |
| `SetStepWarning(stepNumber, title, description, details)` | 把步骤更新为警告 |
| `SetStepError(stepNumber, title, description, details)` | 把步骤更新为失败 |
| `SetPreviewUploadAvailability(hasUploadableContent)` | 设置第 3 步是否存在合法可上传内容，通常来自第 2 步校验结果 |
| `AppendStepDetails(stepNumber, details)` | 给某一步追加详情日志 |
| `UploadRequested` | 用户在第 3 步点击底部【上传】按钮时触发 |
| `UploadCanceled` | 用户在第 1、2、3、4 步点击底部【取消】按钮时触发，触发后弹窗默认关闭 |
| `Confirmed` | 用户在第 3 步无合法可上传内容或第 5 步点击底部【确认】按钮时触发，触发后弹窗默认关闭 |

只有 `Active` 步骤会显示右侧动态圆环，其他状态不显示右侧圆环。

底部按钮根据当前 `Active` 步骤自动变化：

| 当前步骤 | 底部按钮 |
| --- | --- |
| 第 1 步：数据准备 | 【取消】 |
| 第 2 步：字段验证 | 【取消】 |
| 第 3 步：变更预览 | 有合法可上传内容时【上传】【取消】；无合法可上传内容时【确认】 |
| 第 4 步：数据上传 | 【取消】 |
| 第 5 步：上传结果 | 非 `Pending` 时显示【确认】 |

详情框只允许出现在第 3 步和第 5 步。第 3 步的变更预览详情框显示后不会因为进入第 4 步而收起；第 5 步的结果详情框显示后也会保留。第 1、2、4 步不会显示详情框，即使调用方传入了 `details`。上传过程日志如果需要展示，应在第 5 步结果里展示。

## 8. 失败场景示例

如果字段验证失败，不应该继续标记后续步骤为完成，可以这样展示：

```csharp
var steps = new[]
{
    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "数据准备",
        "已读取 Excel 可见选区",
        BatchUploadProgressDialog.BatchUploadStepState.Completed),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "字段验证",
        "发现无法上传的字段",
        BatchUploadProgressDialog.BatchUploadStepState.Error,
        "第 12 行 SITEOWNER 不能为空\r\n第 15 行 ProjectId 无效"),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "变更预览",
        "等待字段验证通过",
        BatchUploadProgressDialog.BatchUploadStepState.Pending),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "数据上传",
        "未开始",
        BatchUploadProgressDialog.BatchUploadStepState.Pending),

    new BatchUploadProgressDialog.BatchUploadProgressStep(
        "上传结果",
        "未上传",
        BatchUploadProgressDialog.BatchUploadStepState.Pending),
};
```

## 9. 内网 AI 接入规则

内网 AI 生成代码时，请遵守下面规则：

| 规则 | 要求 |
| --- | --- |
| 使用交互 API | 不要修改弹窗内部布局，业务流程推进时调用 `SetStepActive`、`SetStepCompleted` 等方法 |
| 绑定 Excel 父窗口 | 优先使用 `ExcelDialogOwner.FromCurrentApplication()` |
| 详情文本换行 | 使用 `\r\n` 拼接多行详情 |
| 长详情放 `details` | 不要把长文本塞进 `description` |
| 成功流程 | 已完成步骤用 `Completed`，当前步骤用 `Active` |
| 失败流程 | 失败步骤用 `Error`，后续未执行步骤用 `Pending` |
| 不使用示例入口 | `CreateSample()` 只用于截图和本地预览，不要在业务代码里调用 |

## 10. 缩放适配说明

该弹窗按下面方式处理 50% 到 300% 缩放：

| 区域 | 处理方式 |
| --- | --- |
| 系统标题栏 | 使用 WinForms 系统标题栏关闭按钮，不再自绘右上角 `X` |
| 弹窗位置 | 默认显示在屏幕正中间 |
| 顶部标题 | 根据字体高度动态计算 header 高度，避免 200% 以上压缩 |
| 中间步骤区 | 使用可滚动内容区，步骤内容按真实文本高度自适应 |
| 详情区域 | 详情框使用白色背景，并有更大的最小和最大高度，内容过长时可上下和左右滚动 |
| 右侧圆环 | 只有正在进行的 `Active` 步骤右侧显示动态圆环 |
| 底部按钮 | 根据当前步骤显示；第 1、2、4 步显示【取消】，第 3 步显示【上传】【取消】，第 5 步非 `Pending` 时显示【确认】 |

## 11. 本地化要求

仓库约定：新增弹窗的用户可见文字必须通过 `HostLocalizedStrings.cs` 本地化。

当前弹窗壳子的业务步骤文字由调用方传入，因此调用方应传入已经本地化后的文本。

固定文案也需要本地化，主要包括：

| 文案 | 当前值 | 建议来源 |
| --- | --- | --- |
| 弹窗标题 | `批量上传` | `HostLocalizedStrings` |
| 关闭按钮 | `关闭` | `HostLocalizedStrings` |

如果正式接入主流程，内网 AI 应先检查 `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`，新增或复用对应的本地化属性，再替换弹窗里的固定硬编码文字。

## 12. 本地预览截图

已有本地预览工具：

```text
tools/RenderBatchUploadProgressDialog.cs
```

可生成 50%、75%、100%、125%、150%、175%、200%、225%、250%、300% 的预览图。

输出目录：

```text
dialog-preview/
```

生成截图时可以使用仓库里已有的编译方式：

```powershell
$env:TEMP = (Resolve-Path .).Path
$env:TMP = (Resolve-Path .).Path
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /nologo /target:winexe /out:RenderBatchUploadProgressDialog.exe /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll src\OfficeAgent.ExcelAddIn\Dialogs\BatchUploadProgressDialog.cs tools\RenderBatchUploadProgressDialog.cs
.\RenderBatchUploadProgressDialog.exe dialog-preview
Remove-Item .\RenderBatchUploadProgressDialog.exe
```

## 13. 给内网 AI 的最短指令模板

如果要让内网 AI 接入该弹窗，可以直接给它下面这段要求：

```text
请使用 src/OfficeAgent.ExcelAddIn/Dialogs/BatchUploadProgressDialog.cs 作为批量上传进度弹窗壳子。
不要重写 WinForms 布局。
请先构造 BatchUploadProgressDialog.BatchUploadProgressStep[] 初始步骤，再用 SetStepActive、SetStepCompleted、SetStepError、AppendStepDetails 等方法交互式推进状态。
步骤状态只能使用 Pending、Active、Completed、Warning、Error。
长文本放 details，追加日志使用 AppendStepDetails。
底部按钮根据当前步骤变化：第1/2步【取消】，第3步【上传】【取消】，第4步无按钮，第5步【确认】。业务代码通过 UploadRequested、UploadCanceled、Confirmed 事件处理点击。
弹窗展示时使用 ExcelDialogOwner.FromCurrentApplication() 绑定 Excel 父窗口。
不要调用 CreateSample()，它只用于预览。
所有新增用户可见固定文案必须走 HostLocalizedStrings.cs 本地化。
```
