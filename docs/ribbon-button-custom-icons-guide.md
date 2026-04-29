# Ribbon Button Custom Icons Guide

日期：2026-04-28

本文说明后续如何把 Excel Ribbon 按钮的 Office 内置图标替换为项目自己的图片文件。

## 1. 当前状态

当前 Ribbon 按钮使用 Office 内置 `imageMso` 图标，代码入口在：

- `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`

典型配置如下：

```csharp
this.partialDownloadButton.OfficeImageId = "Refresh";
this.partialDownloadButton.ShowImage = true;
this.partialDownloadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
```

如果要使用自己的图片，应改为给按钮设置 `Image`，并清空 `OfficeImageId`。

重要：图标替换只应改变按钮图片来源，不应改变 Ribbon 标签来源。当前 Ribbon 组名、按钮名和项目下拉框状态文本由 `HostLocalizedStrings` 按 Excel UI 语言解析为中文或英文；不要在图标替换代码里重新写死中文或英文 `Label`。

## 2. 推荐图片规格

- 使用透明背景 PNG。
- 大按钮优先准备 `32x32` 图片。
- 常规小按钮优先准备 `16x16` 图片；当前 `Initialize sheet` 使用常规小按钮布局。
- 同一组按钮应保持统一线宽、色彩和留白。
- 不要使用本机绝对路径作为运行时图片路径，图片应随 add-in 编译发布。

## 3. 推荐资源存放方式

推荐把图片加入 `src/OfficeAgent.ExcelAddIn/Properties/Resources.resx`，由 `Properties.Resources` 强类型访问。

项目当前已经有：

- `src/OfficeAgent.ExcelAddIn/Properties/Resources.resx`
- `src/OfficeAgent.ExcelAddIn/Properties/Resources.Designer.cs`
- 示例资源属性：`Properties.Resources.Logo`

新增按钮图片时，建议命名为：

- `RibbonOpen`
- `RibbonInitializeSheet`
- `RibbonApplyTemplate`
- `RibbonSaveTemplate`
- `RibbonSaveAsTemplate`
- `RibbonPartialDownload`
- `RibbonPartialUpload`
- `RibbonLogin`
- `RibbonDocumentation`
- `RibbonAbout`

命名规则：`Ribbon` + 按钮语义，避免使用 `Icon1`、`Image2` 这类不可维护名称。

## 4. 用 Visual Studio 添加图片资源

1. 打开 `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`。
2. 打开 `Properties/Resources.resx`。
3. 选择 `Add Resource` -> `Add Existing File`。
4. 选择 PNG 图片文件。
5. 将资源名称改成语义化名称，例如 `RibbonPartialDownload`。
6. 保存 `.resx`，让 Visual Studio 重新生成 `Resources.Designer.cs`。

完成后应能在代码里访问：

```csharp
Properties.Resources.RibbonPartialDownload
```

## 5. 代码替换方式

推荐把自定义图片赋值集中放在 `AgentRibbon_Load`，避免图标配置散落在多个按钮初始化块里。保留现有本地化调用顺序：先应用本地化标签，再覆盖图片；项目下拉框使用 `ProjectDropDownPlaceholderText`，不要写成固定中文字符串。

示例：

```csharp
private void AgentRibbon_Load(object sender, RibbonUIEventArgs e)
{
    ApplyLocalizedLabels();
    ApplyCustomRibbonImages();

    SetProjectDropDownText(ProjectDropDownPlaceholderText);
    RefreshTemplateButtonsFromController();
    BindToControllersAndRefresh();
}

private void ApplyCustomRibbonImages()
{
    SetRibbonImage(toggleTaskPaneButton, Properties.Resources.RibbonOpen);
    SetRibbonImage(initializeSheetButton, Properties.Resources.RibbonInitializeSheet);
    SetRibbonImage(applyTemplateButton, Properties.Resources.RibbonApplyTemplate);
    SetRibbonImage(saveTemplateButton, Properties.Resources.RibbonSaveTemplate);
    SetRibbonImage(saveAsTemplateButton, Properties.Resources.RibbonSaveAsTemplate);
    SetRibbonImage(partialDownloadButton, Properties.Resources.RibbonPartialDownload);
    SetRibbonImage(partialUploadButton, Properties.Resources.RibbonPartialUpload);
    SetRibbonImage(loginButton, Properties.Resources.RibbonLogin);
    SetRibbonImage(documentationButton, Properties.Resources.RibbonDocumentation);
    SetRibbonImage(aboutButton, Properties.Resources.RibbonAbout);
}

private static void SetRibbonImage(RibbonButton button, System.Drawing.Image image)
{
    button.OfficeImageId = string.Empty;
    button.Image = image;
    button.ShowImage = true;
}
```

注意：

- `OfficeImageId` 和 `Image` 不要同时作为同一个按钮的图标来源。
- 设置自定义图片时，应清空 `OfficeImageId`。
- 保留 `ShowImage = true`。
- 不要在图片替换 helper 中统一覆盖 `ControlSize`。Designer 负责维护每个按钮的布局：`Initialize sheet` 是 `RibbonControlSizeRegular`；其余命令按钮仍是 `RibbonControlSizeLarge`。
- `ApplyCustomRibbonImages` 不应设置 `button.Label`、Ribbon group `Label` 或项目下拉框文本；这些文本仍由 `ApplyLocalizedLabels()` 和 `HostLocalizedStrings` 管理。

## 6. Designer 文件处理原则

`AgentRibbon.Designer.cs` 当前包含按钮声明、分组、标签、内置图标和布局配置。Designer 中的标签应保持英文兜底值，运行时再由 `ApplyLocalizedLabels()` 根据 Excel UI 语言覆盖为中文或英文。

后续替换为自定义图片时有两种选择：

- 保留 Designer 中的按钮、分组、英文兜底标签、布局，只在 `AgentRibbon.cs` 的 `AgentRibbon_Load` 里覆盖图片。
- 或者直接在 Designer / Visual Studio Ribbon Designer 中设置 `Image` 属性。

推荐第一种：图片赋值集中在 `AgentRibbon.cs`，更容易审查和测试；Designer 继续负责控件结构。

如果保留 Designer 里的 `OfficeImageId`，运行时 `SetRibbonImage` 必须清空它。否则后续排查时很难判断实际显示来源。

不要通过 Designer 或图标初始化代码把 `Project`、`Data sync`、`Documentation` 等标签改回中文硬编码；否则英文 Excel 会再次出现中文按钮名。

## 7. 测试更新

如果正式切换为自定义图片，需要同步更新：

- `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`

测试重点从 “包含指定 `OfficeImageId`” 改为：

- 按钮仍然 `ShowImage = true`
- 按钮保留各自配置的 `ControlSize`：`Initialize sheet` 为 `RibbonControlSizeRegular`，其余命令按钮为 `RibbonControlSizeLarge`
- `AgentRibbon.cs` 中存在集中赋值方法
- 每个按钮都绑定了对应的 `Properties.Resources.Ribbon...` 图片
- 不再依赖 `OfficeImageId` 作为最终图标来源
- `RibbonDesignerUsesEnglishSafeDefaultsForLocalizedControls`、`RibbonAppliesLocalizedLabelsFromHostStringsAtRuntime` 和 `HostLocalizedStringsTests` 仍然通过，确保图标替换没有破坏中英文 Ribbon 标签。

## 8. 构建和开发刷新

替换图片后，运行：

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~AgentRibbonConfigurationTests"
pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel
```

然后重新打开 Excel 验证 Ribbon。

如果只替换了资源图片但 Excel 已经打开，也必须重启 Excel。Ribbon 图标不会在当前 Excel 进程里自动刷新。

## 9. 手动验收清单

打开 Excel 后确认：

- `xISDP` tab 能正常显示。
- 每个目标按钮显示自定义图片。
- `Initialize sheet` 仍是常规小按钮布局；其余命令按钮仍是上方图标、下方文字。
- `Apply template`、`Save template`、`Save as template` 仍是大按钮。
- 图片没有被拉伸、裁切、模糊或出现白底。
- 中文 Excel 下显示 `项目`、`数据同步`、`文档`、`关于`；英文 Excel 下显示 `Project`、`Data sync`、`Documentation`、`About`。
- `Open`、`下载` / `Download`、`上传` / `Upload`、`文档` / `Documentation`、`关于` / `About` 等按钮点击行为不变。
- `关于` / `About` 仍显示版本号、程序集版本、构建配置和构建时间。
- `文档` / `Documentation` 仍用默认浏览器打开配置的文档 URL。

## 10. 常见问题

### Excel 里还是旧图标

优先检查：

- 是否运行了 `eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel`。
- 是否完全关闭了所有 `EXCEL.EXE`。
- 注册表是否仍指向当前仓库的 Debug manifest：

```powershell
Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Excel\Addins\OfficeAgent.ExcelAddIn'
```

`Manifest` 应指向：

```text
file:///D:/Workspace/demos/office-agent/src/OfficeAgent.ExcelAddIn/bin/Debug/OfficeAgent.ExcelAddIn.vsto|vstolocal
```

### 代码里找不到 `Properties.Resources.Ribbon...`

通常是资源没有进入 `.resx`，或 `Resources.Designer.cs` 没有重新生成。

处理方式：

- 确认图片已加入 `Properties/Resources.resx`。
- 在 Visual Studio 中保存 `.resx`。
- 重新编译 `OfficeAgent.ExcelAddIn` 项目。

### 图片不显示但按钮文字显示

检查：

- `button.Image` 是否为 `null`。
- `button.ShowImage` 是否为 `true`。
- `button.OfficeImageId` 是否仍占用图标来源。
- 资源图片格式是否为正常 PNG / BMP。

### 图标位置或按钮大小不符合预期

检查：

```csharp
button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
```

`Initialize sheet` 应保持 `RibbonControlSizeRegular`。其余命令按钮应保持 `RibbonControlSizeLarge`，大按钮布局才会按 Office Ribbon 的常规方式显示为上方图标、下方文字。不要为了替换图片在 helper 中统一覆盖所有按钮的 `ControlSize`。

### 英文 Excel 里又显示中文按钮名

这通常不是图标资源问题，而是图标替换时误改了 Ribbon 标签来源。优先检查：

- `AgentRibbon_Load` 是否仍调用 `ApplyLocalizedLabels()`。
- 项目下拉框是否仍调用 `SetProjectDropDownText(ProjectDropDownPlaceholderText)`，而不是 `SetProjectDropDownText("先选择项目")`。
- `ApplyCustomRibbonImages` 是否只设置 `OfficeImageId`、`Image` 和 `ShowImage`，没有设置任何 `Label` 或统一覆盖 `ControlSize`。
- `AgentRibbon.Designer.cs` 中本地化控件是否仍保持英文兜底标签。
- `HostLocalizedStrings` 是否仍包含对应中文和英文标签。

## 11. 维护约定

当 Ribbon 图标从 Office 内置图标切换为自定义图片时，同一改动应同步更新：

- `docs/modules/ribbon-sync-current-behavior.md`
- `docs/vsto-manual-test-checklist.md`
- `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`

如果只是替换某个图片文件但按钮语义和布局不变，仍建议记录替换原因和图片来源，避免后续不知道资源用途。
