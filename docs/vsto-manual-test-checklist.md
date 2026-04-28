# OfficeAgent VSTO Manual Test Checklist

## Installer

- Confirm `installer/OfficeAgent.SetupBundle/prereqs/` contains `vstor_redist.exe` before starting the build.
- Run `installer/OfficeAgent.Setup/build.ps1` and confirm `artifacts/installer/xISDP.Setup.exe`, `artifacts/installer/xISDP.Setup-x86.msi`, and `artifacts/installer/xISDP.Setup-x64.msi` are created.
- Run `xISDP.Setup.exe` on a machine missing the VSTO runtime and confirm it installs VSTO Runtime and then installs xISDP.
- Run `xISDP.Setup.exe` on a machine missing the WebView2 runtime and confirm it still installs xISDP without attempting to install WebView2.
- Run `xISDP.Setup.exe` on a machine with the VSTO runtime already installed and confirm it skips the VSTO installer.
- Run `xISDP.Setup.exe` twice on the same machine and confirm the second run does not reinstall VSTO and falls through to normal xISDP maintenance behavior.
- Choose the MSI that matches the target Excel bitness only for direct enterprise distribution or debugging. Do not install the x86 package for x64 Excel or the x64 package for x86 Excel.
- Install the direct MSI under a standard user profile.
- Confirm files are deployed under `%LocalAppData%\\OfficeAgent\\ExcelAddIn`.
- Confirm Excel add-in registry entries exist under `HKCU\\Software\\Microsoft\\Office\\Excel\\Addins\\OfficeAgent.ExcelAddIn`.
- On a machine missing the VSTO runtime, confirm the direct MSI blocks with a clear prerequisite message.
- On a machine missing the WebView2 runtime, confirm the direct MSI blocks with a clear prerequisite message.
- Note the current MVP manifests are signed with the development publisher `OfficeAgent Dev Certificate`; for distribution outside the build machine, replace it with a trusted code-signing certificate or import the publisher certificate through your enterprise deployment flow.

## Excel Startup

- Start Excel 2019 x86 on Windows with the x86 MSI and confirm `OfficeAgent` loads without manual sideload.
- Start Excel 2019 x64 on Windows with the x64 MSI and confirm `OfficeAgent` loads without manual sideload.
- Close and reopen Excel and confirm the add-in still loads automatically.

## Task Pane

- Run the task-pane checks twice: once under Chinese Excel (`zh-*` UI) and once under English Excel (any non-`zh-*` UI).
- In Chinese Excel, confirm the fixed task-pane UI and frontend system messages render in Chinese.
- In English Excel, confirm the fixed task-pane UI and frontend system messages render in English.
- In English Excel, trigger one host-generated task-pane error or fallback path and confirm the copy renders in English.
- In browser preview mode, confirm the fixed UI defaults to English before any explicit `uiLanguageOverride` is saved.
- In English Excel, send a Chinese free-form prompt and confirm the AI reply still follows the prompt language when possible instead of being forced to English.
- In Chinese Excel, send an English free-form prompt and confirm the AI reply can still stay in English when appropriate.
- Use the Ribbon button to open and close the task pane repeatedly.
- Confirm the task pane does not duplicate after repeated toggles.
- Confirm the WebView2 missing-runtime fallback message appears if WebView2 is not installed.

## Session And Settings

- Open Settings and save `API Key`, `Base URL`, `Business Base URL`, `Model`, `SSO URL`, and `登录成功路径` / `Login success path`.
- Confirm `Base URL` stays reserved for the LLM endpoint and `Business Base URL` points to the business API or mock server.
- Restart Excel and confirm settings reload correctly.
- Create or switch sessions and confirm existing thread history is preserved per session.

## Selection Context

- Select a contiguous range and confirm workbook, sheet, address, row count, column count, and headers update.
- Select a non-contiguous range and confirm the warning is shown.

## upload_data

- Trigger `upload_data` with natural language and confirm a preview card appears.
- Trigger `upload_data` with `/upload_data ...` and confirm it also routes to the skill.
- Cancel the preview and confirm the thread logs the cancellation without changing Excel.
- Confirm the preview and verify the external API is called with the selected rows.
- Simulate a 4xx/5xx API failure and confirm the error message is shown in the task pane.
- Configure a `Business Base URL` with a path prefix such as `/v1/` and confirm the request preserves the prefix.

## Excel Command Confirmation

- Run a read command and confirm it executes immediately.
- Run a write command and confirm it requires preview + confirmation.
- Leave a confirmation card open and verify the composer stays disabled until confirm or cancel.

## Ribbon Sync

- Run the Ribbon Sync checks twice: once under Chinese Excel (`zh-*` UI) and once under English Excel (any non-`zh-*` UI).
- In Chinese Excel, confirm the Ribbon group/button labels, project dropdown statuses, login popup, layout dialog, and native confirmation/result dialogs render in Chinese.
- In English Excel, confirm the Ribbon group/button labels, project dropdown statuses, login popup, layout dialog, and native confirmation/result dialogs render in English.
- For the steps below, use the localized host labels for the current Excel UI. Key pairs: `初始化当前表` / `Initialize sheet`, `应用模板` / `Apply template`, `保存模板` / `Save template`, `另存模板` / `Save as template`, `部分下载` / `Partial download`, `部分上传` / `Partial upload`, `先选择项目` / `Select project`, `请先登录` / `Sign in first`, `无可用项目` / `No projects available`.
- Bind a blank worksheet through the Ribbon project dropdown and confirm the layout dialog appears with defaults `HeaderStartRow = 1`, `HeaderRowCount = 2`, `DataStartRow = 3`.
- Confirm the layout dialog and enter custom values, then verify `ISDP_Setting` writes one `SheetBindings` row with the user-entered layout values.
- Confirming project selection should still not auto-initialize the current sheet; `SheetFieldMappings` remains unchanged until `初始化当前表` / `Initialize sheet` is clicked.
- Open `ISDP_Setting` and confirm it uses one worksheet with three readable sections: `TemplateBindings` on top, `SheetBindings` in the middle, `SheetFieldMappings` below, each with a title row, a header row, and data rows.
- Confirm `SheetFieldMappings` displays headers in this order: `HeaderType`, `ISDP L1`, `Excel L1`, `ISDP L2`, `Excel L2`, `HeaderId`, `ApiFieldKey`, `IsIdColumn`, `ActivityId`, `PropertyId`.
- Confirm there are two blank separator rows between each adjacent metadata section, and that metadata is no longer stored as flattened `tableName + values` rows.
- Switch to a worksheet with existing binding metadata and confirm the Ribbon dropdown automatically rehydrates that project as `ProjectId-DisplayName` instead of showing the localized empty placeholder (`先选择项目` / `Select project`).
- Save a workbook with `ISDP_Setting` as the active sheet, reopen Excel from the desktop shortcut, and confirm the Ribbon dropdown shows the localized empty placeholder (`先选择项目` / `Select project`) unless `SheetBindings` contains an explicit binding row for `ISDP_Setting`.
- Switch from a bound business sheet to `ISDP_Setting` and confirm the Ribbon dropdown clears back to the localized empty placeholder (`先选择项目` / `Select project`) when `ISDP_Setting` itself has no binding.
- Switch to a worksheet without binding metadata and confirm the Ribbon dropdown shows the localized empty placeholder (`先选择项目` / `Select project`).
- Open two workbooks in the same Excel process, bind `Sheet1` in each workbook to different projects, switch back and forth between the two files, and confirm the Ribbon dropdown plus download/upload behavior always follow the active workbook's own `ISDP_Setting` metadata.
- On a sheet that already has binding metadata, switch to another project and confirm the layout dialog defaults reuse the current sheet's saved layout values.
- Increase Windows or Office UI font scaling, reopen the project layout dialog, and confirm all labels, inputs, and buttons remain fully visible without overlap.
- Under both Chinese Excel and English Excel, trigger the unauthenticated path and confirm the native auth-required prompt plus login button text are localized for that host language.
- Reselect the already bound project (`same systemKey + projectId`) and confirm no layout dialog appears and `SheetBindings` is not rewritten.
- Cancel the layout dialog while switching projects and confirm both `ISDP_Setting` binding data and Ribbon dropdown project stay unchanged.
- After switching to another project and confirming the dialog, verify old `SheetFieldMappings` are cleared; before clicking `初始化当前表` / `Initialize sheet`, running download/upload should report that the current sheet is not initialized.
- Enter invalid values in the layout dialog (for example overlaps between header/data regions) and confirm validation error is shown while keeping the dialog open.
- Start Excel while unauthenticated against a protected project API and confirm the project dropdown shows the localized sign-in-required status (`请先登录` / `Sign in first`).
- Configure the project API to return an empty array and confirm the project dropdown shows the localized empty-project status (`无可用项目` / `No projects available`).
- Click `初始化当前表` / `Initialize sheet` on a sheet that already contains business cells and confirm only `ISDP_Setting` changes; the business area should remain untouched.
- Click `部分下载` / `Partial download` and `部分上传` / `Partial upload` and confirm each action uses a native Office/WinForms confirmation dialog instead of the task pane.
- Before `部分下载` / `Partial download`, put an old value in one selected managed non-ID cell, run the download, and confirm `xISDP_Log` is created with columns `key`, `表头`, `修改模式`, `修改值`, `原始值`, `修改时间`; the new row should show `修改模式 = 下载`, `原始值` as the overwritten Excel value, and `修改值` as the downloaded value.
- Edit one managed non-ID cell, run `部分上传` / `Partial upload`, and confirm `xISDP_Log` appends one `修改模式 = 上传` row using the user's pre-edit Excel value as `原始值` and the uploaded cell value as `修改值`.
- Force a `部分上传` / `Partial upload` failure from the mock server or API, then confirm no new `上传` row is added to `xISDP_Log`; retry successfully and confirm the original pre-edit value is still used.
- Add more than 2000 sync log rows through repeated upload/download validation or seeded workbook data, then trigger another logged sync and confirm `xISDP_Log` keeps only the latest 2000 data rows plus the header row.
- Confirm the Ribbon includes a dedicated `模板` / `Template` group with `应用模板` / `Apply template`, `保存模板` / `Save template`, and `另存模板` / `Save as template`.
- Confirm all Ribbon buttons display Office built-in icons that match their action semantics. `初始化当前表` / `Initialize sheet` should use the small regular button layout; the other command buttons should remain in the large icon-above-label layout.
- Confirm the `模板` / `Template` group shows `应用模板` / `Apply template`, `保存模板` / `Save template`, and `另存模板` / `Save as template` as large buttons.
- Confirm the `xISDP AI` group button is labeled `Open`.
- Confirm the Ribbon includes one `数据同步` / `Data sync` group containing `部分下载` / `Partial download` and `部分上传` / `Partial upload`, and that there is no `全量下载`, `全量上传`, or `增量上传` button.
- Confirm the Ribbon includes a `帮助` / `Help` group with `文档` / `Documentation` and `关于` / `About`; `文档` / `Documentation` opens `https://github.com/SanChiaki/OfficeAgent` in the default browser, and `关于` / `About` shows version and build information.
- In the same project, save two different local templates and confirm `应用模板` / `Apply template` can list both.
- Apply one template and confirm `TemplateBindings` updates to the selected template while `SheetBindings` / `SheetFieldMappings` are expanded into the current sheet.
- Manually edit `ISDP_Setting` field mapping text after applying a template, click `保存模板` / `Save template`, then reapply that template and confirm the edited mapping is preserved.
- With a sheet already bound to a template, use `另存模板` / `Save as template`, confirm the new template name appears in the local template list, and confirm the current sheet's `TemplateBindings.TemplateId` switches to the new template.
- Force a template revision conflict by editing the same template outside the workbook, then click `保存模板` / `Save template` and confirm the dialog offers overwrite, save-as, and cancel.
- Open an older workbook that has no `TemplateBindings` section and confirm download, upload, and initialize still work.
- Edit `ISDP_Setting` so `HeaderStartRow = 3`, `HeaderRowCount = 2`, and `DataStartRow = 6`, then run the hidden full-download path and confirm headers/data are written at the configured rows.
- On a sheet that already has recognizable headers, run the hidden full-download path and confirm the plugin refreshes data cells without rewriting those existing headers.
- Modify `Excel L1` or `Excel L2` in `SheetFieldMappings`, update the matching Excel header text manually, then run `部分下载` / `Partial download` or `部分上传` / `Partial upload` and confirm the column still resolves by current header text.
- Set one `single` mapping row to use both `Excel L1` and `Excel L2`, keep `HeaderRowCount = 2`, prepare matching grouped headers on the sheet, then run `部分下载` / `Partial download` and confirm the grouped-single column resolves and only the selected child cells are refreshed.
- Using the same grouped-single metadata and visible grouped headers, edit a grouped-single cell and run `部分上传` / `Partial upload`, then confirm the upload resolves that `single` field correctly and does not require converting it to a non-`single` field type.
- Keep the grouped-single headers already present on the worksheet, run the hidden full-download path, and confirm the plugin reuses that existing grouped layout instead of flattening or rewriting the recognized headers.
- Clear the worksheet header area, keep the grouped-single metadata in `ISDP_Setting`, then run the hidden full-download path and confirm regenerated headers fall back to flat child-only single headers without any grouped parent header row for that `single` field.
- Verify the task pane button and login button still work after the Ribbon Sync controls are added.
