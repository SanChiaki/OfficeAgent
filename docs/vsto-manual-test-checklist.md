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

- Open Settings and save `API Key`, `Base URL`, `Business Base URL`, `Model`, `API Format`, `SSO URL`, and `登录成功路径` / `Login success path`.
- Confirm `Base URL` stays reserved for the LLM endpoint and `Business Base URL` points to the business API or mock server.
- Confirm the analytics URL is not shown in Settings; it is an internal hidden configuration.
- Confirm `API Format = OpenAI Compatible` keeps existing OpenAI-compatible planner behavior, then switch to `API Format = Anthropic Messages` with an Anthropic Messages-compatible endpoint and confirm ordinary planner chat still returns a valid response.
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
- For the steps below, use the localized host labels for the current Excel UI. Key pairs: `初始化当前表` / `Initialize sheet`, `AI映射列` / `AI map columns`, `应用配置` / `Apply Setting`, `保存配置` / `Save Setting`, `另存配置` / `Save as Setting`, `下载` / `Download`, `上传` / `Upload`, `登录` / `Login`, `退出` / `Logout`, `先选择项目` / `Select project`, `请先登录` / `Sign in first`, `无可用项目` / `No projects available`.
- Bind a blank worksheet through the Ribbon project dropdown and confirm the layout dialog appears with defaults `HeaderStartRow = 1`, `HeaderRowCount = 2`, `DataStartRow = 3`.
- Confirm the layout dialog and enter custom values, then verify `xISDP_Setting` writes one `SheetBindings` row with the user-entered layout values.
- Confirming project selection should still not auto-initialize the current sheet; `SheetFieldMappings` remains unchanged until `初始化当前表` / `Initialize sheet` is clicked.
- Open `xISDP_Setting` and confirm it uses one worksheet with three readable sections: `TemplateBindings` on top, `SheetBindings` in the middle, `SheetFieldMappings` below, each with a title row, a header row, and data rows.
- Confirm `SheetFieldMappings` displays headers in this order: `HeaderType`, `ISDP L1`, `ISDP L2`, `Excel L1`, `Excel L2`, `HeaderId`, `ApiFieldKey`, `IsIdColumn`, `ActivityId`, `PropertyId`.
- Confirm there are two blank separator rows between each adjacent metadata section, and that metadata is no longer stored as flattened `tableName + values` rows.
- Switch to a worksheet with existing binding metadata and confirm the Ribbon dropdown automatically rehydrates that project as `ProjectId-DisplayName` instead of showing the localized empty placeholder (`先选择项目` / `Select project`).
- Save a workbook with `xISDP_Setting` as the active sheet, reopen Excel from the desktop shortcut, and confirm the Ribbon dropdown shows the localized empty placeholder (`先选择项目` / `Select project`) unless `SheetBindings` contains an explicit binding row for `xISDP_Setting`.
- Switch from a bound business sheet to `xISDP_Setting` and confirm the Ribbon dropdown clears back to the localized empty placeholder (`先选择项目` / `Select project`) when `xISDP_Setting` itself has no binding.
- Switch to a worksheet without binding metadata and confirm the Ribbon dropdown shows the localized empty placeholder (`先选择项目` / `Select project`).
- Open two workbooks in the same Excel process, bind `Sheet1` in each workbook to different projects, switch back and forth between the two files, and confirm the Ribbon dropdown plus download/upload behavior always follow the active workbook's own `xISDP_Setting` metadata.
- Open an older workbook that contains `ISDP_Setting` but not `xISDP_Setting`, trigger any Ribbon Sync metadata read such as project refresh, and confirm the worksheet is automatically renamed to `xISDP_Setting` with its existing metadata preserved.
- On a sheet that already has binding metadata, switch to another project and confirm the layout dialog defaults reuse the current sheet's saved layout values.
- Increase Windows or Office UI font scaling, reopen the project layout dialog, and confirm all labels, inputs, and buttons remain fully visible without overlap.
- Under both Chinese Excel and English Excel, trigger the unauthenticated path and confirm the native auth-required prompt plus login button text are localized for that host language.
- Reselect the already bound project (`same systemKey + projectId`) and confirm no layout dialog appears and `SheetBindings` is not rewritten.
- Cancel the layout dialog while switching projects and confirm both `xISDP_Setting` binding data and Ribbon dropdown project stay unchanged.
- After switching to another project and confirming the dialog, verify old `SheetFieldMappings` are cleared; before clicking `初始化当前表` / `Initialize sheet`, running download/upload should report that the current sheet is not initialized.
- Enter invalid values in the layout dialog (for example overlaps between header/data regions) and confirm validation error is shown while keeping the dialog open.
- Start Excel while unauthenticated against a protected project API and confirm the project dropdown shows the localized sign-in-required status (`请先登录` / `Sign in first`).
- Close the automatic sign-in-required prompt without logging in, then open or activate another workbook in the same Excel process and confirm the prompt is not shown again; the project dropdown should remain at `请先登录` / `Sign in first`.
- Confirm the account group shows small regular `登录` / `Login` and `退出` / `Logout` buttons; while unauthenticated, `登录` / `Login` is enabled and `退出` / `Logout` is disabled.
- Complete SSO login from the Ribbon and confirm `登录` / `Login` becomes disabled, `退出` / `Logout` becomes enabled, and the project dropdown can load projects.
- Click `退出` / `Logout` and confirm the project dropdown switches to `请先登录` / `Sign in first`, `登录` / `Login` becomes enabled, `退出` / `Logout` becomes disabled, and reopening the project dropdown does not automatically show the login prompt or reuse the previous cookie.
- Configure the project API to return an empty array and confirm the project dropdown shows the localized empty-project status (`无可用项目` / `No projects available`).
- Click `初始化当前表` / `Initialize sheet` on a sheet that already contains business cells and confirm only `xISDP_Setting` changes; the business area should remain untouched.
- Configure `Base URL`, `API Key`, `Model`, and `API Format` in the existing Settings UI for an OpenAI-compatible model endpoint, then use `AI映射列` / `AI map columns` and confirm it reuses those settings rather than asking for a separate AI mapping configuration.
- Switch `API Format` to `Anthropic Messages` with a compatible model endpoint, run `AI映射列` / `AI map columns`, and confirm the preview appears after the Anthropic Messages response completes.
- Run `AI映射列` / `AI map columns` against a slow model response and confirm a native processing dialog appears above Excel with `中止` / `Abort`; click it and verify the model call is cancelled, no preview is shown, and Excel control returns after cancellation.
- While the AI mapping preview, completion, or error dialog is open, click/focus the Excel window and confirm the dialog remains owned by Excel rather than disappearing behind the workbook.
- On an initialized sheet, rename one visible business header so it no longer matches `ISDP L1` / `ISDP L2`, click `AI映射列` / `AI map columns`, and confirm the preview dialog appears before any metadata is saved.
- Confirm headers that already match `SheetFieldMappings.Excel L1 / Excel L2`, including the `ID` column, are not shown as accepted rows in the AI mapping preview; if all scanned headers already match, the model should not be called and the command should report no accepted mappings.
- Confirm the AI mapping preview shows only four columns: apply checkbox (`是否修改` / `Apply`), Excel letter column, current actual header `L1/L2`, and matched ISDP header `L1/L2`.
- In the AI mapping preview, cancel the dialog and confirm `xISDP_Setting.SheetFieldMappings` remains unchanged.
- In the AI mapping preview, clear the apply checkbox for one recommendation, confirm the dialog, and verify that unchecked recommendation is not written to `xISDP_Setting.SheetFieldMappings`.
- Run `AI映射列` / `AI map columns` again, confirm the preview, and verify only `Excel L1` / `Excel L2` are updated for accepted recommendations; `ISDP L1` / `ISDP L2`, `HeaderId`, `ApiFieldKey`, `IsIdColumn`, `ActivityId`, and `PropertyId` must remain unchanged.
- Include one model recommendation that writes an L1/L2 combination different from the original field type or visible header depth, such as mapping a single-level actual header to an `activityProperty` row or mapping a two-level actual header to a `single` row, and confirm accepted recommendations still update only `Excel L1` / `Excel L2`.
- Prepare headers outside the current selection but inside the configured header rows, select an unrelated cell, run `AI映射列` / `AI map columns`, and confirm the preview still includes the full configured header area rather than only the selected cell.
- Include one low-confidence or unmatched actual header in the model response or test fixture and confirm it is not shown as an applicable mapping and is skipped after confirmation.
- Click `下载` / `Download` and `上传` / `Upload` and confirm each action uses a native Office/WinForms confirmation dialog instead of the task pane.
- Configure `下载` / `Download` to return no matching rows and confirm it shows a no-matching-records info message instead of a confirmation dialog with field count, and that no worksheet cells are written or cleared.
- Select one or more full worksheet columns that include managed non-ID fields, click `下载` / `Download`, and confirm Excel stays responsive, only rows with `row_id` from `DataStartRow` through the last used row are requested, and only selected managed columns are refreshed.
- Select the whole worksheet, click `下载` / `Download`, and confirm Excel stays responsive, the ID column is not overwritten, and all recognized non-ID managed fields are refreshed only for rows with `row_id`.
- Select two non-contiguous areas in different rows and columns, click `下载` / `Download`, and confirm only the exact selected cells are refreshed; cells at the unselected row/column intersections must remain unchanged.
- Before `下载` / `Download`, put an old value in one selected managed non-ID cell, run the download, and confirm `xISDP_Log` is created with columns `Key`, `Header`, `Change Mode`, `New Value`, `Old Value`, `Changed At`; the new row should show `Change Mode = Download`, `Old Value` as the overwritten Excel value, and `New Value` as the downloaded value.
- Edit one managed non-ID cell, run `上传` / `Upload`, and confirm `xISDP_Log` appends one `Change Mode = Upload` row using the user's pre-edit Excel value as `Old Value` and the uploaded cell value as `New Value`.
- Force an `上传` / `Upload` failure from the mock server or API, then confirm no new `Upload` row is added to `xISDP_Log`; retry successfully and confirm the original pre-edit value is still used.
- Add more than 2000 sync log rows through repeated upload/download validation or seeded workbook data, then trigger another logged sync and confirm `xISDP_Log` keeps only the latest 2000 data rows plus the header row.
- Confirm the Ribbon includes a dedicated `配置` / `Setting` group with `应用配置` / `Apply Setting`, `保存配置` / `Save Setting`, and `另存配置` / `Save as Setting`.
- Confirm all Ribbon buttons display icons that match their action semantics. `关于` / `About` may use a host-generated custom icon for the version reminder red dot; other Ribbon buttons should display Office built-in icons. `初始化当前表` / `Initialize sheet`, `AI映射列` / `AI map columns`, all buttons in the `配置` / `Setting` group, and the account `登录` / `Login` plus `退出` / `Logout` buttons should use the small regular button layout; data sync and help command buttons should remain in the large icon-above-label layout.
- Confirm the `xISDP AI` group task-pane button shows only its icon and does not display the `Open` label.
- Confirm the Ribbon includes one `数据同步` / `Data sync` group containing `下载` / `Download` and `上传` / `Upload`, and that there is no `全量下载`, `全量上传`, or `增量上传` button.
- Confirm the Ribbon includes a `帮助` / `Help` group with `文档` / `Documentation` and `关于` / `About`; `文档` / `Documentation` opens `https://github.com/SanChiaki/OfficeAgent` in the default browser, and `关于` / `About` shows version and build information.
- 更新提醒：配置内部更新 manifest URL，使其返回 `Content-Type: application/octet-stream` 的 JSON 字节流，且 `latestVersion` 高于当前 `VersionInfo.AppVersion`。Debug 和 Release 均应使用同一套检查逻辑；打开 Excel 后确认 `关于` / `About` 图标显示红点；点击 `关于` 后确认显示当前版本、最新版本、更新摘要和下载入口，且不显示发布说明按钮；点击 `忽略此版本` / `Ignore this version` 后确认红点消失，再次点击 `关于` 仍能看到该版本的更新信息和下载入口；把 manifest 提高到更高版本后确认红点重新出现。
- 更新检查失败隔离：让更新 manifest URL 断开或返回非法 JSON，重新打开 Excel，确认 Ribbon、任务窗格、登录、下载、上传和模板操作仍可用，且没有更新失败弹窗。
- 未配置更新源隔离：清空或删除更新 manifest URL 配置后重新打开 Excel，确认不会请求更新 manifest URL，且 Ribbon、任务窗格、登录、下载、上传和模板操作仍可用。
- In the same project, save two different local templates and confirm `应用配置` / `Apply Setting` can list both.
- Apply one template and confirm `TemplateBindings` updates to the selected template while `SheetBindings` / `SheetFieldMappings` are expanded into the current sheet.
- Manually edit `xISDP_Setting` field mapping text after applying a template, click `保存配置` / `Save Setting`, then reapply that template and confirm the edited mapping is preserved.
- With a sheet already bound to a template, use `另存配置` / `Save as Setting`, confirm the new template name appears in the local template list, and confirm the current sheet's `TemplateBindings.TemplateId` switches to the new template.
- Force a template revision conflict by editing the same template outside the workbook, then click `保存配置` / `Save Setting` and confirm the dialog offers overwrite, save-as, and cancel.
- Open an older workbook that has no `TemplateBindings` section and confirm download, upload, and initialize still work.
- Edit `xISDP_Setting` so `HeaderStartRow = 3`, `HeaderRowCount = 2`, and `DataStartRow = 6`, then run the hidden full-download path and confirm headers/data are written at the configured rows.
- On a sheet that already has recognizable headers, run the hidden full-download path and confirm the plugin refreshes data cells without rewriting those existing headers.
- Modify `Excel L1` or `Excel L2` in `SheetFieldMappings`, update the matching Excel header text manually, then run `下载` / `Download` or `上传` / `Upload` and confirm the column still resolves by current header text.
- Set `HeaderRowCount = 1`, keep an `activityProperty` row's visible single-row name in `Excel L1` with `Excel L2` empty, then run `下载` / `Download` or `上传` / `Upload` and confirm the activity property column resolves.
- Set one `single` mapping row to use both `Excel L1` and `Excel L2`, keep `HeaderRowCount = 2`, prepare matching grouped headers on the sheet, then run `下载` / `Download` and confirm the grouped-single column resolves and only the selected child cells are refreshed.
- Using the same grouped-single metadata and visible grouped headers, edit a grouped-single cell and run `上传` / `Upload`, then confirm the upload resolves that `single` field correctly and does not require converting it to a non-`single` field type.
- Keep the grouped-single headers already present on the worksheet, run the hidden full-download path, and confirm the plugin reuses that existing grouped layout instead of flattening or rewriting the recognized headers.
- Clear the worksheet header area, keep the grouped-single metadata in `xISDP_Setting`, then run the hidden full-download path and confirm regenerated headers fall back to flat child-only single headers without any grouped parent header row for that `single` field.
- Verify the task pane button and account buttons still work after the Ribbon Sync controls are added.

## Analytics

- Start `tests/mock-server` and set hidden `AnalyticsUrl = http://localhost:3200/insertLog`.
- Confirm hidden `AnalyticsUrl` is stored in `%LocalAppData%\OfficeAgent\settings.json`, not in the task-pane Settings UI.
- Clear existing events with `DELETE http://localhost:3200/analytics/logs`.
- Click Ribbon initialize, download, and upload; confirm `/analytics/logs` contains `ribbon.initialize.*`, `ribbon.download.*`, and `ribbon.upload.*` events with `projectId` and `projectName`.
- In the task pane, send a prompt, save settings, and confirm/cancel a preview card; confirm `/analytics/logs` contains `panel.*` events.
- After SSO login, confirm each mock analytics log entry includes the login cookie under `cookies`.
- Confirm the analytics `payload.answer` content does not contain API keys, cookies, raw prompt text, cell values, or business API request/response bodies.
