# OfficeAgent VSTO MVP Release Notes

## What is included

- Excel VSTO add-in host with Ribbon entry and a single reusable task pane.
- WebView2-hosted React chat UI packaged locally with the add-in.
- Local session persistence and local settings persistence for `API Key`, `Base URL`, `Business Base URL`, `Model`, `SSO URL`, and `登录成功路径`.
- Live Excel selection context updates in the task pane.
- `upload_data` skill with preview, confirmation, external API call, and result message flow.
- Read-immediately and write-with-confirmation command routing.
- Structured local logging for host startup, task pane lifecycle, WebView2 initialization, bridge requests, skill routing, selection changes, and business API failures.
- WiX-based packaging that produces separate `x86` and `x64` MSI artifacts.

## Packaging outputs

- `artifacts/installer/xISDP.Setup-x86.msi`
- `artifacts/installer/xISDP.Setup-x64.msi`

## Verification summary

- `dotnet test tests\\OfficeAgent.Core.Tests\\OfficeAgent.Core.Tests.csproj`
- `dotnet test tests\\OfficeAgent.Infrastructure.Tests\\OfficeAgent.Infrastructure.Tests.csproj`
- `dotnet test tests\\OfficeAgent.ExcelAddIn.Tests\\OfficeAgent.ExcelAddIn.Tests.csproj`
- `installer\\OfficeAgent.Setup\\build.ps1`

## Operational notes

- Install the MSI that matches the target Excel bitness.
- Ensure `VSTO Runtime` and `WebView2 Runtime` are already present before MSI installation.
- `Base URL` is reserved for the LLM service. Use `Business Base URL` for business APIs such as the mock server, Ribbon Sync, and `upload_data`.
- Runtime logs are written under `%LocalAppData%\\OfficeAgent\\logs\\officeagent.log`.
