# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OfficeAgent (branded as **Resy AI**) is an AI-powered Excel agent — a chat-based task pane inside Excel that lets users interact with spreadsheet data through natural language. The project pivoted from an Office.js Add-in to a VSTO approach and lives across two git worktrees.

## Repository Structure

The repo root contains design docs and environment config. All source code lives in worktrees under `.worktrees/`:

- **`.worktrees/office-agent-mvp/`** — Original Office.js Add-in (React + Office.js, deprecated)
- **`.worktrees/office-agent-vsto-mvp/`** — Active VSTO implementation (C# + WebView2 + React)
- **`docs/`** — Design specs and implementation plans (primary language: Chinese)

**Branches:**
- `codex/office-agent-vsto-mvp` — active development (in the VSTO worktree)
- `codex/vsto-redesign-plan` — docs/planning (in the repo root)
- `master` — default branch

The VSTO worktree pushes via `git push office-agent codex/office-agent-vsto-mvp` (HTTPS remote).

## VSTO Architecture

```
Excel Process
  └── VSTO Add-in (ThisAddIn)
       ├── AgentRibbon (Ribbon with two groups)
       │    ├── Group "Resy AI" — Open/Close task pane button
       │    └── Group "账号" — SSO login button
       └── TaskPaneController
            └── CustomTaskPane (420px, docked right)
                 └── TaskPaneHostControl (WinForms UserControl)
                      └── WebView2 (Edge embedded browser)
                           └── React/TypeScript Frontend
```

**Backend layers (C# / .NET Framework 4.8):**

| Layer | Project | Responsibility |
|---|---|---|
| Add-in host | `OfficeAgent.ExcelAddIn` | VSTO entrypoint, Ribbon, task pane lifecycle, WebView2 bootstrap, Excel event bridge |
| Core | `OfficeAgent.Core` | Domain models, AgentOrchestrator, SkillRegistry, ConfirmationService, PlanExecutor |
| Infrastructure | `OfficeAgent.Infrastructure` | Excel Interop adapter, HTTP clients (LLM, Business API), file storage, DPAPI encryption |

**WebView2 JS/.NET Bridge:** Frontend calls `window.chrome.webview.postMessage(json)`, routed by `WebMessageRouter` to the appropriate handler. Responses come back via `CoreWebView2.PostWebMessageAsJson`. All bridge messages use the `bridge.*` namespace (e.g., `bridge.runAgent`, `bridge.executeExcelCommand`, `bridge.getSelectionContext`, `bridge.saveSessions`, `bridge.login`, `bridge.logout`, `bridge.getLoginStatus`).

**Agent dispatch modes:** Auto (detect route from input), Skill (direct to named skill), Agent (LLM planner for multi-step plans with user confirmation).

**SSO login:** Users configure an SSO URL and an optional login success path (登录成功路径) in settings. Clicking "登录" in the task pane opens a WebView2 popup (`SsoLoginPopup`) to the SSO page. Login success is detected through two parallel channels: (1) `WebResourceResponseReceived` fires when a page fetch/XHR request matching the success path returns HTTP 200; (2) `NavigationCompleted` fires when the page navigates to a URL whose path contains the success path marker. Once detected, the popup captures cookies via `CookieManager.GetCookiesAsync`, stores them in a shared `CookieContainer`, and persists them via `FileCookieStore` (DPAPI-encrypted). A manual "已登录" button at the bottom of the popup also serves as a fallback. `BusinessApiClient` uses the same `CookieContainer` to send cookies with business API requests. The popup initializes WebView2 via `InitializeAsync()` before `ShowDialog()` to avoid async void continuation issues in modal dialogs.

**Conversation history:** The agent sends multi-turn conversation history to the LLM. The frontend extracts the last 10 turns (20 messages) from `sessionThreads` and passes them as `ConversationTurn[]` via `bridge.runAgent`. The backend inserts history between the system prompt and current user message in the OpenAI-compatible chat completions API. History flows: `AgentCommandEnvelope.ConversationHistory` → `PlannerRequest.ConversationHistory` → `LlmPlannerClient.BuildChatMessages`. Only `user` and `assistant` role messages are included; assistant content is plain text (`AssistantMessage`), not full structured JSON.

**Async architecture:** LLM calls use a full async chain to avoid blocking the Excel UI thread. The chain is: `WebViewBootstrapper` (async void event handler) → `WebMessageRouter.RouteAsync` → `AgentOrchestrator.ExecuteAsync` → `LlmPlannerClient.CompleteAsync`. `ConfigureAwait(true)` in the orchestrator keeps continuations on the UI thread for safe COM access; `ConfigureAwait(false)` in the HTTP layer for efficiency. A `WindowsFormsSynchronizationContext` is installed explicitly in `WebViewBootstrapper.InitializeAsync` because VSTO does not call `Application.Run`.

**MSI installer:** Auto-increments version from `git rev-list --count HEAD` (e.g. `1.0.28`). `MajorUpgrade` with a stable `UpgradeCode` enables seamless overwrite-install. Prerequisites: VSTO Runtime 4.0, WebView2 Runtime.

## Build & Run Commands

### VSTO Solution (`.worktrees/office-agent-vsto-mvp/`)

```bash
# Build .NET projects (Core + Infrastructure only, no signing needed)
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" src/OfficeAgent.Core/OfficeAgent.Core.csproj -p:Configuration=Release
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" src/OfficeAgent.Infrastructure/OfficeAgent.Infrastructure.csproj -p:Configuration=Release

# Build full VSTO solution (requires ClickOnce manifest signing certificate)
MSBuild OfficeAgent.sln /restore

# Build and run tests
dotnet test

# Run a specific test project
dotnet test tests/OfficeAgent.Core.Tests

# Build the React frontend
cd src/OfficeAgent.Frontend && npm run build

# Dev server for frontend only
cd src/OfficeAgent.Frontend && npm run dev

# Build MSI installer (builds frontend + VSTO add-in + MSI for x86/x64)
powershell installer/OfficeAgent.Setup/build.ps1
```

The .NET solution requires Visual Studio 2022 and targets .NET Framework 4.8. Tests use xUnit.

### Office.js MVP (`.worktrees/office-agent-mvp/` — deprecated)

```bash
npm run build       # tsc --noEmit && vite build
npm test            # vitest run
npm run test:watch  # vitest --watch
npm run dev         # HTTPS dev server on localhost:3000
```

## Key Design Decisions

- Write operations (Excel commands, skills) require explicit user confirmation via `ConfirmationService`
- Agent plans are previewed before execution; plan steps are validated against a whitelist of supported actions
- Settings and sessions are persisted as local JSON files; secrets use DPAPI encryption
- Session management (create, rename, delete, switch) runs entirely in the frontend; state is persisted to backend via `bridge.saveSessions` with 1-second debounce
- Auto-rename: new sessions titled "New chat" are automatically renamed after the first user message (first 20 chars)
- The frontend is a thin presentation layer — all business logic lives in C# Core/Infrastructure
- MSI installer checks for VSTO Runtime and WebView2 Runtime prerequisites; installs to `%LocalAppData%\OfficeAgent\`
- UI language is Chinese; technical terms (API Key, Base URL, Model, Excel, Agent, Skill) kept in English
- Brand name in UI is "Resy AI"; Ribbon button shows the product logo from embedded resources

## Environment Configuration

The root `.env` file configures: `API_KEY`, `BASE_URL`, `MODEL` (used by the LLM planner).

## Mock Server (Testing)

A standalone mock SSO + Business API server lives in `tests/mock-server/`:

```bash
cd tests/mock-server && npm install && node server.js
```

| Service | Port | Endpoints |
|---|---|---|
| SSO Login | 3100 | `GET /login` (form page); `POST /rest/login` (returns 200 + Set-Cookie) |
| Business API | 3200 | `GET /api/performance`, `GET /api/performance/:name`, `POST /api/performance`, `POST /upload_data`, `GET /api/download/:projectName` |

Configure the add-in: SSO URL = `http://localhost:3100/login`, 登录成功路径 = `/rest/login`, Base URL = `http://localhost:3200`, API Key = leave empty (uses SSO cookies).

## Logging

Runtime logs are written to `%LocalAppData%\OfficeAgent\logs\officeagent.log` (JSON, one object per line). Check this file for diagnosing bridge errors, LLM timeouts, or WebView2 failures.

## CI/CD

GitHub Actions workflow (`.github/workflows/build-msi.yml`) builds MSI installers on every push to `codex/office-agent-vsto-mvp`. The workflow runs on `windows-latest` (VS 2022 Enterprise), builds the frontend and VSTO add-in, packages x86/x64 MSI installers via WiX v4, and uploads them as artifacts. VSTO ClickOnce manifest signing uses a temporary self-signed certificate generated per build. The `.dotnet-tools.json` manifest declares the `wix` v4.0.5 .NET tool.

## Further Reading

- **`EXPERIENCE.md`** — Development pitfalls, debugging tips, and lessons learned from building this VSTO add-in.
