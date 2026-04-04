# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**OfficeAgent** (branded **Resy AI**) — AI-powered Excel agent. Chat-based task pane inside Excel for natural-language spreadsheet interaction. VSTO + WebView2 + React/TypeScript, .NET Framework 4.8.

## Repository Structure

```
src/
  OfficeAgent.ExcelAddIn/    — VSTO entrypoint, Ribbon, task pane, WebView2 bootstrap
  OfficeAgent.Core/          — Domain models, AgentOrchestrator, SkillRegistry, PlanExecutor
  OfficeAgent.Infrastructure/ — Excel Interop adapter, HTTP clients, file storage, DPAPI
  OfficeAgent.Frontend/      — React/TypeScript task pane UI
tests/
  OfficeAgent.Core.Tests/    — xUnit unit/integration tests
  mock-server/               — Standalone mock SSO + Business API (Node.js)
installer/OfficeAgent.Setup/ — WiX v4 MSI packaging
```

**Branches:** `main` (active), `codex/office-agent-vsto-mvp` (legacy). No active worktrees.

## Architecture

### Component Diagram

```
Excel Process
  └── VSTO Add-in (ThisAddIn)
       ├── AgentRibbon — "Resy AI" (open/close pane), "账号" (SSO login)
       └── TaskPaneController → CustomTaskPane (420px, right-docked)
            └── TaskPaneHostControl → WebView2 → React/TypeScript Frontend
```

### Bridge Protocol

Frontend ↔ Backend via `window.chrome.webview.postMessage` / `CoreWebView2.PostWebMessageAsJson`.
All messages use `bridge.*` namespace: `bridge.runAgent`, `bridge.executeExcelCommand`, `bridge.getSelectionContext`, `bridge.saveSessions`, `bridge.login`, `bridge.logout`, `bridge.getLoginStatus`.

### Agent Dispatch

| Mode | Behavior |
|---|---|
| Auto | Detect route from input |
| Skill | Direct to named skill |
| Agent | LLM planner → multi-step plan → user confirmation → execute |

### Key Flows

- **SSO login:** Configurable SSO URL + success path. Dual detection (`WebResourceResponseReceived` for HTTP 200, `NavigationCompleted` for URL path match). Cookies captured via `CookieManager`, persisted via `FileCookieStore` (DPAPI). Manual "已登录" button as fallback.
- **Conversation history:** Frontend sends last 10 turns (20 messages) as `ConversationTurn[]` via `bridge.runAgent`. Backend inserts between system prompt and current user message. Only `user`/`assistant` roles; assistant content is plain text.
- **Async chain:** `WebViewBootstrapper` → `WebMessageRouter.RouteAsync` → `AgentOrchestrator.ExecuteAsync` → `LlmPlannerClient.CompleteAsync`. `ConfigureAwait(true)` for UI thread (COM safety), `ConfigureAwait(false)` for HTTP. `WindowsFormsSynchronizationContext` installed explicitly in `InitializeAsync`.

## Code Style — C#

- **Target:** `net48`, `LangVersion: latest`, nullable disabled. No XML doc comments.
- **Naming:** Classes/Methods/Properties = PascalCase; Interfaces = `I` prefix; Privates = camelCase; Constants = PascalCase in static classes.
- **Formatting:** 4-space indent. New-line braces for namespaces/classes/methods; same-line for `if`/`for`/`try`.
- **Strings:** `string.Equals(a, b, StringComparison.Ordinal)` — never `==`.
- **Async:** `ConfigureAwait(true)` for UI continuations; `ConfigureAwait(false)` for HTTP.
- **JSON:** Newtonsoft.Json, `[JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]` on DTOs.
- **Tests:** xUnit. Descriptive PascalCase names. Nested `private sealed class` fakes.

## Code Style — TypeScript/React

- **Strict:** `strict: true`, `noUnusedLocals: true`, `noUnusedParameters: true`. No ESLint/Prettier.
- **Formatting:** 2-space indent, single quotes, semicolons, trailing commas.
- **Components:** Functional only, `export function App() {}` style (not arrow). Hooks: `useState`, `useEffect`, `useRef` only.
- **Imports:** Use `type` keyword for type-only imports. No barrel exports.
- **Async handlers:** `void handleSomething()` for fire-and-forget; `isActive` flag in `useEffect` cleanup.
- **Tests:** Vitest + Testing Library. `vi.mock()` for module mocking. `userEvent.setup()` for interactions.

## General Conventions

- All business logic in C#; frontend is thin presentation layer.
- No state management libraries, no CSS-in-JS, no routing libraries — plain CSS only.
- Chinese UI strings inline in JSX (no i18n library).
- Write operations require explicit confirmation via `ConfirmationService`.
- Brand: "Resy AI"; UI language: Chinese; technical terms (API Key, Base URL, Model, Agent, Skill) in English.

## Build & Run

```bash
# Frontend
cd src/OfficeAgent.Frontend && npm run build    # production
cd src/OfficeAgent.Frontend && npm run dev      # dev server

# .NET
dotnet test                                     # all tests
dotnet test tests/OfficeAgent.Core.Tests        # specific project
dotnet test --filter "FullyQualifiedName~Name"  # single test

# MSBuild (needs VS 2022, cert for full solution)
MSBuild OfficeAgent.sln /restore

# MSI installer (builds frontend + VSTO + MSI x86/x64)
pwsh -NoProfile -ExecutionPolicy Bypass -File installer/OfficeAgent.Setup/build.ps1

> **Important:** Always use `pwsh` (PowerShell 7.x) to run the build script. Windows PowerShell 5.1 lacks the `Cert:` PSDrive and will fail when generating the code-signing certificate.
```

## Environment & Testing

**Root `.env`:** `API_KEY`, `BASE_URL`, `MODEL`.

**Mock server:** `cd tests/mock-server && npm install && node server.js`
- SSO: `http://localhost:3100/login`, 登录成功路径 = `/rest/login`
- Business API: `http://localhost:3200` (leave API Key empty → uses SSO cookies)

## CI/CD

GitHub Actions (`.github/workflows/build-msi.yml`) on every `main` push: VS 2022 Enterprise → frontend + VSTO build → WiX v4 x86/x64 MSI → artifacts. Self-signed cert generated per build.

## Logging

`%LocalAppData%\OfficeAgent\logs\officeagent.log` — JSON, one object per line.

## Further Reading

- **`EXPERIENCE.md`** — Development pitfalls, debugging tips, lessons learned.
