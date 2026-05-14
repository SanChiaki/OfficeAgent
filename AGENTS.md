# Repository Guidelines

## Project Overview
OfficeAgent, branded as Resy AI, is an Excel VSTO add-in with a WebView2-hosted React task pane. Keep business logic in C# and the frontend thin over `bridge.*` messages.

## Worktree Model

This repository uses a worktree-first workflow.

The main repo directory is only for:
- git control operations
- worktree management
- merge/rebase operations

Do not develop directly inside the main repo.

Each agent must:
- work in its own branch
- work in its own worktree
- stay within its assigned task scope

---

## Starting Work

Before editing, confirm where you are:

```sh
git rev-parse --show-toplevel
git status --short
```

If you are in the main repo directory, create or switch to a dedicated worktree before making changes.

Recommended naming:
- branch: `agent/<task-slug>` or `codex/<task-slug>`
- worktree: `../<repo-name>-worktrees/<task-slug>`

Example:

```sh
git fetch
git worktree add ../<repo-name>-worktrees/<task-slug> -b agent/<task-slug> origin/main
cd ../<repo-name>-worktrees/<task-slug>
```

Reuse an existing branch or worktree only when it clearly belongs to the same task.

---

## Working In A Worktree

After setup, stay inside the assigned worktree.

Do not modify files from:
- the main repo working directory
- another agent's worktree
- unrelated task branches

Keep changes limited to the assigned task. Avoid touching high-conflict shared files unless necessary:
- package.json
- lockfiles
- shared configs
- CI workflows
- database schemas

If existing changes are present in the worktree, assume they may belong to a human or another agent. Do not overwrite, revert, or rebase them away blindly.

---

## Finishing Work

Before handing off:

```sh
git status --short
git diff --stat
```

Summarize:
- current branch
- current worktree path
- changed files
- any high-conflict files touched
- whether the branch was rebased or merged before handoff

## Project Structure & Module Organization
`src/OfficeAgent.Core` contains orchestration, models, skills, and service contracts. `src/OfficeAgent.Infrastructure` holds HTTP clients, storage, diagnostics, and DPAPI helpers. `src/OfficeAgent.ExcelAddIn` hosts the ribbon, task pane, Excel interop, and WebView bridge. `src/OfficeAgent.Frontend` is the React/Vite UI. Tests live in `tests/OfficeAgent.Core.Tests`, `tests/OfficeAgent.Infrastructure.Tests`, `tests/OfficeAgent.ExcelAddIn.Tests`, and `tests/OfficeAgent.IntegrationTests`; `tests/mock-server` provides local SSO and API fixtures. Installer sources live in `installer/OfficeAgent.Setup` and `installer/OfficeAgent.SetupBundle`.

## Build, Test, And Development Commands
- `cd src/OfficeAgent.Frontend && npm run dev` for frontend dev.
- `cd src/OfficeAgent.Frontend && npm run build` for the bundle.
- `cd src/OfficeAgent.Frontend && npm run test` for Vitest.
- `pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1` for the recommended dev refresh flow: rebuild frontend `dist`, rebuild the Debug VSTO add-in, and refresh Excel's local registration for the development add-in manifest.
- `pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel` to force-close all running `EXCEL.EXE` processes before rebuilding when validating Ribbon, VSTO startup, or sheet event changes.
- `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj`
- `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj`
- `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`
- `dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj`
- `pwsh -NoProfile -ExecutionPolicy Bypass -File installer/OfficeAgent.Setup/build.ps1` for frontend + add-in + MSI + offline `xISDP.Setup.exe` builds.
- `cd tests/mock-server && npm install && npm start` for mock services.

Recommended development flow:

- Frontend-only changes: run `cd src/OfficeAgent.Frontend && npm run build`, then reopen the task pane.
- Add-in / Ribbon / Excel interop changes: run `pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel`, then reopen Excel from the normal user launch path. This rebuild also refreshes the local `OfficeAgent.ExcelAddIn` Excel registration to the repo `bin/Debug` manifest.
- Installer validation only: stage the offline prerequisite installers under `installer/OfficeAgent.SetupBundle/prereqs/`, run `pwsh -NoProfile -ExecutionPolicy Bypass -File installer/OfficeAgent.Setup/build.ps1`, then validate `artifacts/installer/xISDP.Setup.exe`.

## Coding Style & Naming Conventions
C# uses 4-space indentation, PascalCase for public members, `I`-prefixed interfaces, camelCase private fields, and new-line braces for namespaces, classes, and methods. Prefer `string.Equals(..., StringComparison.Ordinal)` over `==`. Preserve the UI thread only where COM interop requires it. TypeScript uses 2-space indentation, single quotes, semicolons, trailing commas, type-only imports, and functional components such as `export function App() {}`. Avoid barrel exports, routing libraries, and CSS-in-JS.

## Testing Guidelines
Use xUnit for .NET and Vitest plus Testing Library for the frontend. Name .NET tests with descriptive PascalCase behavior names such as `ExecuteReturnsChatFallbackForUnknownUserInput`; frontend tests should live beside the code as `*.test.ts` or `*.test.tsx`. Use `vi.mock()` and `userEvent.setup()` where appropriate. For Excel write flows, SSO, or installer work, run `docs/vsto-manual-test-checklist.md`.

After changing code, do not report testing as successful until the relevant build validation also passes. Choose the build command that matches the touched area, such as frontend `npm run build`, Debug VSTO add-in build / `eng/Dev-RefreshExcelAddIn.ps1`, or installer build. If a build cannot be run, state that clearly and do not mark verification as fully successful.

## Commit & Pull Request Guidelines
Follow the existing Conventional Commit style: `feat:`, `fix:`, `docs:`, `build:`, and `test:`. Keep each commit scoped to one logical change. PRs should summarize user-visible impact, list verification commands, link the related issue when available, and include screenshots for task-pane or installer UI changes.

## Security & Configuration Tips
Root `.env` values include `API_KEY`, `BASE_URL`, and `MODEL`; do not commit secrets or local settings. If the business API uses SSO cookies, leave the API key empty and use `tests/mock-server`. Use `pwsh` for installer builds because Windows PowerShell 5.1 cannot create the signing certificate. Treat `src/OfficeAgent.ExcelAddIn/Properties/Version.g.cs` as generated, and use `%LocalAppData%\\OfficeAgent\\logs\\officeagent.log` for diagnostics.

## Module Documentation Entry
Before modifying a feature module or starting a fresh implementation session, read [docs/module-index.md](docs/module-index.md) first.

Current recommended flow:

1. Open [docs/module-index.md](docs/module-index.md)
2. Jump to the target module's current behavior snapshot under `docs/modules/`
3. Only after reading the snapshot, continue into design docs, plans, test checklists, or integration guides as needed

For Ribbon Sync connector changes, also keep [docs/ribbon-sync-real-system-integration-guide.md](docs/ribbon-sync-real-system-integration-guide.md) aligned with the actual registration and routing model.

For Ribbon Sync specifically, start with:

- [docs/modules/ribbon-sync-current-behavior.md](docs/modules/ribbon-sync-current-behavior.md)

When a module's user-visible behavior changes, update its `docs/modules/*-current-behavior.md` file in the same change whenever practical.
