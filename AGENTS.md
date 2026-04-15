# Repository Guidelines

## Project Overview
OfficeAgent, branded as Resy AI, is an Excel VSTO add-in with a WebView2-hosted React task pane. Keep business logic in C# and the frontend thin over `bridge.*` messages.

## Project Structure & Module Organization
`src/OfficeAgent.Core` contains orchestration, models, skills, and service contracts. `src/OfficeAgent.Infrastructure` holds HTTP clients, storage, diagnostics, and DPAPI helpers. `src/OfficeAgent.ExcelAddIn` hosts the ribbon, task pane, Excel interop, and WebView bridge. `src/OfficeAgent.Frontend` is the React/Vite UI. Tests live in `tests/OfficeAgent.Core.Tests`, `tests/OfficeAgent.Infrastructure.Tests`, `tests/OfficeAgent.ExcelAddIn.Tests`, and `tests/OfficeAgent.IntegrationTests`; `tests/mock-server` provides local SSO and API fixtures. Installer sources live in `installer/OfficeAgent.Setup`.

## Build, Test, And Development Commands
- `cd src/OfficeAgent.Frontend && npm run dev` for frontend dev.
- `cd src/OfficeAgent.Frontend && npm run build` for the bundle.
- `cd src/OfficeAgent.Frontend && npm run test` for Vitest.
- `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj`
- `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj`
- `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`
- `dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj`
- `pwsh -NoProfile -ExecutionPolicy Bypass -File installer/OfficeAgent.Setup/build.ps1` for frontend + add-in + MSI builds.
- `cd tests/mock-server && npm install && npm start` for mock services.

## Coding Style & Naming Conventions
C# uses 4-space indentation, PascalCase for public members, `I`-prefixed interfaces, camelCase private fields, and new-line braces for namespaces, classes, and methods. Prefer `string.Equals(..., StringComparison.Ordinal)` over `==`. Preserve the UI thread only where COM interop requires it. TypeScript uses 2-space indentation, single quotes, semicolons, trailing commas, type-only imports, and functional components such as `export function App() {}`. Avoid barrel exports, routing libraries, and CSS-in-JS.

## Testing Guidelines
Use xUnit for .NET and Vitest plus Testing Library for the frontend. Name .NET tests with descriptive PascalCase behavior names such as `ExecuteReturnsChatFallbackForUnknownUserInput`; frontend tests should live beside the code as `*.test.ts` or `*.test.tsx`. Use `vi.mock()` and `userEvent.setup()` where appropriate. For Excel write flows, SSO, or installer work, run `docs/vsto-manual-test-checklist.md`.

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
