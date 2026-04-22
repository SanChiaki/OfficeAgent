# Ribbon Sync Auth-Required Login Prompt Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Treat `401` and `403` business-system responses as unauthenticated Ribbon Sync failures, show a unified `当前未登录，请先登录` dialog with a `点我登录` button, and reuse the existing Ribbon login flow from all Ribbon Sync entry points.

**Architecture:** Introduce a typed authentication exception in the infrastructure layer so UI code does not depend on message matching. Add a reusable login prompt dialog service in the Excel add-in layer, then route both `AgentRibbon` project loading and `RibbonSyncController` sync actions through the same login helper that invokes the existing Ribbon SSO flow.

**Tech Stack:** C#, .NET Framework 4.8, WinForms, VSTO Excel add-in, xUnit

---

## File Structure

- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`
  Responsibility: translate `401/403` HTTP responses into a typed authentication-required exception with the new user-facing message.
- `src/OfficeAgent.Core` or `src/OfficeAgent.Infrastructure`
  Responsibility: host the shared authentication-required exception type used across UI flows.
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
  Responsibility: reuse a single login execution path for the Ribbon login button and auth-required prompts, refresh project dropdown after successful login, and stop relying on message matching.
- `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`
  Responsibility: expose a reusable auth-required dialog with `点我登录` affordance.
- `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
  Responsibility: intercept authentication-required failures from initialize/download/upload paths and invoke the login prompt instead of a generic error dialog.
- `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`
  Responsibility: lock `401` and `403` translation plus the new error message.
- `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
  Responsibility: lock Ribbon auth-required project-loading behavior and login reuse.
- `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
  Responsibility: lock controller behavior for authentication-required failures on sync actions.
- `docs/modules/ribbon-sync-current-behavior.md`
  Responsibility: document the new unauthenticated behavior and prompt.

### Task 1: Lock Connector Authentication Translation

**Files:**
- Modify: `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`
- Modify: `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`

- [ ] **Step 1: Write the failing tests**

Add tests asserting that both `401` and `403` from `/projects` throw the same authentication-required message:

```csharp
[Theory]
[InlineData(HttpStatusCode.Unauthorized)]
[InlineData(HttpStatusCode.Forbidden)]
public void GetProjectsPromptsLoginWhenProjectsEndpointRequiresAuthentication(HttpStatusCode statusCode)
{
    var connector = CurrentBusinessSystemConnector.ForTests(
        "https://example.test",
        new AuthRequiredProjectsHandler(statusCode));

    var error = Assert.Throws<AuthenticationRequiredException>(() => connector.GetProjects());

    Assert.Equal("当前未登录，请先登录", error.Message);
}
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~CurrentBusinessSystemConnectorTests
```

Expected: FAIL because the connector only treats `401` as unauthenticated and still throws `InvalidOperationException`.

- [ ] **Step 3: Implement the minimal connector change**

Introduce a typed authentication-required exception and use it for both `401` and `403`:

```csharp
if (response.StatusCode == HttpStatusCode.Unauthorized ||
    response.StatusCode == HttpStatusCode.Forbidden)
{
    throw new AuthenticationRequiredException("当前未登录，请先登录");
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run the same command from Step 2.

Expected: PASS with the new exception type and message.

### Task 2: Lock Ribbon Project Loading Prompt Behavior

**Files:**
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`

- [ ] **Step 1: Write the failing Ribbon configuration tests**

Add assertions that Ribbon project-loading auth failures use a dedicated auth-required flow and reuse the login path:

```csharp
[Fact]
public void ProjectLoadingUsesDedicatedAuthenticationPrompt()
{
    var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
        "src",
        "OfficeAgent.ExcelAddIn",
        "AgentRibbon.cs"));

    Assert.Contains("catch (AuthenticationRequiredException", ribbonCodeText, StringComparison.Ordinal);
    Assert.Contains("ShowAuthenticationRequiredPrompt", ribbonCodeText, StringComparison.Ordinal);
}

[Fact]
public void AuthenticationPromptOffersPointMeToLoginButton()
{
    var dialogCodeText = File.ReadAllText(ResolveRepositoryPath(
        "src",
        "OfficeAgent.ExcelAddIn",
        "Dialogs",
        "OperationResultDialog.cs"));

    Assert.Contains("点我登录", dialogCodeText, StringComparison.Ordinal);
}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~AgentRibbonConfigurationTests
```

Expected: FAIL because the Ribbon still catches `InvalidOperationException` and only shows a generic message box.

- [ ] **Step 3: Implement the minimal Ribbon prompt flow**

Refactor the Ribbon login button logic into a reusable method and call it from the auth-required prompt:

```csharp
private bool ExecuteRibbonLoginFlow(bool refreshProjectsAfterSuccess)
{
    // existing SsoLoginPopup flow
}
```

Add a dedicated auth-required catch branch that:

- sets the dropdown to `请先登录`
- opens the auth-required prompt
- runs `ExecuteRibbonLoginFlow(refreshProjectsAfterSuccess: true)` when the user chooses login

- [ ] **Step 4: Run the tests to verify they pass**

Run the same command from Step 2.

Expected: PASS with the new prompt entry point visible in the Ribbon source.

### Task 3: Lock Controller Authentication Prompt Routing

**Files:**
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`

- [ ] **Step 1: Write the failing controller tests**

Add tests showing initialize/download/upload auth failures use the login prompt instead of `ShowError`:

```csharp
[Fact]
public void ExecuteInitializeCurrentSheetPromptsLoginWhenAuthenticationIsRequired()
{
    var dialogService = new FakeRibbonSyncDialogService();
    var executionService = CreateExecutionService(initializeException:
        new AuthenticationRequiredException("当前未登录，请先登录"));
    var controller = CreateController(executionService, dialogService);

    controller.ExecuteInitializeCurrentSheet();

    Assert.Equal(1, dialogService.AuthenticationPromptCount);
    Assert.Empty(dialogService.ErrorMessages);
}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~RibbonSyncControllerTests
```

Expected: FAIL because the controller currently routes all exceptions to `ShowError`.

- [ ] **Step 3: Implement the minimal controller change**

Add a dedicated catch for `AuthenticationRequiredException` that calls the shared auth-required prompt and optional login callback:

```csharp
catch (AuthenticationRequiredException ex)
{
    dialogService.ShowAuthenticationRequired(ex.Message, triggerLogin: loginAction);
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run the same command from Step 2.

Expected: PASS with authentication-required failures no longer treated as generic errors.

### Task 4: Update the Module Snapshot

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`

- [ ] **Step 1: Update the user-visible auth behavior**

Document that:

- `401` and `403` both map to `请先登录`
- Ribbon Sync prompts users with `当前未登录，请先登录`
- the prompt includes `点我登录`
- successful login reloads project list but does not auto-retry sync actions

- [ ] **Step 2: Verify the docs change is present**

Run:

```powershell
Select-String -Path docs/modules/ribbon-sync-current-behavior.md -Pattern "403|点我登录|当前未登录，请先登录"
```

Expected: the updated lines are returned.
