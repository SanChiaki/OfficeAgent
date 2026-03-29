using System;
using System.IO;
using System.Reflection;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Infrastructure.Security;
using OfficeAgent.Infrastructure.Storage;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WebMessageRouterTests : IDisposable
    {
        private readonly string tempDirectory;

        public WebMessageRouterTests()
        {
            tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeAgent.Router.Tests", Guid.NewGuid().ToString("N"));
        }

        [Fact]
        public void SaveSettingsRejectsMissingPayloadWithoutOverwritingStoredValues()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                Model = "gpt-5-mini",
            });

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\"}");
            var settingsAfter = settingsStore.Load();

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
            Assert.Equal("secret-token", settingsAfter.ApiKey);
            Assert.Equal("https://api.internal.example", settingsAfter.BaseUrl);
            Assert.Equal("gpt-5-mini", settingsAfter.Model);
        }

        [Fact]
        public void SaveSettingsRejectsEmptyObjectPayloadWithoutOverwritingStoredValues()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                Model = "gpt-5-mini",
            });

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\",\"payload\":{}}");
            var settingsAfter = settingsStore.Load();

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
            Assert.Equal("secret-token", settingsAfter.ApiKey);
            Assert.Equal("https://api.internal.example", settingsAfter.BaseUrl);
            Assert.Equal("gpt-5-mini", settingsAfter.Model);
        }

        [Fact]
        public void SaveSettingsRejectsPayloadWithoutApiKeyWithoutOverwritingStoredValues()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                Model = "gpt-5-mini",
            });

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\",\"payload\":{\"baseUrl\":\"https://api.internal.example\",\"model\":\"gpt-5-mini\"}}");
            var settingsAfter = settingsStore.Load();

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
            Assert.Equal("secret-token", settingsAfter.ApiKey);
            Assert.Equal("https://api.internal.example", settingsAfter.BaseUrl);
            Assert.Equal("gpt-5-mini", settingsAfter.Model);
        }

        [Fact]
        public void GetSettingsRejectsEmptyObjectPayload()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.getSettings\",\"requestId\":\"req-1\",\"payload\":{}}");

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
        }

        [Fact]
        public void GetSelectionContextReturnsCurrentSelectionContext()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var selectionContextService = new FakeExcelContextService(new SelectionContext
            {
                WorkbookName = "Quarterly Report.xlsx",
                SheetName = "Sheet1",
                Address = "A1:C4",
                RowCount = 4,
                ColumnCount = 3,
                IsContiguous = true,
                HeaderPreview = new[] { "Name", "Region", "Amount" },
                SampleRows = new[]
                {
                    new[] { "Project A", "CN", "42" },
                    new[] { "Project B", "US", "36" },
                },
                WarningMessage = null,
            });

            var router = CreateRouter(sessionStore, settingsStore, selectionContextService);
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.getSelectionContext\",\"requestId\":\"req-1\"}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"workbookName\":\"Quarterly Report.xlsx\"", responseJson);
            Assert.Contains("\"sheetName\":\"Sheet1\"", responseJson);
            Assert.Contains("\"address\":\"A1:C4\"", responseJson);
        }

        public void Dispose()
        {
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }

        private static object CreateRouter(FileSessionStore sessionStore, FileSettingsStore settingsStore)
        {
            return CreateRouter(sessionStore, settingsStore, new FakeExcelContextService(SelectionContext.Empty("No selection available.")));
        }

        private static object CreateRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService selectionContextService)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var routerType = addInAssembly.GetType(
                "OfficeAgent.ExcelAddIn.WebBridge.WebMessageRouter",
                throwOnError: true);
            return Activator.CreateInstance(
                routerType,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { sessionStore, settingsStore, selectionContextService },
                culture: null);
        }

        private static string InvokeRoute(object router, string requestJson)
        {
            var routeMethod = router.GetType().GetMethod(
                "Route",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (string)routeMethod.Invoke(router, new object[] { requestJson });
        }

        private static string ResolveAddInAssemblyPath()
        {
            return Path.GetFullPath(
                Path.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
        }

        private sealed class FakeExcelContextService : IExcelContextService
        {
            private readonly SelectionContext selectionContext;

            public FakeExcelContextService(SelectionContext selectionContext)
            {
                this.selectionContext = selectionContext;
            }

            public SelectionContext GetCurrentSelectionContext()
            {
                return selectionContext;
            }
        }
    }
}
