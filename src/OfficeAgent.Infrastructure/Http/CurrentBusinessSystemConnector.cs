using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Authentication;
using System.Text;
using Newtonsoft.Json;
using OfficeAgent.Core;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class CurrentBusinessSystemConnector : ISystemConnector
    {
        private const string CurrentSystemKey = "current-business-system";
        private const int DefaultHeaderStartRow = 1;
        private const int DefaultHeaderRowCount = 2;
        private const int DefaultDataStartRow = 3;

        private static readonly IReadOnlyDictionary<string, string> PropertyLabels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["name"] = "名称",
            ["start"] = "开始时间",
            ["end"] = "结束时间",
        };

        private sealed class SchemaHeadWrapper
        {
            [JsonProperty("headList")]
            public CurrentBusinessHeadDefinition[] HeadList { get; set; } = Array.Empty<CurrentBusinessHeadDefinition>();
        }

        private readonly CurrentBusinessSchemaMapper schemaMapper;
        private readonly CurrentBusinessFieldMappingSeedBuilder fieldMappingSeedBuilder;
        private readonly Func<AppSettings> loadSettings;
        private readonly HttpClient httpClient;

        public CurrentBusinessSystemConnector(Func<AppSettings> loadSettings, HttpClient httpClient = null, CookieContainer cookieContainer = null)
            : this(loadSettings ?? throw new ArgumentNullException(nameof(loadSettings)), new CurrentBusinessSchemaMapper(PropertyLabels), new CurrentBusinessFieldMappingSeedBuilder(PropertyLabels), httpClient, handler: null, cookieContainer)
        {
        }

        private CurrentBusinessSystemConnector(
            Func<AppSettings> loadSettings,
            CurrentBusinessSchemaMapper schemaMapper,
            CurrentBusinessFieldMappingSeedBuilder fieldMappingSeedBuilder,
            HttpClient httpClient,
            HttpMessageHandler handler,
            CookieContainer cookieContainer)
        {
            if (schemaMapper == null)
            {
                throw new ArgumentNullException(nameof(schemaMapper));
            }

            if (fieldMappingSeedBuilder == null)
            {
                throw new ArgumentNullException(nameof(fieldMappingSeedBuilder));
            }

            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));
            this.schemaMapper = schemaMapper;
            this.fieldMappingSeedBuilder = fieldMappingSeedBuilder;
            if (httpClient != null)
            {
                this.httpClient = httpClient;
            }
            else
            {
                HttpMessageHandler handlerToUse = handler ?? new HttpClientHandler
                {
                    CookieContainer = cookieContainer ?? new CookieContainer(),
                    UseCookies = true,
                    SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                };

                this.httpClient = handler != null
                    ? new HttpClient(handlerToUse, disposeHandler: false)
                    : new HttpClient(handlerToUse);

                this.httpClient.Timeout = TimeSpan.FromSeconds(15);
            }
        }

        public static CurrentBusinessSystemConnector ForTests(string baseUrl, HttpMessageHandler handler)
        {
            if (handler == null)
            {
                throw new ArgumentNullException(nameof(handler));
            }

            return new CurrentBusinessSystemConnector(
                () => new AppSettings { BusinessBaseUrl = baseUrl },
                new CurrentBusinessSchemaMapper(PropertyLabels),
                new CurrentBusinessFieldMappingSeedBuilder(PropertyLabels),
                httpClient: null,
                handler: handler,
                cookieContainer: null);
        }

        public string SystemKey => CurrentSystemKey;

        public IReadOnlyList<ProjectOption> GetProjects()
        {
            var projects = Get<List<ProjectOption>>("/projects") ?? new List<ProjectOption>();
            return projects
                .Where(project => project != null && !string.IsNullOrWhiteSpace(project.ProjectId))
                .Select(project => new ProjectOption
                {
                    SystemKey = CurrentSystemKey,
                    ProjectId = project.ProjectId ?? string.Empty,
                    DisplayName = project.DisplayName ?? string.Empty,
                })
                .ToArray();
        }

        public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
        {
            if (project == null)
            {
                throw new ArgumentNullException(nameof(project));
            }

            return new SheetBinding
            {
                SheetName = sheetName ?? string.Empty,
                SystemKey = string.IsNullOrWhiteSpace(project.SystemKey) ? CurrentSystemKey : project.SystemKey,
                ProjectId = project.ProjectId ?? string.Empty,
                ProjectName = project.DisplayName ?? string.Empty,
                HeaderStartRow = DefaultHeaderStartRow,
                HeaderRowCount = DefaultHeaderRowCount,
                DataStartRow = DefaultDataStartRow,
            };
        }

        public FieldMappingTableDefinition GetFieldMappingDefinition(string projectId)
        {
            EnsureProjectId(projectId);

            return new FieldMappingTableDefinition
            {
                SystemKey = CurrentSystemKey,
                Columns = new[]
                {
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.HeaderType, Role = FieldMappingSemanticRole.HeaderType },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.IsdpLevel1, Role = FieldMappingSemanticRole.DefaultSingleHeaderText, RoleKey = CurrentBusinessFieldMappingColumns.DefaultLevel1 },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.IsdpLevel1, Role = FieldMappingSemanticRole.DefaultParentHeaderText, RoleKey = CurrentBusinessFieldMappingColumns.DefaultLevel1 },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.IsdpLevel2, Role = FieldMappingSemanticRole.DefaultChildHeaderText, RoleKey = CurrentBusinessFieldMappingColumns.DefaultLevel2 },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.ExcelLevel1, Role = FieldMappingSemanticRole.CurrentSingleHeaderText, RoleKey = CurrentBusinessFieldMappingColumns.CurrentLevel1 },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.ExcelLevel1, Role = FieldMappingSemanticRole.CurrentParentHeaderText, RoleKey = CurrentBusinessFieldMappingColumns.CurrentLevel1 },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.ExcelLevel2, Role = FieldMappingSemanticRole.CurrentChildHeaderText, RoleKey = CurrentBusinessFieldMappingColumns.CurrentLevel2 },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.HeaderId, Role = FieldMappingSemanticRole.HeaderIdentity },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.ApiFieldKey, Role = FieldMappingSemanticRole.ApiFieldKey },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.IsIdColumn, Role = FieldMappingSemanticRole.IsIdColumn },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.ActivityId, Role = FieldMappingSemanticRole.ActivityIdentity },
                    new FieldMappingColumnDefinition { ColumnName = CurrentBusinessFieldMappingHeaders.PropertyId, Role = FieldMappingSemanticRole.PropertyIdentity },
                },
            };
        }

        public IReadOnlyList<SheetFieldMappingRow> BuildFieldMappingSeed(string sheetName, string projectId)
        {
            EnsureProjectId(projectId);

            var headWrapper = Post<SchemaHeadWrapper>("/head", new { projectId });
            var headList = headWrapper?.HeadList ?? Array.Empty<CurrentBusinessHeadDefinition>();
            var sampleRows = Find(projectId, Array.Empty<string>(), Array.Empty<string>());

            return fieldMappingSeedBuilder.Build(sheetName, headList, sampleRows);
        }

        public WorksheetSchema GetSchema(string projectId)
        {
            var headWrapper = Post<SchemaHeadWrapper>("/head", new { projectId });
            var headList = headWrapper?.HeadList ?? Array.Empty<CurrentBusinessHeadDefinition>();
            var rows = Post<List<Dictionary<string, object>>>("/find", new
            {
                projectId,
                ids = Array.Empty<string>(),
                fieldKeys = Array.Empty<string>(),
            }) ?? new List<Dictionary<string, object>>();

            return schemaMapper.Build(projectId, headList, rows);
        }

        public IReadOnlyList<IDictionary<string, object>> Find(string projectId, IReadOnlyList<string> rowIds, IReadOnlyList<string> fieldKeys)
        {
            var requestedRowIds = rowIds ?? Array.Empty<string>();
            var payload = new
            {
                projectId,
                ids = requestedRowIds,
                rowIds = requestedRowIds,
                fieldKeys = fieldKeys ?? Array.Empty<string>(),
            };

            return Post<List<Dictionary<string, object>>>("/find", payload) ?? new List<Dictionary<string, object>>();
        }

        public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
        {
            if (changes == null)
            {
                throw new ArgumentNullException(nameof(changes));
            }

            if (changes.Count == 0)
            {
                return;
            }

            var items = changes.Select(change => new CurrentBusinessBatchSaveItem
            {
                ProjectId = projectId,
                Id = change.RowId,
                FieldKey = change.ApiFieldKey,
                Value = change.NewValue,
            }).ToArray();

            PostBatchSave(items);
        }

        private T Post<T>(string path, object payload)
        {
            using var response = SendPost(path, payload);
            var content = ReadResponseContent(response);
            EnsureSuccessStatusCode(response, content);
            if (string.IsNullOrWhiteSpace(content))
            {
                return default;
            }

            return JsonConvert.DeserializeObject<T>(content);
        }

        private T Get<T>(string path)
        {
            using var response = Send(HttpMethod.Get, path, payload: null);
            var content = ReadResponseContent(response);
            EnsureSuccessStatusCode(response, content);
            if (string.IsNullOrWhiteSpace(content))
            {
                return default;
            }

            return JsonConvert.DeserializeObject<T>(content);
        }

        private void PostBatchSave(CurrentBusinessBatchSaveItem[] items)
        {
            using var response = SendPost("/batchSave", items);
            var responseBody = ReadResponseContent(response);
            if (response.IsSuccessStatusCode)
            {
                return;
            }

            if (ShouldRetryLegacyBatchSave(response.StatusCode, responseBody))
            {
                OfficeAgentLog.Warn("business_api", "batch_save.legacy_retry", "Retrying batchSave with legacy items wrapper.", responseBody);
                using var legacyResponse = SendPost("/batchSave", new { items });
                var legacyResponseBody = ReadResponseContent(legacyResponse);
                EnsureSuccessStatusCode(legacyResponse, legacyResponseBody);
                return;
            }

            EnsureSuccessStatusCode(response, responseBody);
        }

        private HttpResponseMessage SendPost(string path, object payload)
        {
            return Send(HttpMethod.Post, path, payload);
        }

        private HttpResponseMessage Send(HttpMethod method, string path, object payload)
        {
            var baseUri = ResolveBaseUri();
            using var request = new HttpRequestMessage(method, new Uri(baseUri, path));
            var projectId = ExtractProjectId(payload);
            if (payload != null)
            {
                request.Content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
            }

            OfficeAgentLog.Info(
                "business_api",
                "request.begin",
                "Business API request started.",
                BuildRequestDetails(method, path, projectId));
            try
            {
                var response = httpClient.SendAsync(request).GetAwaiter().GetResult();
                OfficeAgentLog.Info(
                    "business_api",
                    "request.completed",
                    $"Business API request completed with {(int)response.StatusCode} {response.ReasonPhrase}.",
                    BuildRequestDetails(method, path, projectId));
                return response;
            }
            catch (OperationCanceledException ex)
            {
                OfficeAgentLog.Error(
                    "business_api",
                    "request.timeout",
                    "Business API request timed out or was canceled.",
                    ex,
                    BuildRequestDetails(method, path, projectId));
                throw;
            }
            catch (HttpRequestException ex)
            {
                OfficeAgentLog.Error(
                    "business_api",
                    "request.exception",
                    "Business API request failed with an HTTP transport error.",
                    ex,
                    BuildRequestDetails(method, path, projectId));
                throw;
            }
        }

        private static bool ShouldRetryLegacyBatchSave(HttpStatusCode statusCode, string responseBody)
        {
            return statusCode == HttpStatusCode.BadRequest
                && responseBody.IndexOf("items", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string ReadResponseContent(HttpResponseMessage response)
        {
            return response.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
        }

        private static void EnsureSuccessStatusCode(HttpResponseMessage response, string responseBody)
        {
            if (response.IsSuccessStatusCode)
            {
                return;
            }

            if (response.StatusCode == HttpStatusCode.Unauthorized ||
                response.StatusCode == HttpStatusCode.Forbidden)
            {
                throw new AuthenticationRequiredException("当前未登录，请先登录");
            }

            response.EnsureSuccessStatusCode();
        }

        private Uri ResolveBaseUri()
        {
            var settings = loadSettings() ?? new AppSettings();
            var normalizedBaseUrl = AppSettings.NormalizeOptionalUrl(settings.BusinessBaseUrl);
            if (!Uri.TryCreate(normalizedBaseUrl, UriKind.Absolute, out var baseUri))
            {
                throw new InvalidOperationException("The configured Business API Base URL is invalid. Update settings and try again.");
            }

            return baseUri;
        }

        private static void EnsureProjectId(string projectId)
        {
            if (string.IsNullOrWhiteSpace(projectId))
            {
                throw new InvalidOperationException("Project id is required for current business system.");
            }
        }

        private static string BuildRequestDetails(HttpMethod method, string path, string projectId)
        {
            var builder = new StringBuilder();
            AppendDetail(builder, "Method", method?.Method);
            AppendDetail(builder, "Path", path);
            AppendDetail(builder, "ProjectId", projectId);
            return builder.ToString();
        }

        private static string ExtractProjectId(object payload)
        {
            if (payload == null)
            {
                return string.Empty;
            }

            var property = payload.GetType().GetProperty("projectId") ??
                payload.GetType().GetProperty("ProjectId");
            return property?.GetValue(payload)?.ToString() ?? string.Empty;
        }

        private static void AppendDetail(StringBuilder builder, string name, string value)
        {
            if (builder.Length > 0)
            {
                builder.Append("; ");
            }

            builder
                .Append(name)
                .Append('=')
                .Append(string.IsNullOrWhiteSpace(value) ? "<empty>" : value);
        }
    }
}
