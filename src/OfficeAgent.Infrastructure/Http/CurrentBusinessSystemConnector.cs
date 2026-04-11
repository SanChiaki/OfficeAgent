using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Authentication;
using System.Text;
using Newtonsoft.Json;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class CurrentBusinessSystemConnector : ISystemConnector
    {
        private static readonly IReadOnlyList<ProjectOption> Projects = new[]
        {
            new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "performance",
                DisplayName = "绩效项目",
            },
        };

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
        private readonly Func<AppSettings> loadSettings;
        private readonly HttpClient httpClient;

        public CurrentBusinessSystemConnector(Func<AppSettings> loadSettings, HttpClient httpClient = null, CookieContainer cookieContainer = null)
            : this(loadSettings ?? throw new ArgumentNullException(nameof(loadSettings)), new CurrentBusinessSchemaMapper(PropertyLabels), httpClient, handler: null, cookieContainer)
        {
        }

        private CurrentBusinessSystemConnector(
            Func<AppSettings> loadSettings,
            CurrentBusinessSchemaMapper schemaMapper,
            HttpClient httpClient,
            HttpMessageHandler handler,
            CookieContainer cookieContainer)
        {
            if (schemaMapper == null)
            {
                throw new ArgumentNullException(nameof(schemaMapper));
            }

            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));
            this.schemaMapper = schemaMapper;
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
                httpClient: null,
                handler: handler,
                cookieContainer: null);
        }

        public IReadOnlyList<ProjectOption> GetProjects() => Projects;

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
            var payload = new
            {
                projectId,
                ids = rowIds ?? Array.Empty<string>(),
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

            Post<object>("/batchSave", items);
        }

        private T Post<T>(string path, object payload)
        {
            var baseUri = ResolveBaseUri();
            using var request = new HttpRequestMessage(HttpMethod.Post, new Uri(baseUri, path))
            {
                Content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json"),
            };

            using var response = httpClient.SendAsync(request).GetAwaiter().GetResult();
            response.EnsureSuccessStatusCode();
            var content = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            if (string.IsNullOrWhiteSpace(content))
            {
                return default;
            }

            return JsonConvert.DeserializeObject<T>(content);
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
    }
}
