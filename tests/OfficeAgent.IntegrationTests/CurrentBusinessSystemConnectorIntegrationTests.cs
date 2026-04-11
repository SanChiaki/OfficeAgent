using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Security.Authentication;
using System.Threading.Tasks;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.IntegrationTests
{
    public sealed class CurrentBusinessSystemConnectorIntegrationTests : IClassFixture<MockServerFixture>
    {
        private readonly MockServerFixture fixture;

        public CurrentBusinessSystemConnectorIntegrationTests(MockServerFixture fixture)
        {
            this.fixture = fixture;
        }

        [Fact]
        public async Task FindAndBatchSaveRoundTripAgainstMockServer()
        {
            var connector = await CreateConnectorAsync();
            var rows = connector.Find("performance", new[] { "row-1" }, new[] { "name", "start_12345678" });
            Assert.Single(rows);

            connector.BatchSave("performance", new[]
            {
                new CellChange
                {
                    RowId = "row-1",
                    ApiFieldKey = "name",
                    NewValue = "Updated",
                },
            });

            var updatedRows = connector.Find("performance", new[] { "row-1" }, new[] { "name" });
            Assert.Single(updatedRows);
            Assert.Equal("Updated", updatedRows[0]["name"]?.ToString());
        }

        [Fact]
        public async Task GetSchemaIncludesActivityColumns()
        {
            var connector = await CreateConnectorAsync();
            var schema = connector.GetSchema("performance");

            Assert.NotNull(schema);
            var columns = schema.Columns;
            Assert.Contains(columns, column => column.ApiFieldKey == "name" && column.ColumnKind == WorksheetColumnKind.Single);
            Assert.Contains(columns, column => column.ApiFieldKey == "start_12345678" && column.ColumnKind == WorksheetColumnKind.ActivityProperty);
            Assert.Contains(columns, column => column.ApiFieldKey == "end_12345678" && column.ColumnKind == WorksheetColumnKind.ActivityProperty);
            var activityColumn = columns.First(column => column.ApiFieldKey == "start_12345678");
            Assert.Equal("测试活动111", activityColumn.ActivityName);
            Assert.Equal("开始时间", activityColumn.ChildHeaderText);
        }

        private async Task<CurrentBusinessSystemConnector> CreateConnectorAsync()
        {
            var cookieJar = await fixture.LoginAs("connector_user", "password123").ConfigureAwait(false);
            var httpClient = new HttpClient(new HttpClientHandler
            {
                UseCookies = true,
                CookieContainer = cookieJar,
                SslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13,
            })
            {
                Timeout = TimeSpan.FromSeconds(10),
            };

            return new CurrentBusinessSystemConnector(
                () => new AppSettings { BusinessBaseUrl = fixture.BusinessUrl, ApiKey = string.Empty },
                httpClient);
        }
    }
}
