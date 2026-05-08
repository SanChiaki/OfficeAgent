using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class AiColumnMappingServiceTests
    {
        [Fact]
        public void BuildRequestUsesRoleKeysAndIncludesActualHeadersAndCandidates()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var actualHeaders = new[]
            {
                new AiColumnMappingActualHeader { ExcelColumn = 3, DisplayText = "项目负责人", ActualL1 = "项目负责人" },
            };

            var request = service.BuildRequest("Sheet1", definition, rows, actualHeaders);

            Assert.Equal("Sheet1", request.SheetName);
            Assert.Equal("current-business-system", request.SystemKey);
            Assert.Same(actualHeaders[0], Assert.Single(request.ActualHeaders));
            var candidate = Assert.Single(request.Candidates);
            Assert.Equal("owner_name", candidate.ApiFieldKey);
            Assert.Equal("owner_name", candidate.HeaderId);
            Assert.Equal("single", candidate.HeaderType);
            Assert.Equal("负责人", candidate.IsdpL1);
            Assert.Equal(string.Empty, candidate.IsdpL2);
            Assert.Equal("负责人", candidate.DefaultL1);
            Assert.Equal(string.Empty, candidate.DefaultL2);
            Assert.Equal("负责人", candidate.CurrentExcelL1);
            Assert.Equal(string.Empty, candidate.CurrentExcelL2);
        }

        [Fact]
        public void BuildRequestUsesOnlyCandidatesForRequestedSheet()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[]
            {
                CreateOwnerRow("Sheet1"),
                CreateOwnerRow("OtherSheet"),
            };

            var request = service.BuildRequest(
                "Sheet1",
                definition,
                rows,
                new[] { new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "项目负责人" } });

            var candidate = Assert.Single(request.Candidates);
            Assert.Equal("owner_name", candidate.HeaderId);
        }

        [Fact]
        public void JsonSerializationOmitsCompatibilityAliasProperties()
        {
            var response = new AiColumnMappingResponse
            {
                Suggestions = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumnIndex = 2,
                        HeaderId = "owner_name",
                        ApiFieldKey = "owner_name",
                        ActualL1 = "项目负责人",
                        Confidence = 0.91,
                    },
                },
                UnmatchedHeaders = new[]
                {
                    new AiColumnMappingUnmatchedHeader
                    {
                        ExcelColumnIndex = 5,
                        DisplayText = "备注",
                    },
                },
            };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumnIndex = 2,
                        ActualL1 = "项目负责人",
                        ActualL2 = string.Empty,
                        HeaderId = "owner_name",
                        ApiFieldKey = "owner_name",
                        DefaultL1 = "负责人",
                        Confidence = 0.91,
                    },
                },
            };

            var responseJson = JObject.FromObject(response);
            var suggestionJson = (JObject)responseJson["Mappings"][0];
            var unmatchedJson = (JObject)responseJson["Unmatched"][0];
            var previewJson = JObject.FromObject(preview);
            var previewItemJson = (JObject)previewJson["Items"][0];

            Assert.Null(responseJson["Suggestions"]);
            Assert.Null(responseJson["UnmatchedHeaders"]);
            Assert.Null(suggestionJson["ExcelColumnIndex"]);
            Assert.Null(suggestionJson["HeaderId"]);
            Assert.Null(suggestionJson["ApiFieldKey"]);
            Assert.Null(unmatchedJson["ExcelColumnIndex"]);
            Assert.Null(previewItemJson["ExcelColumnIndex"]);
            Assert.Null(previewItemJson["ActualL1"]);
            Assert.Null(previewItemJson["ActualL2"]);
            Assert.Null(previewItemJson["HeaderId"]);
            Assert.Null(previewItemJson["ApiFieldKey"]);
            Assert.Null(previewItemJson["DefaultL1"]);
            Assert.Null(previewItemJson["DefaultL2"]);
            Assert.Equal(2, (int)suggestionJson["ExcelColumn"]);
            Assert.Equal("owner_name", (string)suggestionJson["TargetHeaderId"]);
            Assert.Equal("owner_name", (string)suggestionJson["TargetApiFieldKey"]);
            Assert.Equal(2, (int)previewItemJson["ExcelColumn"]);
            Assert.Equal("项目负责人", (string)previewItemJson["SuggestedExcelL1"]);
        }

        [Fact]
        public async Task AiColumnMappingClientExposesSyncAndAsyncMapContracts()
        {
            var client = new FakeAiColumnMappingClient();
            var request = new AiColumnMappingRequest { SheetName = "Sheet1" };

            var syncResponse = client.Map(request);
            var asyncResponse = await client.MapAsync(request);

            Assert.Same(client.Response, syncResponse);
            Assert.Same(client.Response, asyncResponse);
            Assert.Same(request, client.LastRequest);
        }

        [Fact]
        public void ApplyConfirmedPreviewUpdatesOnlyCurrentHeaderRoles()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var request = CreateRequest(rows, new AiColumnMappingActualHeader { ExcelColumnIndex = 2, ActualL1 = "项目负责人" });
            var response = new AiColumnMappingResponse
            {
                Mappings = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumn = 2,
                        ActualL1 = "项目负责人",
                        TargetApiFieldKey = "owner_name",
                        TargetHeaderId = "owner_name",
                        Confidence = 0.92,
                    },
                },
            };
            var preview = service.CreatePreview(request, response, headerRowCount: 1);

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(1, result.AppliedCount);
            var item = Assert.Single(preview.Items);
            Assert.Equal(AiColumnMappingPreviewStatuses.Accepted, item.Status);
            Assert.Equal(2, item.ExcelColumn);
            Assert.Equal("负责人", item.TargetIsdpL1);
            Assert.Equal(string.Empty, item.TargetIsdpL2);
            Assert.Equal("项目负责人", item.SuggestedExcelL1);
            Assert.Equal(string.Empty, item.SuggestedExcelL2);
            Assert.NotSame(rows, result.Rows);
            Assert.NotSame(rows[0], result.Rows[0]);
            Assert.Equal("负责人", rows[0].Values["default_l1"]);
            Assert.Equal("负责人", result.Rows[0].Values["default_l1"]);
            Assert.Equal("项目负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("项目负责人", result.Rows[0].Values["current_parent"]);
            Assert.Equal(string.Empty, result.Rows[0].Values["current_child"]);
            Assert.Equal("owner_name", result.Rows[0].Values["header_id"]);
            Assert.Equal("owner_name", result.Rows[0].Values["api_field_key"]);
            Assert.Equal("false", result.Rows[0].Values["is_id_column"]);
            Assert.Equal(string.Empty, result.Rows[0].Values["activity_id"]);
            Assert.Equal(string.Empty, result.Rows[0].Values["property_id"]);
            Assert.Equal("负责人", rows[0].Values["current_single"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewUpdatesOnlyRowsForRequestedSheet()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var sheet1Row = CreateOwnerRow();
            var otherSheetRow = CreateOwnerRow("OtherSheet");
            var rows = new[] { sheet1Row, otherSheetRow };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.92,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(1, result.AppliedCount);
            Assert.Equal("Sheet1", result.Rows[0].SheetName);
            Assert.Equal("项目负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("OtherSheet", result.Rows[1].SheetName);
            Assert.Equal("负责人", result.Rows[1].Values["current_single"]);
            Assert.Equal("负责人", otherSheetRow.Values["current_single"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewRequiresSheetName()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.92,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var error = Assert.Throws<ArgumentException>(() =>
                service.ApplyConfirmedPreview(" ", definition, rows, preview, headerRowCount: 1));

            Assert.Equal("sheetName", error.ParamName);
            Assert.Contains("Sheet name is required.", error.Message, StringComparison.Ordinal);
            Assert.Equal("负责人", rows[0].Values["current_single"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewLeavesLowConfidenceAndUnmatchedRowsUnchanged()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var request = CreateRequest(rows, new AiColumnMappingActualHeader { ExcelColumnIndex = 2, ActualL1 = "项目负责人" });
            var response = new AiColumnMappingResponse
            {
                Mappings = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumnIndex = 2,
                        ActualL1 = "项目负责人",
                        ApiFieldKey = "owner_name",
                        HeaderId = "owner_name",
                        Confidence = 0.61,
                    },
                },
                Unmatched = new[]
                {
                    new AiColumnMappingUnmatchedHeader { ExcelColumn = 3, DisplayText = "备注", ActualL1 = "备注" },
                },
            };

            var preview = service.CreatePreview(request, response, headerRowCount: 1);
            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Contains(preview.Items, item => item.Status == AiColumnMappingPreviewStatuses.LowConfidence);
            Assert.Contains(preview.Items, item => item.Status == AiColumnMappingPreviewStatuses.Unmatched);
            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("负责人", result.Rows[0].Values["current_parent"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewRejectsTamperedAcceptedLowConfidenceItem()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        HeaderType = "single",
                        Confidence = 0.61,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("负责人", result.Rows[0].Values["current_parent"]);
            Assert.Equal(string.Empty, result.Rows[0].Values["current_child"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewRejectsTamperedAcceptedItemWithOnlyHeaderId()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        HeaderType = "single",
                        Confidence = 0.92,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("负责人", result.Rows[0].Values["current_parent"]);
            Assert.Equal(string.Empty, result.Rows[0].Values["current_child"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewRejectsTamperedAcceptedItemWithOnlyApiFieldKey()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "项目负责人",
                        TargetApiFieldKey = "owner_name",
                        HeaderType = "single",
                        Confidence = 0.92,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("负责人", result.Rows[0].Values["current_parent"]);
            Assert.Equal(string.Empty, result.Rows[0].Values["current_child"]);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        public void ApplyConfirmedPreviewRejectsTamperedAcceptedInvalidExcelColumns(int excelColumn)
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = excelColumn,
                        SuggestedExcelL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        HeaderType = "single",
                        Confidence = 0.92,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("负责人", result.Rows[0].Values["current_parent"]);
            Assert.Equal(string.Empty, result.Rows[0].Values["current_child"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewWritesL1AndL2WithoutUsingHeaderType()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateActivityPropertyRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 5,
                        SuggestedExcelL1 = "实际开始",
                        SuggestedExcelL2 = "实际结束",
                        TargetHeaderId = "start_12345678",
                        TargetApiFieldKey = "start_12345678",
                        HeaderType = "single",
                        Confidence = 0.93,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 2);

            Assert.Equal(1, result.AppliedCount);
            Assert.Equal("实际开始", result.Rows[0].Values["current_single"]);
            Assert.Equal("实际开始", result.Rows[0].Values["current_parent"]);
            Assert.Equal("实际结束", result.Rows[0].Values["current_child"]);
        }

        [Fact]
        public void ActivityPropertyMappingsCanUseSingleLevelModelRecommendationWhenHeaderRowCountIsOne()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateActivityPropertyRow() };
            var request = CreateRequest(rows, new AiColumnMappingActualHeader { ExcelColumn = 5, ActualL1 = "实际开始" });
            var response = new AiColumnMappingResponse
            {
                Mappings = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumn = 5,
                        ActualL1 = "实际开始",
                        TargetHeaderId = "start_12345678",
                        TargetApiFieldKey = "start_12345678",
                        Confidence = 0.93,
                    },
                },
            };

            var preview = service.CreatePreview(request, response, headerRowCount: 1);
            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(AiColumnMappingPreviewStatuses.Accepted, Assert.Single(preview.Items).Status);
            Assert.Equal(1, result.AppliedCount);
            Assert.Equal("实际开始", result.Rows[0].Values["current_single"]);
            Assert.Equal("实际开始", result.Rows[0].Values["current_parent"]);
            Assert.Equal(string.Empty, result.Rows[0].Values["current_child"]);
        }

        [Theory]
        [InlineData("", "开始时间")]
        [InlineData("测试活动111", "")]
        public void ApplyConfirmedPreviewAllowsPartialL1AndL2ForAnyHeaderType(string actualL1, string actualL2)
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateActivityPropertyRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 5,
                        SuggestedExcelL1 = actualL1,
                        SuggestedExcelL2 = actualL2,
                        TargetHeaderId = "start_12345678",
                        TargetApiFieldKey = "start_12345678",
                        HeaderType = "activityProperty",
                        Confidence = 0.93,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 2);

            Assert.Equal(1, result.AppliedCount);
            Assert.Equal(actualL1, result.Rows[0].Values["current_single"]);
            Assert.Equal(actualL1, result.Rows[0].Values["current_parent"]);
            Assert.Equal(actualL2, result.Rows[0].Values["current_child"]);
        }

        [Fact]
        public void CreatePreviewRejectsDuplicateTargetFields()
        {
            var service = new AiColumnMappingService();
            var rows = new[] { CreateOwnerRow() };
            var request = CreateRequest(
                rows,
                new AiColumnMappingActualHeader { ExcelColumnIndex = 2, ActualL1 = "项目负责人" },
                new AiColumnMappingActualHeader { ExcelColumnIndex = 3, ActualL1 = "负责人姓名" });
            var response = new AiColumnMappingResponse
            {
                Mappings = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumnIndex = 2,
                        ActualL1 = "项目负责人",
                        ApiFieldKey = "owner_name",
                        HeaderId = "owner_name",
                        Confidence = 0.92,
                    },
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumnIndex = 3,
                        ActualL1 = "负责人姓名",
                        ApiFieldKey = "owner_name",
                        HeaderId = "owner_name",
                        Confidence = 0.91,
                    },
                },
            };

            var preview = service.CreatePreview(request, response, headerRowCount: 1);
            var result = service.ApplyConfirmedPreview("Sheet1", CreateDefinition(), rows, preview, headerRowCount: 1);

            Assert.Equal(AiColumnMappingPreviewStatuses.Rejected, preview.Items[0].Status);
            Assert.Equal(AiColumnMappingPreviewStatuses.Rejected, preview.Items[1].Status);
            Assert.Contains("duplicate target", preview.Items[0].Reason, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("duplicate target", preview.Items[1].Reason, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
        }

        [Fact]
        public void CreatePreviewRejectsDuplicateExcelColumnSuggestions()
        {
            var service = new AiColumnMappingService();
            var rows = new[] { CreateOwnerRow(), CreateStatusRow() };
            var request = CreateRequest(rows, new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "项目负责人" });
            var response = new AiColumnMappingResponse
            {
                Mappings = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumn = 2,
                        ActualL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.92,
                    },
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumn = 2,
                        ActualL1 = "状态",
                        TargetHeaderId = "status",
                        TargetApiFieldKey = "status",
                        Confidence = 0.91,
                    },
                },
            };

            var preview = service.CreatePreview(request, response, headerRowCount: 1);
            var result = service.ApplyConfirmedPreview("Sheet1", CreateDefinition(), rows, preview, headerRowCount: 1);

            Assert.Equal(AiColumnMappingPreviewStatuses.Rejected, preview.Items[0].Status);
            Assert.Equal(AiColumnMappingPreviewStatuses.Rejected, preview.Items[1].Status);
            Assert.Contains("duplicate Excel column", preview.Items[0].Reason, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("duplicate Excel column", preview.Items[1].Reason, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("状态", result.Rows[1].Values["current_single"]);
        }

        [Fact]
        public void CreatePreviewRejectsDuplicateApiFieldTargetsWithDistinctHeaderIds()
        {
            var service = new AiColumnMappingService();
            var rows = new[]
            {
                CreateOwnerRowWithDistinctTargetIdentities(),
                CreateOwnerAliasRowWithSameApiFieldKey(),
            };
            var request = CreateRequest(
                rows,
                new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "项目负责人" },
                new AiColumnMappingActualHeader { ExcelColumn = 3, ActualL1 = "负责人姓名" });
            var response = new AiColumnMappingResponse
            {
                Mappings = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumn = 2,
                        TargetHeaderId = "owner_header",
                        Confidence = 0.92,
                    },
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumn = 3,
                        TargetHeaderId = "owner_alias_header",
                        Confidence = 0.91,
                    },
                },
            };

            var preview = service.CreatePreview(request, response, headerRowCount: 1);
            var result = service.ApplyConfirmedPreview("Sheet1", CreateDefinition(), rows, preview, headerRowCount: 1);

            Assert.Equal(AiColumnMappingPreviewStatuses.Rejected, preview.Items[0].Status);
            Assert.Equal(AiColumnMappingPreviewStatuses.Rejected, preview.Items[1].Status);
            Assert.Contains("duplicate target", preview.Items[0].Reason, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("duplicate target", preview.Items[1].Reason, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("负责人别名", result.Rows[1].Values["current_single"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewRejectsTamperedAcceptedDuplicateTargets()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.92,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 3,
                        SuggestedExcelL1 = "负责人姓名",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.91,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewRejectsTamperedAcceptedDuplicateExcelColumns()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow(), CreateStatusRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.92,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "状态",
                        TargetHeaderId = "status",
                        TargetApiFieldKey = "status",
                        Confidence = 0.91,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("状态", result.Rows[1].Values["current_single"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewRejectsOverlappingDuplicateConflictSets()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow(), CreateStatusRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 1,
                        SuggestedExcelL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.92,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "负责人姓名",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.91,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "状态",
                        TargetHeaderId = "status",
                        TargetApiFieldKey = "status",
                        Confidence = 0.9,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
            Assert.Equal("状态", result.Rows[1].Values["current_single"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewRejectsDuplicateTargetsBeforeEligibilityFiltering()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 1,
                        SuggestedExcelL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.92,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "负责人姓名",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.61,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewRejectsDuplicateApiFieldTargetsBeforeEligibilityFiltering()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRowWithDistinctTargetIdentities() };
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 1,
                        SuggestedExcelL1 = "项目负责人",
                        TargetHeaderId = "owner_header",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.92,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "负责人姓名",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.61,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
        }

        [Fact]
        public void CreatePreviewRejectsSuggestionsForColumnsMissingFromActualHeaders()
        {
            var service = new AiColumnMappingService();
            var rows = new[] { CreateOwnerRow() };
            var request = CreateRequest(rows, new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "项目负责人" });
            var response = new AiColumnMappingResponse
            {
                Mappings = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumn = 99,
                        ActualL1 = "AI 发明列",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.93,
                    },
                },
            };

            var preview = service.CreatePreview(request, response, headerRowCount: 1);
            var result = service.ApplyConfirmedPreview("Sheet1", CreateDefinition(), rows, preview, headerRowCount: 1);

            Assert.Equal(AiColumnMappingPreviewStatuses.Rejected, Assert.Single(preview.Items).Status);
            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
        }

        [Fact]
        public void CreatePreviewRejectsInvalidHeaderIdEvenWhenApiFieldKeyIsValid()
        {
            var service = new AiColumnMappingService();
            var rows = new[] { CreateOwnerRow() };
            var request = CreateRequest(rows, new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "项目负责人" });
            var response = new AiColumnMappingResponse
            {
                Mappings = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumn = 2,
                        ActualL1 = "项目负责人",
                        TargetHeaderId = "missing_header",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.92,
                    },
                },
            };

            var preview = service.CreatePreview(request, response, headerRowCount: 1);
            var result = service.ApplyConfirmedPreview("Sheet1", CreateDefinition(), rows, preview, headerRowCount: 1);

            Assert.Equal(AiColumnMappingPreviewStatuses.Rejected, Assert.Single(preview.Items).Status);
            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
        }

        [Fact]
        public void CreatePreviewRejectsInvalidApiFieldKeyEvenWhenHeaderIdIsValid()
        {
            var service = new AiColumnMappingService();
            var rows = new[] { CreateOwnerRow() };
            var request = CreateRequest(rows, new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "项目负责人" });
            var response = new AiColumnMappingResponse
            {
                Mappings = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumn = 2,
                        ActualL1 = "项目负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "missing_api_field",
                        Confidence = 0.92,
                    },
                },
            };

            var preview = service.CreatePreview(request, response, headerRowCount: 1);
            var result = service.ApplyConfirmedPreview("Sheet1", CreateDefinition(), rows, preview, headerRowCount: 1);

            Assert.Equal(AiColumnMappingPreviewStatuses.Rejected, Assert.Single(preview.Items).Status);
            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows[0].Values["current_single"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewAllowsL2SuggestionsWhenHeaderRowCountIsOne()
        {
            var service = new AiColumnMappingService();
            var definition = CreateDefinition();
            var rows = new[] { CreateOwnerRow() };
            var request = CreateRequest(rows, new AiColumnMappingActualHeader { ExcelColumnIndex = 2, ActualL1 = "项目", ActualL2 = "负责人" });
            var response = new AiColumnMappingResponse
            {
                Mappings = new[]
                {
                    new AiColumnMappingSuggestion
                    {
                        ExcelColumnIndex = 2,
                        ActualL1 = "项目",
                        ActualL2 = "负责人",
                        ApiFieldKey = "owner_name",
                        HeaderId = "owner_name",
                        Confidence = 0.92,
                    },
                },
            };

            var preview = service.CreatePreview(request, response, headerRowCount: 1);
            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(AiColumnMappingPreviewStatuses.Accepted, Assert.Single(preview.Items).Status);
            Assert.Equal(1, result.AppliedCount);
            Assert.Equal("项目", result.Rows[0].Values["current_single"]);
            Assert.Equal("项目", result.Rows[0].Values["current_parent"]);
            Assert.Equal("负责人", result.Rows[0].Values["current_child"]);
        }

        private static AiColumnMappingRequest CreateRequest(
            IReadOnlyList<SheetFieldMappingRow> rows,
            params AiColumnMappingActualHeader[] actualHeaders)
        {
            return new AiColumnMappingService().BuildRequest("Sheet1", CreateDefinition(), rows, actualHeaders);
        }

        private static FieldMappingTableDefinition CreateDefinition()
        {
            return new FieldMappingTableDefinition
            {
                SystemKey = "current-business-system",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition { ColumnName = "HeaderId", Role = FieldMappingSemanticRole.HeaderIdentity, RoleKey = "header_id" },
                    new FieldMappingColumnDefinition { ColumnName = "HeaderType", Role = FieldMappingSemanticRole.HeaderType, RoleKey = "header_type" },
                    new FieldMappingColumnDefinition { ColumnName = "ISDP L1", Role = FieldMappingSemanticRole.DefaultSingleHeaderText, RoleKey = "default_l1" },
                    new FieldMappingColumnDefinition { ColumnName = "ISDP Parent", Role = FieldMappingSemanticRole.DefaultParentHeaderText, RoleKey = "default_parent" },
                    new FieldMappingColumnDefinition { ColumnName = "ISDP L2", Role = FieldMappingSemanticRole.DefaultChildHeaderText, RoleKey = "default_l2" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L1", Role = FieldMappingSemanticRole.CurrentSingleHeaderText, RoleKey = "current_single" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel Parent", Role = FieldMappingSemanticRole.CurrentParentHeaderText, RoleKey = "current_parent" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L2", Role = FieldMappingSemanticRole.CurrentChildHeaderText, RoleKey = "current_child" },
                    new FieldMappingColumnDefinition { ColumnName = "ApiFieldKey", Role = FieldMappingSemanticRole.ApiFieldKey, RoleKey = "api_field_key" },
                    new FieldMappingColumnDefinition { ColumnName = "IsIdColumn", Role = FieldMappingSemanticRole.IsIdColumn, RoleKey = "is_id_column" },
                    new FieldMappingColumnDefinition { ColumnName = "ActivityId", Role = FieldMappingSemanticRole.ActivityIdentity, RoleKey = "activity_id" },
                    new FieldMappingColumnDefinition { ColumnName = "PropertyId", Role = FieldMappingSemanticRole.PropertyIdentity, RoleKey = "property_id" },
                },
            };
        }

        private sealed class FakeAiColumnMappingClient : IAiColumnMappingClient
        {
            public AiColumnMappingResponse Response { get; } = new AiColumnMappingResponse();

            public AiColumnMappingRequest LastRequest { get; private set; }

            public AiColumnMappingResponse Map(AiColumnMappingRequest request)
            {
                LastRequest = request;
                return Response;
            }

            public Task<AiColumnMappingResponse> MapAsync(AiColumnMappingRequest request)
            {
                LastRequest = request;
                return Task.FromResult(Response);
            }
        }

        private static SheetFieldMappingRow CreateOwnerRow(string sheetName = "Sheet1")
        {
            return new SheetFieldMappingRow
            {
                SheetName = sheetName,
                Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["header_id"] = "owner_name",
                    ["header_type"] = "single",
                    ["default_l1"] = "负责人",
                    ["default_parent"] = "负责人",
                    ["default_l2"] = string.Empty,
                    ["current_single"] = "负责人",
                    ["current_parent"] = "负责人",
                    ["current_child"] = string.Empty,
                    ["api_field_key"] = "owner_name",
                    ["is_id_column"] = "false",
                    ["activity_id"] = string.Empty,
                    ["property_id"] = string.Empty,
                },
            };
        }

        private static SheetFieldMappingRow CreateOwnerRowWithDistinctTargetIdentities()
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in CreateOwnerRow().Values)
            {
                values[pair.Key] = pair.Value;
            }

            values["header_id"] = "owner_header";
            values["api_field_key"] = "owner_name";

            return new SheetFieldMappingRow
            {
                SheetName = "Sheet1",
                Values = values,
            };
        }

        private static SheetFieldMappingRow CreateOwnerAliasRowWithSameApiFieldKey()
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in CreateOwnerRow().Values)
            {
                values[pair.Key] = pair.Value;
            }

            values["header_id"] = "owner_alias_header";
            values["default_l1"] = "负责人别名";
            values["default_parent"] = "负责人别名";
            values["current_single"] = "负责人别名";
            values["current_parent"] = "负责人别名";
            values["api_field_key"] = "owner_name";

            return new SheetFieldMappingRow
            {
                SheetName = "Sheet1",
                Values = values,
            };
        }

        private static SheetFieldMappingRow CreateStatusRow()
        {
            return new SheetFieldMappingRow
            {
                SheetName = "Sheet1",
                Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["header_id"] = "status",
                    ["header_type"] = "single",
                    ["default_l1"] = "状态",
                    ["default_parent"] = "状态",
                    ["default_l2"] = string.Empty,
                    ["current_single"] = "状态",
                    ["current_parent"] = "状态",
                    ["current_child"] = string.Empty,
                    ["api_field_key"] = "status",
                    ["is_id_column"] = "false",
                    ["activity_id"] = string.Empty,
                    ["property_id"] = string.Empty,
                },
            };
        }

        private static SheetFieldMappingRow CreateActivityPropertyRow()
        {
            return new SheetFieldMappingRow
            {
                SheetName = "Sheet1",
                Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["header_id"] = "start_12345678",
                    ["header_type"] = "activityProperty",
                    ["default_l1"] = string.Empty,
                    ["default_parent"] = "测试活动111",
                    ["default_l2"] = "开始时间",
                    ["current_single"] = string.Empty,
                    ["current_parent"] = "测试活动111",
                    ["current_child"] = "开始时间",
                    ["api_field_key"] = "start_12345678",
                    ["is_id_column"] = "false",
                    ["activity_id"] = "12345678",
                    ["property_id"] = "start",
                },
            };
        }
    }
}
