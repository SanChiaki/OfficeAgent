using System;
using Newtonsoft.Json;

namespace OfficeAgent.Core.Models
{
    public static class AiColumnMappingPreviewStatuses
    {
        public const string Accepted = "accepted";
        public const string LowConfidence = "lowConfidence";
        public const string Unmatched = "unmatched";
        public const string Rejected = "rejected";
    }

    public static class AiColumnMappingPreviewStatus
    {
        public const string Accepted = AiColumnMappingPreviewStatuses.Accepted;
        public const string LowConfidence = AiColumnMappingPreviewStatuses.LowConfidence;
        public const string Unmatched = AiColumnMappingPreviewStatuses.Unmatched;
        public const string Rejected = AiColumnMappingPreviewStatuses.Rejected;
    }

    public sealed class AiColumnMappingActualHeader
    {
        public int ExcelColumn { get; set; }

        [JsonIgnore]
        public int ExcelColumnIndex
        {
            get { return ExcelColumn; }
            set { ExcelColumn = value; }
        }

        public string DisplayText { get; set; } = string.Empty;

        public string ActualL1 { get; set; } = string.Empty;

        public string ActualL2 { get; set; } = string.Empty;
    }

    public sealed class AiColumnMappingCandidate
    {
        public string HeaderId { get; set; } = string.Empty;

        public string ApiFieldKey { get; set; } = string.Empty;

        public string HeaderType { get; set; } = string.Empty;

        public string IsdpL1 { get; set; } = string.Empty;

        [JsonIgnore]
        public string DefaultL1
        {
            get { return IsdpL1; }
            set { IsdpL1 = value ?? string.Empty; }
        }

        public string IsdpL2 { get; set; } = string.Empty;

        [JsonIgnore]
        public string DefaultL2
        {
            get { return IsdpL2; }
            set { IsdpL2 = value ?? string.Empty; }
        }

        public string CurrentExcelL1 { get; set; } = string.Empty;

        public string CurrentExcelL2 { get; set; } = string.Empty;

        public bool IsIdColumn { get; set; }

        public string ActivityId { get; set; } = string.Empty;

        public string PropertyId { get; set; } = string.Empty;
    }

    public sealed class AiColumnMappingRequest
    {
        public string SystemKey { get; set; } = string.Empty;

        public string SheetName { get; set; } = string.Empty;

        public AiColumnMappingActualHeader[] ActualHeaders { get; set; } = Array.Empty<AiColumnMappingActualHeader>();

        public AiColumnMappingCandidate[] Candidates { get; set; } = Array.Empty<AiColumnMappingCandidate>();
    }

    public sealed class AiColumnMappingSuggestion
    {
        public int ExcelColumn { get; set; }

        [JsonIgnore]
        public int ExcelColumnIndex
        {
            get { return ExcelColumn; }
            set { ExcelColumn = value; }
        }

        public string ActualL1 { get; set; } = string.Empty;

        public string ActualL2 { get; set; } = string.Empty;

        public string TargetHeaderId { get; set; } = string.Empty;

        [JsonIgnore]
        public string HeaderId
        {
            get { return TargetHeaderId; }
            set { TargetHeaderId = value; }
        }

        public string TargetApiFieldKey { get; set; } = string.Empty;

        [JsonIgnore]
        public string ApiFieldKey
        {
            get { return TargetApiFieldKey; }
            set { TargetApiFieldKey = value; }
        }

        public double Confidence { get; set; }

        public string Reason { get; set; } = string.Empty;
    }

    public sealed class AiColumnMappingUnmatchedHeader
    {
        public int ExcelColumn { get; set; }

        [JsonIgnore]
        public int ExcelColumnIndex
        {
            get { return ExcelColumn; }
            set { ExcelColumn = value; }
        }

        public string DisplayText { get; set; } = string.Empty;

        public string ActualL1 { get; set; } = string.Empty;

        public string ActualL2 { get; set; } = string.Empty;

        public string Reason { get; set; } = string.Empty;
    }

    public sealed class AiColumnMappingResponse
    {
        public AiColumnMappingSuggestion[] Mappings { get; set; } = Array.Empty<AiColumnMappingSuggestion>();

        [JsonIgnore]
        public AiColumnMappingSuggestion[] Suggestions
        {
            get { return Mappings; }
            set { Mappings = value ?? Array.Empty<AiColumnMappingSuggestion>(); }
        }

        public AiColumnMappingUnmatchedHeader[] Unmatched { get; set; } = Array.Empty<AiColumnMappingUnmatchedHeader>();

        [JsonIgnore]
        public AiColumnMappingUnmatchedHeader[] UnmatchedHeaders
        {
            get { return Unmatched; }
            set { Unmatched = value ?? Array.Empty<AiColumnMappingUnmatchedHeader>(); }
        }
    }

    public sealed class AiColumnMappingPreview
    {
        public AiColumnMappingPreviewItem[] Items { get; set; } = Array.Empty<AiColumnMappingPreviewItem>();
    }

    public sealed class AiColumnMappingPreviewItem
    {
        public int ExcelColumn { get; set; }

        [JsonIgnore]
        public int ExcelColumnIndex
        {
            get { return ExcelColumn; }
            set { ExcelColumn = value; }
        }

        public string SuggestedExcelL1 { get; set; } = string.Empty;

        [JsonIgnore]
        public string ActualL1
        {
            get { return SuggestedExcelL1; }
            set { SuggestedExcelL1 = value ?? string.Empty; }
        }

        public string SuggestedExcelL2 { get; set; } = string.Empty;

        [JsonIgnore]
        public string ActualL2
        {
            get { return SuggestedExcelL2; }
            set { SuggestedExcelL2 = value ?? string.Empty; }
        }

        public string TargetHeaderId { get; set; } = string.Empty;

        [JsonIgnore]
        public string HeaderId
        {
            get { return TargetHeaderId; }
            set { TargetHeaderId = value ?? string.Empty; }
        }

        public string TargetApiFieldKey { get; set; } = string.Empty;

        [JsonIgnore]
        public string ApiFieldKey
        {
            get { return TargetApiFieldKey; }
            set { TargetApiFieldKey = value ?? string.Empty; }
        }

        public string HeaderType { get; set; } = string.Empty;

        public string TargetIsdpL1 { get; set; } = string.Empty;

        [JsonIgnore]
        public string DefaultL1
        {
            get { return TargetIsdpL1; }
            set { TargetIsdpL1 = value ?? string.Empty; }
        }

        public string TargetIsdpL2 { get; set; } = string.Empty;

        [JsonIgnore]
        public string DefaultL2
        {
            get { return TargetIsdpL2; }
            set { TargetIsdpL2 = value ?? string.Empty; }
        }

        public double Confidence { get; set; }

        public string Status { get; set; } = AiColumnMappingPreviewStatuses.Unmatched;

        public bool ShouldApply { get; set; } = true;

        public string Reason { get; set; } = string.Empty;
    }

    public sealed class AiColumnMappingApplyResult
    {
        public SheetFieldMappingRow[] Rows { get; set; } = Array.Empty<SheetFieldMappingRow>();

        public int AppliedCount { get; set; }

        public int SkippedCount { get; set; }
    }
}
