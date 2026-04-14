using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IWorksheetMetadataStore
    {
        void SaveBinding(SheetBinding binding);

        SheetBinding LoadBinding(string sheetName);

        void SaveFieldMappings(string sheetName, FieldMappingTableDefinition definition, IReadOnlyList<SheetFieldMappingRow> rows);

        SheetFieldMappingRow[] LoadFieldMappings(string sheetName, FieldMappingTableDefinition definition);

        void ClearFieldMappings(string sheetName);

        WorksheetSnapshotCell[] LoadSnapshot(string sheetName);

        void SaveSnapshot(string sheetName, WorksheetSnapshotCell[] cells);
    }
}
