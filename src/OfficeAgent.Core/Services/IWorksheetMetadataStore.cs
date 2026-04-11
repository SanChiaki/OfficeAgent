using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IWorksheetMetadataStore
    {
        void SaveBinding(SheetBinding binding);

        SheetBinding LoadBinding(string sheetName);

        WorksheetSnapshotCell[] LoadSnapshot(string sheetName);

        void SaveSnapshot(string sheetName, WorksheetSnapshotCell[] cells);
    }
}
