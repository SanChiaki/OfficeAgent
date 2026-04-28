using System.Collections.Generic;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal interface IWorksheetChangeLogStore
    {
        void Append(IReadOnlyList<WorksheetChangeLogEntry> entries);
    }
}
