namespace OfficeAgent.ExcelAddIn.Excel
{
    internal interface IWorksheetMetadataAdapter
    {
        string GetWorkbookScopeKey();

        void EnsureWorksheet(string name, bool visible);

        void WriteTable(string tableName, string[] headers, string[][] rows);

        string[] ReadHeaders(string tableName);

        string[][] ReadTable(string tableName);
    }
}
