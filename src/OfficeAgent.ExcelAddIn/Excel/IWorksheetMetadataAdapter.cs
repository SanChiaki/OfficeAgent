namespace OfficeAgent.ExcelAddIn.Excel
{
    internal interface IWorksheetMetadataAdapter
    {
        void EnsureWorksheet(string name, bool visible);

        void WriteTable(string tableName, string[] headers, string[][] rows);

        string[][] ReadTable(string tableName);
    }
}
