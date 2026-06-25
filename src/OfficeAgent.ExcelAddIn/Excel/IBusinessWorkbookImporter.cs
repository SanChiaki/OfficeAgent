namespace OfficeAgent.ExcelAddIn.Excel
{
    internal interface IBusinessWorkbookImporter
    {
        bool IsWorkSheetContentBlank(string sheetName);

        void EnsureCanWriteToWorkSheet(string sheetName);

        void ImportBusinessDataSheet(byte[] workbookBytes, string targetSheetName);

        void ActivateWorkSheetAtA1(string sheetName);
    }
}
