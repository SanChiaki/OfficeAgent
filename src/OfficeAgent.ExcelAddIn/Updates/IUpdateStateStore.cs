namespace OfficeAgent.ExcelAddIn.Updates
{
    internal interface IUpdateStateStore
    {
        UpdateState Load();
        void Save(UpdateState state);
    }
}
