using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IExcelContextService
    {
        SelectionContext GetCurrentSelectionContext();
    }
}
