using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IWorksheetSelectionReader
    {
        IReadOnlyList<SelectedVisibleCell> ReadVisibleSelection();
    }
}
