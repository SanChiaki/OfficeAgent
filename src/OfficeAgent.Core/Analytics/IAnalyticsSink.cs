using System.Threading;
using System.Threading.Tasks;

namespace OfficeAgent.Core.Analytics
{
    public interface IAnalyticsSink
    {
        Task WriteAsync(AnalyticsEvent analyticsEvent, CancellationToken cancellationToken);
    }
}
