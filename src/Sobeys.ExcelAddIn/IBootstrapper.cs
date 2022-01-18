using Sobeys.ExcelAddIn.Services;

namespace Sobeys.ExcelAddIn
{
    public interface IBootstrapper
    {
        IAddInService AddInService { get; }

        IWorkbookService ActiveWorkbookService { get; }
    }
}
