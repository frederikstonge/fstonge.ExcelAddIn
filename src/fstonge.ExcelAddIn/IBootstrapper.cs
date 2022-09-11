using fstonge.ExcelAddIn.Services;

namespace fstonge.ExcelAddIn
{
    public interface IBootstrapper
    {
        IAddInService AddInService { get; }

        IWorkbookService ActiveWorkbookService { get; }
    }
}
