using System.ComponentModel.Composition.Hosting;
using Sobeys.ExcelAddIn.Services;

namespace Sobeys.ExcelAddIn.Models
{
    public class WorkbookContainer
    {
        public WorkbookContainer(CompositionContainer container, IWorkbookService workbookService)
        {
            Container = container;
            WorkbookService = workbookService;
        }

        public CompositionContainer Container { get; }

        public IWorkbookService WorkbookService { get; }
    }
}
