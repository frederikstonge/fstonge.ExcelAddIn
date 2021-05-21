using Sobeys.ExcelAddIn.Services;
using System.ComponentModel.Composition.Hosting;

namespace Sobeys.ExcelAddIn.Models
{
    public class WorkbookContainer
    {
        public WorkbookContainer(CompositionContainer container, WorkbookService workbookService)
        {
            Container = container;
            WorkbookService = workbookService;
        }

        public CompositionContainer Container { get; }

        public WorkbookService WorkbookService { get; }
    }
}
