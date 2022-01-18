using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using Sobeys.ExcelAddIn.Models;
using Sobeys.ExcelAddIn.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sobeys.ExcelAddIn
{
    public class Bootstrapper : IBootstrapper, IDisposable
    {
        private readonly IRibbon _ribbon;
        private readonly CompositionContainer _container;
        private readonly Dictionary<string, WorkbookContainer> _workbookContainers;

        public Bootstrapper(IRibbon ribbon)
        {
            _ribbon = ribbon;
            var catalog = new AggregateCatalog();
            catalog.Catalogs.Add(new AssemblyCatalog(typeof(AddIn).Assembly));
            _container = new CompositionContainer(catalog);
            _workbookContainers = new Dictionary<string, WorkbookContainer>();

            var batch = new CompositionBatch();
            batch.AddExportedValue(_ribbon);
            batch.AddExportedValue(Globals.AddIn);
            batch.AddExportedValue<IBootstrapper>(this);
            _container.Compose(batch);
            _container.ComposeParts(_ribbon);

            AddInService = _container.GetExportedValue<IAddInService>();

            Globals.AddIn.Application.WorkbookOpen += ApplicationWorkbookOpen;
            Globals.AddIn.Application.WorkbookBeforeClose += ApplicationWorkbookBeforeClose;
            Globals.AddIn.Application.WorkbookActivate += ApplicationWorkbookActivate;
        }

        public IAddInService AddInService { get; }

        public IWorkbookService ActiveWorkbookService => 
            _workbookContainers.ContainsKey(Globals.AddIn.Application.ActiveWorkbook.FullName)
            ? _workbookContainers[Globals.AddIn.Application.ActiveWorkbook.FullName].WorkbookService
            : null;

        public void Dispose()
        {
            Globals.AddIn.Application.WorkbookOpen -= ApplicationWorkbookOpen;
            Globals.AddIn.Application.WorkbookBeforeClose -= ApplicationWorkbookBeforeClose;
            Globals.AddIn.Application.WorkbookActivate -= ApplicationWorkbookActivate;

            foreach (var workbookContainer in _workbookContainers)
            {
                RemoveWorkbook(workbookContainer.Key);
            }

            _container.Dispose();
        }

        private void ApplicationWorkbookActivate(Excel.Workbook workbook)
        {
            if (!_workbookContainers.ContainsKey(workbook.FullName))
            {
                AddWorkbook(workbook);
            }

            _ribbon.Invalidate();
        }

        private void ApplicationWorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            if (!cancel)
            {
                RemoveWorkbook(workbook.FullName);
            }
        }

        private void ApplicationWorkbookOpen(Excel.Workbook workbook)
        {
            AddWorkbook(workbook);
        }

        private void AddWorkbook(Excel.Workbook workbook)
        {
            var catalog = new AggregateCatalog();
            catalog.Catalogs.Add(_container.Catalog);
            var container = new CompositionContainer(catalog);

            // Add singletons
            var batch = new CompositionBatch();
            batch.AddExportedValue(_container.GetExportedValue<IRibbon>());
            batch.AddExportedValue(_container.GetExportedValue<AddIn>());
            batch.AddExportedValue(_container.GetExportedValue<IBootstrapper>());
            batch.AddExportedValue(_container.GetExportedValue<IAddInService>());
            batch.AddExportedValue(workbook);
            container.Compose(batch);

            var workBookService = container.GetExportedValue<IWorkbookService>();
            _workbookContainers.Add(workbook.FullName, new WorkbookContainer(container, workBookService));
        }

        private void RemoveWorkbook(string key)
        {
            var container = _workbookContainers[key];
            container.Container.Dispose();
            _workbookContainers.Remove(key);
        }
    }
}
