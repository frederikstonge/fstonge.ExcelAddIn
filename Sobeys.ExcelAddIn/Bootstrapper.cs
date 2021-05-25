using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using Microsoft.Office.Core;
using Sobeys.ExcelAddIn.Models;
using Sobeys.ExcelAddIn.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sobeys.ExcelAddIn
{
    public class Bootstrapper : IDisposable
    {
        private readonly IRibbon _ribbon;
        private readonly CompositionContainer _container;
        private readonly Dictionary<string, WorkbookContainer> _workbookContainers;

        public Bootstrapper(IRibbon ribbon)
        {
            _ribbon = ribbon;
            var catalog = new AggregateCatalog();
            catalog.Catalogs.Add(new AssemblyCatalog(typeof(Bootstrapper).Assembly));
            _container = new CompositionContainer(catalog);
            _workbookContainers = new Dictionary<string, WorkbookContainer>();

            var batch = new CompositionBatch();
            batch.AddExportedValue(_ribbon);
            batch.AddExportedValue(Globals.ThisAddIn);
            batch.AddExportedValue(this);
            _container.Compose(batch);
            _container.ComposeParts(_ribbon);

            AddInService = _container.GetExportedValue<IAddInService>();

            Globals.ThisAddIn.Application.WorkbookOpen += ApplicationWorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookBeforeClose += ApplicationWorkbookBeforeClose;
            Globals.ThisAddIn.Application.WorkbookActivate += ApplicationWorkbookActivate;
        }

        public IAddInService AddInService { get; }

        public IWorkbookService ActiveWorkbookService => _workbookContainers.ContainsKey(Globals.ThisAddIn.Application.ActiveWorkbook.FullName)
            ? _workbookContainers[Globals.ThisAddIn.Application.ActiveWorkbook.FullName].WorkbookService
            : null;

        public void Dispose()
        {
            Globals.ThisAddIn.Application.WorkbookOpen -= ApplicationWorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookBeforeClose -= ApplicationWorkbookBeforeClose;
            Globals.ThisAddIn.Application.WorkbookActivate -= ApplicationWorkbookActivate;

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
            batch.AddExportedValue(_container.GetExportedValue<ThisAddIn>());
            batch.AddExportedValue(_container.GetExportedValue<Bootstrapper>());
            batch.AddExportedValue(_container.GetExportedValue<IAddInService>());
            batch.AddExportedValue(workbook);
            container.Compose(batch);

            var workBookWrapper = container.GetExportedValue<IWorkbookService>();
            _workbookContainers.Add(workbook.FullName, new WorkbookContainer(container, workBookWrapper));
        }

        private void RemoveWorkbook(string key)
        {
            var container = _workbookContainers[key];
            container.Container.Dispose();
            _workbookContainers.Remove(key);
        }
    }
}
