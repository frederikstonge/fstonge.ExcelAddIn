using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sobeys.ExcelAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private Dictionary<string, WorkbookWrapper> _workbooks;

        public Ribbon()
        {
            _workbooks = new Dictionary<string, WorkbookWrapper>();
        }

        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            if (!_workbooks.ContainsKey(workbook.FullName))
            {
                _workbooks.Add(workbook.FullName, new WorkbookWrapper(workbook, _ribbon));
            }

            _ribbon.Invalidate();
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            if (!cancel)
            {
                var wrapper = _workbooks[workbook.FullName];
                _workbooks.Remove(workbook.FullName);
                wrapper.Dispose();
            }
        }

        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            _workbooks.Add(workbook.FullName, new WorkbookWrapper(workbook, _ribbon));
        }

        private WorkbookWrapper GetActiveWorkbookWrapper()
        {
            if (Globals.ThisAddIn.Application.ActiveWorkbook != null)
            {
                return _workbooks[Globals.ThisAddIn.Application.ActiveWorkbook.FullName];
            }

            return null;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Sobeys.ExcelAddIn.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public bool SuperCopyEnabled(Office.IRibbonControl control)
        {
            var activeWorkbook = GetActiveWorkbookWrapper();
            if (activeWorkbook != null)
            {
                return activeWorkbook.SuperCopyEnabled();
            }

            return false;
        }

        public void OnSuperCopy(Office.IRibbonControl control)
        {
            var activeWorkbook = GetActiveWorkbookWrapper();
            if (activeWorkbook != null)
            {
                activeWorkbook.OnSuperCopy();
            }
        }

        public void OnAbout(Office.IRibbonControl control)
        {
            System.Diagnostics.Process.Start("https://github.com/frederikstonge/sobeys-excel-addin");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this._ribbon = ribbonUI;
            Globals.ThisAddIn.Application.WorkbookOpen += Application_WorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            Globals.ThisAddIn.Application.WorkbookActivate += Application_WorkbookActivate;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
