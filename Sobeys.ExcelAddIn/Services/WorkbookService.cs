using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Windows.Forms;
using Sobeys.ExcelAddIn.Controls;
using Sobeys.ExcelAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;

namespace Sobeys.ExcelAddIn.Services
{
    [Export(typeof(IWorkbookService))]
    [PartCreationPolicy(CreationPolicy.NonShared)]
    public class WorkbookService : IWorkbookService, IDisposable
    {
        private readonly Excel.Workbook _workbook;
        private readonly IRibbon _ribbon;
        private readonly Tools.CustomTaskPane _settingsTaskPane;

        [ImportingConstructor]
        public WorkbookService(Excel.Workbook workbook, IRibbon ribbon, ITaskPaneFactory taskPaneFactory)
        {
            _workbook = workbook;
            _ribbon = ribbon;
            _settingsTaskPane = taskPaneFactory.CreateTaskPane(
                new SettingsUserControl(), 
                "Settings",
                _workbook.Application.ActiveWindow, 
                Office.MsoCTPDockPosition.msoCTPDockPositionRight);

            _workbook.SheetSelectionChange += WorkbookSheetSelectionChange;
            _settingsTaskPane.VisibleChanged += SettingsTaskPaneVisibleChanged;
        }

        public void OnAction(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case RibbonButtons.SuperCopy:
                    SuperCopy();
                    break;
            }
        }

        public void OnPressedAction(Office.IRibbonControl control, bool isPressed)
        {
            switch (control.Id)
            {
                case RibbonButtons.Settings:
                    _settingsTaskPane.Visible = isPressed;
                    _ribbon.Invalidate();
                    break;
            }
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            return control.Id switch
            {
                RibbonButtons.SuperCopy => SuperCopyEnabled(),
                _ => true
            };
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            return control.Id switch
            {
                RibbonButtons.Settings => _settingsTaskPane.Visible,
                _ => true
            };
        }

        public void Dispose()
        {
            _workbook.SheetSelectionChange -= WorkbookSheetSelectionChange;
            _settingsTaskPane.VisibleChanged -= SettingsTaskPaneVisibleChanged;
            _settingsTaskPane.Dispose();
        }

        private bool SuperCopyEnabled()
        {
            Excel.Range range = GetUsedSelectionRange();

            if (range == null)
            {
                return false;
            }

            if (Properties.Settings.Default.SuperCopyMode == SuperCopyMode.Column)
            {
                return range.Columns.Count == 1 && range.Rows.Count > 1;
            }
            else
            {
                return range.Columns.Count > 1 && range.Rows.Count == 1;
            }
        }

        private void SuperCopy()
        {
            try
            {
                Excel.Range range = GetUsedSelectionRange();
                if (range == null)
                {
                    return;
                }

                var items = new List<string>();

                var cells = (SuperCopyMode)Properties.Settings.Default.SuperCopyMode == SuperCopyMode.Column
                    ? range.Rows.OfType<Excel.Range>().Skip(Properties.Settings.Default.SuperCopySkipCells)
                    : range.Columns.OfType<Excel.Range>().Skip(Properties.Settings.Default.SuperCopySkipCells);

                foreach (var cell in cells)
                {
                    var value = Convert.ToString(cell.Value2);
                    if (string.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    items.Add(value);
                }
               
                System.Windows.Clipboard.SetText(string.Join(Properties.Settings.Default.SuperCopyDelimiter, items));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "An error occurred");
            }
        }

        private void WorkbookSheetSelectionChange(object sheet, Excel.Range target)
        {
            _ribbon.Invalidate();
        }

        private void SettingsTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            _ribbon.Invalidate();
        }

        private Excel.Range GetUsedSelectionRange()
        {
            Excel.Range selection = _workbook.Application.Selection;
            Excel.Range usedRange = _workbook.ActiveSheet.UsedRange;

            return _workbook.Application.Intersect(selection, usedRange);
        }
    }
}