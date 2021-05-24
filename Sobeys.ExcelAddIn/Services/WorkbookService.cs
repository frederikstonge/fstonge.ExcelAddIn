using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
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
            _settingsTaskPane = taskPaneFactory.CreateTaskPane(new SettingsUserControl(), "Settings", _workbook.Application.ActiveWindow, Office.MsoCTPDockPosition.msoCTPDockPositionRight);
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
            Excel.Range range = _workbook.Application.Selection;
            return range.Columns.Count == 1 && range.Rows.Count > 1;
        }

        private void SuperCopy()
        {
            try
            {
                Excel.Range range = _workbook.Application.Selection;
                var items = new List<string>();
                foreach (Excel.Range row in range.Rows)
                {
                    var value = Convert.ToString(_workbook.ActiveSheet.Cells[row.Row, row.Column].Value2);
                    if (string.IsNullOrEmpty(value))
                    {
                        break;
                    }

                    items.Add(value);
                }

                System.Windows.Clipboard.SetText(string.Join(";", items));
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
    }
}
