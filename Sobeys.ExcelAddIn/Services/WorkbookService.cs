using Sobeys.ExcelAddIn.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sobeys.ExcelAddIn.Services
{
    [Export]
    public class WorkbookService : IDisposable
    {
        private Excel.Workbook _workbook;
        private Ribbon _ribbon;

        [ImportingConstructor]
        public WorkbookService(Excel.Workbook workbook, Ribbon ribbon)
        {
            _workbook = workbook;
            _ribbon = ribbon;
            _workbook.SheetSelectionChange += WorkbookSheetSelectionChange;
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

        public bool GetEnabled(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case RibbonButtons.SuperCopy:
                    return SuperCopyEnabled();
                default:
                    return true;
            }
        }

        public void Dispose()
        {
            _workbook.SheetSelectionChange -= WorkbookSheetSelectionChange;
        }

        private bool SuperCopyEnabled()
        {
            Excel.Range range = Globals.ThisAddIn.Application.Selection;
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

        private void WorkbookSheetSelectionChange(object shheet, Excel.Range target)
        {
            _ribbon.Invalidate();
        }
    }
}
