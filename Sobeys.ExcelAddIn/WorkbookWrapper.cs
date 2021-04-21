using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sobeys.ExcelAddIn
{
    public class WorkbookWrapper : IDisposable
    {
        private Excel.Workbook _workbook;
        private Office.IRibbonUI _ribbon;

        public WorkbookWrapper(Excel.Workbook workbook, Office.IRibbonUI ribbon)
        {
            _workbook = workbook;
            _ribbon = ribbon;
            _workbook.SheetSelectionChange += _workbook_SheetSelectionChange;
        }

        public string FullName => _workbook.FullName;

        public bool SuperCopyEnabled()
        {
            Excel.Range range = Globals.ThisAddIn.Application.Selection;
            return range.Columns.Count == 1 && range.Rows.Count > 1;
        }

        public void OnSuperCopy()
        {
            try
            {
                Excel.Range range = Globals.ThisAddIn.Application.Selection;
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

        public void Dispose()
        {
            _workbook.SheetSelectionChange -= _workbook_SheetSelectionChange;
        }

        private void _workbook_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            _ribbon.Invalidate();
        }

    }
}
