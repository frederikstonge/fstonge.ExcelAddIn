using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Windows.Forms;
using fstonge.ExcelAddIn.Controls;
using fstonge.ExcelAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;

namespace fstonge.ExcelAddIn.Services
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
                Properties.Resources.Settings_Label,
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
                    SuperCopy(GetUsedSelectionRange());
                    break;
                case RibbonButtons.Merge:
                    var files = OpenMergeFiles();
                    Merge(files);
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
                RibbonButtons.SuperCopy => SuperCopyEnabled(GetUsedSelectionRange()),
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

        private bool SuperCopyEnabled(Excel.Range range)
        {
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

        private void SuperCopy(Excel.Range range)
        {
            try
            {
                if (!SuperCopyEnabled(range))
                {
                    return;
                }

                var items = new List<string>();

                var cells = Properties.Settings.Default.SuperCopyMode == SuperCopyMode.Column
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
            catch
            {
                // ignored
            }
        }

        private List<string> OpenMergeFiles()
        {
            using var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Multiselect = true;
            openFileDialog.RestoreDirectory = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileNames.ToList();
            }

            return new List<string>();
        }

        private void Merge(List<string> filePaths)
        {
            if (!filePaths.Any())
            { 
                return;
            }

            var oXL = new Excel.Application
            {
                Visible = false
            };

            try
            {
                foreach (var file in filePaths)
                {
                    Excel.Workbook workbook = oXL.Workbooks.Open(file, ReadOnly: true);
                    foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                    {
                        Excel.Worksheet destinationWorksheet = _workbook.Worksheets[worksheet.Name];
                        if (destinationWorksheet != null)
                        {
                            Excel.Range usedRange = worksheet.UsedRange;
                            var skip = Properties.Settings.Default.MergeSkipCells;
                            usedRange = usedRange.Offset[skip].Resize[usedRange.Rows.Count - skip];
                            usedRange.Copy();
                            Excel.Range last = destinationWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                            Excel.Range destinationUsedRange = destinationWorksheet.Cells[last.Row + 1, 1];
                            destinationUsedRange.Select();
                            destinationWorksheet.Paste();
                        }
                    }

                    oXL.CutCopyMode = 0;
                    workbook.Close(false);
                }
            }
            catch
            {
            }
            finally
            {
                oXL.Quit();
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