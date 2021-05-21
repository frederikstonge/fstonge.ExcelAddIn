using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;


namespace Sobeys.ExcelAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private ThisAddIn _addIn;

        public Ribbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public void Invalidate()
        {
            _ribbon.Invalidate();
        }


        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Sobeys.ExcelAddIn.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public bool GetWorkbookEnabled(Office.IRibbonControl control)
        {
            return _addIn.AddInWrapper.GetWorkbookEnabled(control);
        }

        public void OnWorkbookAction(Office.IRibbonControl control)
        {
            _addIn.AddInWrapper.OnWorkbookAction(control);
        }

        public void OnAction(Office.IRibbonControl control)
        {
            _addIn.AddInWrapper.OnAction(control);
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
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