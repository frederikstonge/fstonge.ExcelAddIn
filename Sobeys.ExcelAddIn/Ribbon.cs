using System;
using System.ComponentModel.Composition;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace Sobeys.ExcelAddIn
{
    [ComVisible(true)]
    public class Ribbon : IRibbon, Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        [Import]
        public Bootstrapper Bootstrapper { get; set; }

        public void Invalidate()
        {
            _ribbon.Invalidate();
        }

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("Sobeys.ExcelAddIn.Ribbon.xml");
        }

        public bool GetWorkbookEnabled(Office.IRibbonControl control)
        {
            return Bootstrapper.ActiveWorkbookService?.GetEnabled(control) ?? false;
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            return Bootstrapper.AddInService?.GetEnabled(control) ?? false;
        }

        public void OnWorkbookAction(Office.IRibbonControl control)
        {
            Bootstrapper.ActiveWorkbookService?.OnAction(control);
        }

        public void OnAction(Office.IRibbonControl control)
        {
            Bootstrapper.AddInService?.OnAction(control);
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        private static string GetResourceText(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resource in resourceNames)
            {
                if (string.Equals(resourceName, resource, StringComparison.OrdinalIgnoreCase))
                {
                    using var stream = assembly.GetManifestResourceStream(resource);
                    if (stream != null)
                    {
                        using var resourceReader = new StreamReader(stream);
                        return resourceReader.ReadToEnd();
                    }
                }
            }

            return null;
        }
    }
}