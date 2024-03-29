﻿using System;
using System.ComponentModel.Composition;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace fstonge.ExcelAddIn
{
    [ComVisible(true)]
    public class Ribbon : IRibbon, Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        [Import]
        public IBootstrapper Bootstrapper { get; set; }

        public void Invalidate()
        {
            _ribbon.Invalidate();
        }

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("fstonge.ExcelAddIn.Ribbon.xml");
        }

        public bool GetWorkbookEnabled(Office.IRibbonControl control)
        {
            return Bootstrapper.ActiveWorkbookService?.GetEnabled(control) ?? false;
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            return Bootstrapper.AddInService?.GetEnabled(control) ?? false;
        }
        
        public string GetLabel(Office.IRibbonControl control)
        {
            return Properties.Resources.ResourceManager.GetString($"Ribbon_{control.Id}");
        }

        public void OnWorkbookAction(Office.IRibbonControl control)
        {
            Bootstrapper.ActiveWorkbookService?.OnAction(control);
        }

        public void OnWorkbookPressedAction(Office.IRibbonControl control, bool isPressed)
        {
            Bootstrapper.ActiveWorkbookService?.OnPressedAction(control, isPressed);
        }

        public void OnAction(Office.IRibbonControl control)
        {
            Bootstrapper.AddInService?.OnAction(control);
        }

        public bool GetWorkbookPressed(Office.IRibbonControl control)
        {
            return Bootstrapper.ActiveWorkbookService?.GetPressed(control) ?? false;
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