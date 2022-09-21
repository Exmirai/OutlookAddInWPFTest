using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OutlookAddInWPFTest.Utils;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using System.Globalization;
using Microsoft.Office.Interop.Outlook;
using OutlookAddInWPFTest.Managers;

namespace OutlookAddInWPFTest
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            StateManager.Init();
            using (var ctx = new DPIContextBlock(WinAPI.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE))
            {
                new Forms.JButton().Show();
                new Forms.Overlay().Show();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            GlobalContext.Init(GetHostItem<Application>(typeof(Application), "Application"));
            GlobalContext.Language = new CultureInfo(GlobalContext.App.LanguageSettings.LanguageID[MsoAppLanguageID.msoLanguageIDUI]);

            return new JRibbon();
        }
        #endregion
    }
}
