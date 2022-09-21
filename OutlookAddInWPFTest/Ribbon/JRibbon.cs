using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

using Office = Microsoft.Office.Core;

namespace OutlookAddInWPFTest {
    [ComVisible(true)]
    public class JRibbon : Office.IRibbonExtensibility {
        private Office.IRibbonUI _ribbon;

        public JRibbon() {
        }


        public void Ribbon_Load(Office.IRibbonUI ribbonUi) {
            _ribbon = ribbonUi;
            GlobalContext.App.ItemSend += App_ItemSend;
        }



        private void App_ItemSend(object item, ref bool cancel) {
            try {
 
            }
            catch (Exception e) {

            }
            finally {
                Properties.Settings.Default.Save();
            }
        }



        #region Helpers

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("Outlook.Ribbon.JRibbon.xml");
        }

        private static string GetResourceText(string resourceName) {
            var res = Assembly
                            .GetExecutingAssembly()
                            .GetManifestResourceNames()
                            .Where(rn => rn.ToLower().CompareTo(resourceName.ToLower()) == 0)
                            .FirstOrDefault() ?? "";
            return
                string.IsNullOrEmpty(res)
                    ? ""
                    : new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream(res)).ReadToEnd();
        }

        #endregion
    }
}
