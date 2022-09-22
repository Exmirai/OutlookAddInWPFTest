using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutlookAddInWPFTest.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OutlookAddInWPFTest.Managers
{
    public static class AlertManager
    {
        public static Alert[] GetAlerts()
        {
            var list = new List<Alert>();
            try
            {
                if (GlobalContext.App.ActiveInspector() is Inspector insp && insp != null)
                {
                    if (GlobalContext.App.ActiveInspector().CurrentItem is MailItem miItem && miItem != null)
                    {
                        if (GlobalContext.App.ActiveInspector().IsWordMail())
                        {
                            if (GlobalContext.App.ActiveInspector().WordEditor is Word.Document wd2 && wd2 != null)
                            {
                                var bodyText = wd2.Content.Text;
                                int curPos = 0;
                                int foundPos = -1;
                                while ((foundPos = bodyText.IndexOf(' ', curPos)) != -1)
                                {
                                    var left = 0;
                                    var top = 0;
                                    var width = 0;
                                    var height = 0;

                                    var range = wd2.Range(curPos, foundPos);
                                    range.Document.Windows.Cast<Word.Window>().FirstOrDefault()
                                        ?.GetPoint(out left, out top, out width, out height, range);
                                    list.Add(new Alert()
                                    {
                                        rect = new System.Drawing.Rectangle(left, top, width, height),
                                    });
                                }

                                return list.ToArray();
                            }
                        }
                    }
                }

                return Array.Empty<Alert>();
            }
            catch
            {
                return Array.Empty<Alert>();
            }
        }
    }

    public class Alert
    {
        public System.Drawing.Rectangle rect { get; set; }
        public void ProcessClick()
        {

        }

        public void ProcessHover()
        {
            
        }
    }
}
