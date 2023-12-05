using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;
using Ionic.Zip;
using System.Security.Cryptography;
using System.Threading;
using Microsoft.Office.Interop.Outlook;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using System.Net.Mail;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace Yizip
{
    public partial class ThisAddIn
    {
        public static string pasword;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {            
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        void Application_ItemSend(object Item, ref bool Cancel)
        {
            ThisRibbonCollection ribbonCollection = Globals.Ribbons[Globals.ThisAddIn.Application.ActiveInspector()];
            Outlook.MailItem mailItem = Item as Outlook.MailItem;
            Outlook.Recipients recipients = mailItem.Recipients;
            var attachments = mailItem.Attachments.Count;
            if (ribbonCollection.YizipRibbon.btnActive.Enabled == true )
            {
                if (attachments > 0)
                {
                    // open new outlook              
                    string emailAddress = string.Empty;
                    string subject = mailItem.Subject;

                    foreach (Outlook.Recipient recipient in recipients)
                    {
                        emailAddress += recipient.Address + ";";
                    }
                    Outlook.MailItem eMail = (Outlook.MailItem)
                    this.Application.CreateItem(Outlook.OlItemType.olMailItem);
                    eMail.Subject = subject + " password";
                    eMail.To = emailAddress;
                    eMail.Body = "Password: " + pasword;
                    eMail.DeferredDeliveryTime = DateTime.Now.AddMinutes(1);
                    eMail.Importance = Outlook.OlImportance.olImportanceLow;
                    ((Outlook._MailItem)eMail).Send();
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - don't modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
