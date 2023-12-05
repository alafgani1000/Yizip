using Ionic.Zip;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Diagnostics;

namespace Yizip
{
    public partial class YizipRibbon
    {
        public string password;

        private void YizipRibbon_Load(object sender, RibbonUIEventArgs e)
        {
                // %LocalAppData%\Microsoft\Windows\INetCache\Content.Outlook
        }

        void Application_Ziping(Outlook.MailItem mailItem, string password)
        {
            string path = "C:\\yiattahcment\\";
                if (mailItem != null)
                {
                    var attachments = mailItem.Attachments;
                    string[] fileNew = new string[10];
                    int _key = 0;

                    // delete directory
                    if (Directory.Exists(path))
                    {
                        Directory.Delete(path, true);
                    }
                    // save attachment to directory
                    foreach (Outlook.Attachment attachment in attachments)
                    {
                        string file = SaveAttachmentToDirectory(attachment, path);
                        fileNew[_key] = file;
                        _key++;
                    }

                    // delete attachment
                    for (int i = attachments.Count; i >= 1; i--)
                    {
                        DeleteAttacment(i, mailItem);
                    }

                    // create zip file         
                    CreateZipFile(path, path+ "Data.zip", password);

                    // attach file
                    AttachFile(path + "Data.zip", mailItem);
                }
            
        }

        private string SaveAttachmentToDirectory(Outlook.Attachment attachment, string path)
        {
            // Confirm that the attachment is a text file.            
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filePath = path + "\\" + attachment.FileName;
            attachment.SaveAsFile(filePath);
            return filePath;
        }

        public void CreateZipFile(string path, string outputPath, string password)
        {
            if (path != null)
            {
                using (ZipFile zip = new ZipFile())
                {
                    zip.Password = password;
                    zip.AddDirectory(path);
                    zip.Save(outputPath);
                }
            }
        }

        private void DeleteAttacment(int index, Outlook.MailItem mailItem)
        {
            mailItem.Attachments.Remove(index);
        }

        private void AttachFile(string fileName, Outlook.MailItem mailItem)
        {
            if (fileName != null)
            {
                mailItem.Attachments.Add(
                    fileName,
                    Outlook.OlAttachmentType.olByValue,
                    1,
                    fileName
                );
            }
        }

        public string GetRandomAlphanumericString(int length)
        {
            const string alphanumericCharacters =
                "ABCDEFGHIJKLMNOPQRSTUVWXYZ" +
                "abcdefghijklmnopqrstuvwxyz" +
                "0123456789";
            return GetRandomString(length, alphanumericCharacters);
        }

        public string GetRandomString(int length, IEnumerable<char> characterSet)
        {
            if (length < 0)
                throw new ArgumentException("length must not be negative", "length");
            if (length > int.MaxValue / 8)
                throw new ArgumentException("length is too big", "length");
            if (characterSet == null)
                throw new ArgumentNullException("characterSet");
            var characterArray = characterSet.Distinct().ToArray();
            if (characterArray.Length == 0)
                throw new ArgumentException("characterSet must not be empty", "characterSet");

            var bytes = new byte[length * 8];
            new RNGCryptoServiceProvider().GetBytes(bytes);
            var result = new char[length];
            for (int i = 0; i < length; i++)
            {
                ulong value = BitConverter.ToUInt64(bytes, i * 8);
                result[i] = characterArray[value % (uint)characterArray.Length];
            }
            return new string(result);
        }

        private void Button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            var m = e.Control.Context as Inspector;
            var mailItem = m.CurrentItem as MailItem;
            ThisRibbonCollection ribbonCollection = Globals.Ribbons[Globals.ThisAddIn.Application.ActiveInspector()];
            try
            {
                // generate password
                password = GetRandomAlphanumericString(8);
                ThisAddIn.pasword = password;
                ribbonCollection.YizipRibbon.btnActive.Enabled = true;
                Application_Ziping(mailItem, password);
            }
            catch (System.IO.DirectoryNotFoundException)
            {
                MessageBox.Show("Attachment not found", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ButtonActive_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.btnActive.Enabled.Equals(true))
            {
                this.btnActive.Enabled = false;
            }
        }

        private void ButtonCache_Click(object sender, RibbonControlEventArgs e)
        {
            string cmd = "explorer.exe";
            string arg = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Microsoft\\Windows\\INetCache\\Content.Outlook\\";
            Process.Start(cmd, arg);
        }
    }
}
