using Microsoft.Office.Interop.Outlook;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOtwieranieTest
{
    internal class OutlookEmailSender
    {
        public string AttachementPatch { get; set; } = @"C:\dane\Excel\Attachments\";
        public IProgress<int> Progress { get; set; }

        public OutlookEmailSender()
        {
            bool exists = System.IO.Directory.Exists(AttachementPatch);

            if (!exists)
                System.IO.Directory.CreateDirectory(AttachementPatch);
        }

        public void CreateEmail(string[][] tableData)
        {
            //Interop.Microsoft.Office.Interop.Outlook z nuget
            //C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL
            Application outlook = new Application();

            // Create a new MailItem object
            MailItem email = (MailItem)outlook.CreateItem(OlItemType.olMailItem);

            // Set the recipients, subject, and body of the email message
            email.To = "recipient@example.com";
            email.Subject = "Table from Jagged Array Example";
            email.BodyFormat = OlBodyFormat.olFormatHTML;

            // Generate a table from the jagged array data
            string tableHtml = "<table>";
            for (int i = 0; i < tableData.Length; i++)
            {
                tableHtml += "<tr>";
                for (int j = 0; j < tableData[i].Length; j++)
                {
                    tableHtml += $"<td>{tableData[i][j]}</td>";
                }
                tableHtml += "</tr>";
            }
            tableHtml += "</table>";

            // Insert the table into the email message body
            email.HTMLBody = tableHtml;

            // Display the email message
            email.Display(true);
        }

        public void CreateNewEmail(string to, string subject, string body, params MailItem[] oldMailItem)
        {
            Application outlook = new Application();

            // Create a new MailItem object
            MailItem newEmail = (MailItem)outlook.CreateItem(OlItemType.olMailItem);

            // Set the recipients, subject, and body of the email message
            newEmail.To = to;
            newEmail.Subject = subject;
            newEmail.BodyFormat = OlBodyFormat.olFormatHTML;

            // Insert the table into the email message body
            newEmail.HTMLBody = body;

            foreach (var oe in oldMailItem)
                foreach (Attachment attachment in oe.Attachments)
                {
                    string filePath = Path.Combine(AttachementPatch, attachment.FileName);
                    newEmail.Attachments.Add(filePath);
                }

            // Display the email message
            newEmail.Display(true);
        }

        public IEnumerable<MailItem> ReadEmail(string titleFilter)
        {
            // Initialize Outlook application object
            Application outlookApp = new Application();

            MAPIFolder oPublicFolder = (MAPIFolder)outlookApp.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolder‌​Inbox).Parent;
            // Get the inbox folder
            //MAPIFolder inbox = outlookApp.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            // Get all the email items in the inbox
            Items items = oPublicFolder.Items;

            // Filter the items to only include emails with "xxx" in thBook2e subject line
            string filter = $@"@SQL=""urn:schemas:mailheader:subject"" LIKE '%{titleFilter}%'";
            items = items.Restrict(filter);
            string emailContent = string.Empty;
            // Loop through each email item that matches the filter
            foreach (object item in items)
            {
                if (item is MailItem mailItem)
                {
                    // Read the email content to a string variable
                    emailContent = mailItem.Body;

                    // Loop through each attachment and save to disk
                    foreach (Attachment attachment in mailItem.Attachments)
                    {
                        string filePath = Path.Combine(AttachementPatch, attachment.FileName);
                        attachment.SaveAsFile(filePath);
                    }
                    yield return mailItem;
                }
            }
        }

        private List<MailItem> _foundEmails;
        private int progVal = 1;

        public IEnumerable<MailItem> ReadEmailRecur(string titleFilter)
        {
            progVal = 1;

            _foundEmails = new List<MailItem>();
            // Initialize Outlook application object
            Application outlookApp = new Application();

            Folder root = (Folder)outlookApp.Session.DefaultStore.GetRootFolder();
            EnumerateFolders(root, titleFilter);

            return _foundEmails;
        }

        private void EnumerateFolders(Folder folder, string titleFilter)
        {
            //Progress.Report(progVal++);
            Progress.Report((int)((double)progVal++ / Deep * 100)); // liczmy % progresu na podstawie aktualnego progVal oraz Deep (całowita liczba folderów)
            // Write the folder path.

            // MAPIFolder oPublicFolder = (MAPIFolder)outlookApp.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolder‌​Inbox).Parent;
            // Get the inbox folder
            //MAPIFolder inbox = outlookApp.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            // Get all the email items in the inbox
            Items items = folder.Items;

            // Filter the items to only include emails with "xxx" in thBook2e subject line
            string filter = $@"@SQL=""urn:schemas:mailheader:subject"" LIKE '%{titleFilter}%'";
            items = items.Restrict(filter);
            string emailContent = string.Empty;
            // Loop through each email item that matches the filter
            foreach (object item in items)
            {
                if (item is MailItem mailItem)
                {
                    // Read the email content to a string variable
                    emailContent = mailItem.Body;

                    // Loop through each attachment and save to disk
                    foreach (Attachment attachment in mailItem.Attachments)
                    {
                        string filePath = Path.Combine(AttachementPatch, attachment.FileName);
                        attachment.SaveAsFile(filePath);
                    }
                    _foundEmails.Add(mailItem);
                }
            }

            Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    // Call EnumerateFolders using childFolder.
                    EnumerateFolders(childFolder, titleFilter);
                }
            }
        }

        public int? Deep { get; set; } = null;

        public void CalculateDeep()
        {
            if (Deep != null)
                return;

            Deep = 1;
            // Initialize Outlook application object
            Application outlookApp = new Application();

            Folder root = (Folder)outlookApp.Session.DefaultStore.GetRootFolder();
            EnumerateFoldersDeep(root);
        }

        private void EnumerateFoldersDeep(Folder folder)
        {
            Deep++;
            Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    // Call EnumerateFolders using childFolder.
                    EnumerateFoldersDeep(childFolder);
                }
            }
        }
    }
}