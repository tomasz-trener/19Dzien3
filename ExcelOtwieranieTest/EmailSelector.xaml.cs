using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelOtwieranieTest
{
    /// <summary>
    /// Interaction logic for EmailSelector.xaml
    /// </summary>
    public partial class EmailSelector : System.Windows.Window
    {
        private List<MailItem> filteredEmails;

        public EmailSelector()
        {
            InitializeComponent();
        }

        private void readEmails()
        {
            string filter = txtFilter.Text;
            OutlookEmailSender oes = new OutlookEmailSender();

            filteredEmails = oes.ReadEmail(filter).ToList();
        }

        private void GenerateCheckboxes()
        {
            emailsCheckBoxes.Children.Clear();
            foreach (var e in filteredEmails)
            {
                System.Windows.Controls.CheckBox c = new System.Windows.Controls.CheckBox();
                c.Content = e.Subject + $" ({e.Attachments.Count})";
                c.Tag = e;
                emailsCheckBoxes.Children.Add(c);
            }
        }

        private void SendSelectedEmails_Click(object sender, RoutedEventArgs e)
        {
            List<MailItem> selectedEmails = new List<MailItem>();
            foreach (System.Windows.Controls.CheckBox c in emailsCheckBoxes.Children)
                if ((bool)c.IsChecked)
                    selectedEmails.Add((MailItem)c.Tag);

            OutlookEmailSender oes = new OutlookEmailSender();
            oes.CreateNewEmail("test.recipient@example.com", "test", "aaa", selectedEmails.ToArray());

            //mailItem.To = "test.recipient@example.com";
            //mailItem.Subject = "Test Subject";
            //mailItem.Body = "Test Body";

            //oes.CreateNewEmail(mailItem);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            readEmails();
            GenerateCheckboxes();
        }
    }
}