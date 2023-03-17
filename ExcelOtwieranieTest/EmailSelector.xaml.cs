using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
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
using static System.Net.WebRequestMethods;

namespace ExcelOtwieranieTest
{
    /// <summary>
    /// Interaction logic for EmailSelector.xaml
    /// </summary>
    public partial class EmailSelector : System.Windows.Window
    {
        private List<MailItem> filteredEmails;
        private OutlookEmailSender oes = new OutlookEmailSender();

        public EmailSelector()
        {
            InitializeComponent();
        }

        private void emptyAttachementFolder()
        {
            DirectoryInfo di = new DirectoryInfo(oes.AttachementPatch);
            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }

        private void GenerateCheckboxes()
        {
            emailsCheckBoxes.Children.Clear();
            foreach (var e in filteredEmails)
            {
                System.Windows.Controls.CheckBox c = new System.Windows.Controls.CheckBox();
                c.Content = e.Subject + $" ({e.Attachments.Count})";
                c.Tag = e;

                foreach (Attachment at in e.Attachments)
                    c.ToolTip += at.FileName + ", ";

                emailsCheckBoxes.Children.Add(c);
            }
        }

        private void SendSelectedEmails_Click(object sender, RoutedEventArgs e)
        {
            List<MailItem> selectedEmails = new List<MailItem>();
            foreach (System.Windows.Controls.CheckBox c in emailsCheckBoxes.Children)
                if ((bool)c.IsChecked)
                    selectedEmails.Add((MailItem)c.Tag);

            oes.CreateNewEmail("test.recipient@example.com", "test", "aaa", selectedEmails.ToArray());

            //mailItem.To = "test.recipient@example.com";
            //mailItem.Subject = "Test Subject";
            //mailItem.Body = "Test Body";

            //oes.CreateNewEmail(mailItem);
        }

        private async void btnShowEmails_Click(object sender, RoutedEventArgs e)
        {
            var progress = new Progress<int>(value => progressBar.Value = value);  // funkcja anonimowa, która bedzie ustawiac progressbar na zadaną wartość
            oes.Progress = progress; //OutlookEmailSender musi posiadac referencje do progressu bo bedzie ze środa sterować tym progresem
            oes.CalculateDeep(); // musimy najpierw policzyć głębokość czyli liczbę folderów do przeiterowania
            // dlaczego? bo zeby znać % progresu to trzeba wiedzieć jaka jest całkowita liczba folderów do przetworzenia

            progressBar.Maximum = (int)oes.Deep; // ustawiamy progress bar na tę obliczoną wcześniej głębokość

            emptyAttachementFolder();

            string filter = txtFilter.Text;

            await Task.Run(() =>
            {
                filteredEmails = oes.ReadEmailRecur(filter).ToList();
            });
            GenerateCheckboxes();
        }
    }
}