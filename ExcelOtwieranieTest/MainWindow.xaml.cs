using Microsoft.Office.Interop.Outlook;
using SapLogisticAutomatizaion;
using System;
using System.Collections.Generic;
using System.Data;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelOtwieranieTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnSendEmailFromExcel_Click(object sender, RoutedEventArgs e)
        {
            ExcelCollumnsReader ecr = new ExcelCollumnsReader();
            var jaggedArray = ecr.ReadExcelFile(@"C:\dane\Excel\Book1.xlsx");

            DataTable dataTable = new DataTable();
            //for (int i = 0; i < jaggedArray[0].Length; i++)
            //    dataTable.Columns.Add($"Column{i + 1}");
            //int j = 0;
            foreach (var item in jaggedArray[0])
                dataTable.Columns.Add(item);

            for (int i = 1; i < jaggedArray.Length; i++)
                dataTable.Rows.Add(jaggedArray[i]);

            dgvData.ItemsSource = dataTable.DefaultView;

            OutlookEmailSender oes = new OutlookEmailSender();
            oes.CreateEmail(jaggedArray);
        }

        private void btnForwardEmail_Click(object sender, RoutedEventArgs e)
        {
            EmailSelector es = new EmailSelector();
            es.Show();
        }
    }
}