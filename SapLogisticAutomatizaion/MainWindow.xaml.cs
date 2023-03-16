using Microsoft.Extensions.Configuration;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
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

namespace SapLogisticAutomatizaion
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string[][] materialsData;

        private IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .Build();

        public MainWindow()
        {
            InitializeComponent();
            // tutaj wykonuje sie kod , po uruchomieniu formularza

            //Microsoft.Extensions.Configuration.Binder
            //Microsoft.Extensions.Configuration
            //Microsoft.Extensions.Configuration.Json
            string defaultPath = config.GetValue<string>("InputExcelPath");
            txtMaterialsDataFile.Text = defaultPath;

            seedMateralData();
        }

        private void seedMateralData()
        {
            string excelPath = txtMaterialsDataFile.Text;
            var ecr = new ExcelCollumnsReader();
            materialsData = ecr.ReadExcelFile(excelPath);

            string[] materials = materialsData.Select(x => x[0]).Skip(1).ToArray();
            cbPartNumbers.DataContext = materials;
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            // Clear all textboxes
            txtSerialNumber.Clear();
            txtAddictionalData.Clear();

            // Clear all date pickers
            dtpManufacturingDdate.SelectedDate = null;
            dtpReciptDate.SelectedDate = null;

            // Uncheck all radio buttons
            OKRadioButton.IsChecked = false;
            NORadioButton.IsChecked = false;
            T1RadioButton.IsChecked = false;
            CRadioButton.IsChecked = false;

            // Uncheck all checkboxes
            WoodenBoxCheckBox.IsChecked = false;
            ShippingBoxCheckBox.IsChecked = false;
            PlasticBoxCheckBox.IsChecked = false;

            txtBlockMaterialDesc.Text = string.Empty;
            cbPartNumbers.SelectedIndex = -1;
        }

        private void btnCreateNotification_Click(object sender, RoutedEventArgs e)
        {
            // Read values from the controls
            string partNumber = cbPartNumbers.SelectedItem.ToString();
            string materialDescription = txtBlockMaterialDesc.Text;
            string serialNumber = txtSerialNumber.Text;
            DateTime manufacturingDate = dtpManufacturingDdate.SelectedDate ?? DateTime.MinValue;
            DateTime receiptDate = dtpReciptDate.SelectedDate ?? DateTime.MinValue;
            string additionalData = txtAddictionalData.Text;
            string containerCondition = (OKRadioButton.IsChecked == true) ? "OK" : "NO";
            string customsStatus = (T1RadioButton.IsChecked == true) ? "T1" : "C";
            string containerDetails = "";
            if (WoodenBoxCheckBox.IsChecked == true)
                containerDetails += "Wooden box, ";
            if (ShippingBoxCheckBox.IsChecked == true)
                containerDetails += "Shipping box, ";
            if (PlasticBoxCheckBox.IsChecked == true)
                containerDetails += "Plastic box";

            if (containerDetails.EndsWith(", "))
                containerDetails = containerDetails.Remove(containerDetails.Length - 2);
            // Create a new Product object and fill its properties
            Product product = new Product();
            product.PartNumber = partNumber;
            product.MaterialDescription = materialDescription;
            product.SerialNumber = serialNumber;
            product.ManufacturingDate = manufacturingDate;
            product.ReceiptDate = receiptDate;
            product.AdditionalData = additionalData;
            product.ContainerCondition = containerCondition;
            product.CustomsStatus = customsStatus;
            product.ContainerDetails = containerDetails;

            ExcelDataWriter excelDataWriter = new ExcelDataWriter();
            string defaultPath = config.GetValue<string>("OutputExcelPath");
            excelDataWriter.WriteToExcel(defaultPath, product);
        }

        private void cbPartNumbers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int indx = cbPartNumbers.SelectedIndex;
            if (indx > -1)
            {
                string selectedMaterialDescription = materialsData[indx][1];
                txtBlockMaterialDesc.Text = selectedMaterialDescription;
            }
        }

        private void btnSetPath_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                txtMaterialsDataFile.Text = openFileDialog.FileName;
                seedMateralData();
            }
        }
    }
}