using Microsoft.Win32;
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

        public MainWindow()
        {
            InitializeComponent();
            // tutaj wykonuje sie kod , po uruchomieniu formularza
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