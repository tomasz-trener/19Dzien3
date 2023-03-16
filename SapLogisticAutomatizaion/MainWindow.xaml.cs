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
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            // Clear all textboxes
            txtPartNumber.Clear();
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
        }

        private void btnCreateNotification_Click(object sender, RoutedEventArgs e)
        {
        }
    }
}