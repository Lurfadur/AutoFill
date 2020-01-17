using Microsoft.Win32;
using office = Microsoft.Office.Interop;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Runtime.InteropServices;

namespace AutoFill
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        class WorkOrder
        {
            public string Name { get; set; }
            public string Date { get; set; }
            public string OrderNumber { get; set; }
            public string WordPath { get; set; }
            public string ExcelPath{ get; set; }
        }

        // Create new WorkOrder object
        WorkOrder newOrder = new WorkOrder();

        public MainWindow()
        {
            InitializeComponent();
        }

        public bool IsBadInput(string s)
        {
            if (String.IsNullOrEmpty(s))
            {
                return false;
            }
            int i;
            return Int32.TryParse(s, out i);
        }

        private void UserNameData_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox != null)
            {
                string name = textBox.Text.ToString();
                newOrder.Name = name;
            }
        }

        private void OrderNumberData_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox != null)
            {
                newOrder.OrderNumber = textBox.Text.ToString();
            }
        }

        private void SelectDateBox_ValueChanged(object sender, EventArgs e)
        {
            var picker = sender as DatePicker;
            DateTime? date = picker.SelectedDate;
            if (date != null)
            {
                newOrder.Date = date.Value.ToShortDateString();
            }
        }

        private void SelectExcelButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "All Files (*.*)|*.*";
            openFile.FilterIndex = 1;
            openFile.Multiselect = false;

            if (openFile.ShowDialog() == true)
            {
                newOrder.ExcelPath = openFile.FileName;
                ExcelFilePath.Text = newOrder.ExcelPath;
            }
        }

        private void SelectWordButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "All Files (*.*)|*.*";
            openFile.FilterIndex = 1;
            openFile.Multiselect = false;

            if (openFile.ShowDialog() == true)
            {
                newOrder.WordPath = openFile.FileName;
                WordFilePath.Text = newOrder.WordPath;
            }
        }

        private void SaveChangesButton_Click(object sender, RoutedEventArgs e)
        {
            // Do data validation first
            if (IsBadInput(newOrder.Name))
            {
                MessageBox.Show("Name can not contain numbers.", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
                UserNameData.Clear();
            }

            // Create Excel object
            office.Excel.Application xlApp;
            office.Excel.Workbook xlWorkBook;
            office.Excel.Worksheet xlWorkSheet;

            xlApp = new office.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(newOrder.ExcelPath);

            xlWorkSheet = (office.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // Update data entered from user for Excel
            xlWorkSheet.Cells[1, 2] = newOrder.Name;
            xlWorkSheet.Cells[2, 2] = newOrder.Date;
            xlWorkSheet.Cells[3, 2] = newOrder.OrderNumber;
            xlWorkBook.Save();

            // Update data entered from user for Word


            // Release COM resources
            // Excel
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            // Word
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();
            printDialog.UserPageRangeEnabled = true;
        }
    }
}
