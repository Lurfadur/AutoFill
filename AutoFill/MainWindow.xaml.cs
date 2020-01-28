using Microsoft.Win32;
using office = Microsoft.Office.Interop;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

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

        public bool NameIsValid(string s)
        {
            if (String.IsNullOrEmpty(s))
            {
                return false;
            }

            // May only contain a-z and white space
            return Regex.IsMatch(s, @"^[a-zA-Z -]+$");
        }

        public bool DateIsValid(string date)
        {
            if (String.IsNullOrEmpty(date))
            {
                return false;
            }

            char[] delimiterChars = { '-', '/', '.' };
            string[] words = date.Split(delimiterChars);

            // check month
            if (Int32.TryParse(words[0], out int monthVal))
            {
                if ((monthVal < 1 || monthVal > 12))
                {
                    return false;
                }
            }

            // check year
            if (Int32.TryParse(words[2], out int yearVal))
            {
                if (yearVal < DateTime.Now.Year)
                {
                    return false;
                }
            }

            // check day
            if (Int32.TryParse(words[1], out int dayVal))
            {
                if (dayVal < 1)
                {
                    return false;
                }

                // check for leap year
                if (monthVal == 2 && DateTime.IsLeapYear(yearVal))
                {
                    if (dayVal > 29)
                    {
                        return false;
                    }
                }
                else
                {
                    // check for February
                    if (monthVal == 2 && dayVal > 28)
                    {
                        return false;
                    }

                    // months with only 30 days
                    if (monthVal == 4 ||
                        monthVal == 6 ||
                        monthVal == 9 ||
                        monthVal == 11)
                    {
                        if (dayVal > 30)
                        {
                            return false;
                        }
                    }

                    // check for months with only 31 days 
                    else if (monthVal == 1 || 
                        monthVal == 3 ||
                        monthVal == 5 || 
                        monthVal == 7 ||
                        monthVal == 8 ||
                        monthVal == 10 ||
                        monthVal == 12)
                    {
                        if (dayVal > 31)
                        {
                            return false;
                        }
                    }

                    // current month does not exist
                    else
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        public bool OrderIsValid(string order)
        {
            if (string.IsNullOrEmpty(order))
            {
                return false;
            }
            return true;
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
            //var picker = sender as DatePicker;
            DateTime? date = SelectDateBox.SelectedDate;
            if (date.HasValue)
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
            // Check that Name is valid
            if (!NameIsValid(newOrder.Name))
            {
                MessageBox.Show("Name can not be empty or contain numbers.", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
                UserNameData.Clear();
                return;
            }

            // Check that Date is valid
            if (!DateIsValid(newOrder.Date))
            {
                MessageBox.Show("The date is incorrect.", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            // Check that the Order number is valid
            if (!OrderIsValid(newOrder.OrderNumber))
            {
                MessageBox.Show("The order number is incorrect.", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Create Excel object
            office.Excel.Application xlApp;
            office.Excel.Workbook xlWorkBook;
            office.Excel.Worksheet xlWorkSheet;

            xlApp = new office.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(newOrder.ExcelPath);

            xlWorkSheet = (office.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // Update data entered from user into Excel
            xlWorkSheet.Cells[1, 2] = newOrder.Name;
            xlWorkSheet.Cells[2, 2] = newOrder.Date;
            xlWorkSheet.Cells[3, 2] = newOrder.OrderNumber;
            xlWorkBook.Save();

            // Release Excel resources
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            // Create Word object


            // Update data entered from user for Word


            // Release Word resources

            // Set ability to print
            PrintButton.IsEnabled = true;
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();
            printDialog.UserPageRangeEnabled = true;

            // After printing is completed, disable button again
            PrintButton.IsEnabled = false;
        }
    }
}
