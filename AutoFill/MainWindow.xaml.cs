using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Controls;

namespace AutoFill
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        class WorkOrder
        {
            public string Name
            {
                get { return Name; }
                set
                {
                    if (value == "")
                    {
                        throw new ArgumentException("Name must not be empty.");
                    }
                    foreach (char item in value)
                    {
                        if (char.IsDigit(item))
                        {
                            throw new ArgumentException("Digits are not allowed.");
                        }
                    }
                    Name = value;
                }
            }
            public string Date
            {
                get { return Date; }
                set
                {
                    if (value != null)
                    {
                        Date = value;
                    }
                }
            }
            public string OrderNumber
            {
                get { return OrderNumber; }
                set
                {
                    if (value == "")
                    {
                        throw new ArgumentException("Order number must not be empty.");
                    }
                    OrderNumber = value;
                }
            }
            public string WordPath
            {
                get { return WordPath; }
                set
                {
                    WordPath = value;
                }
            }
            public string ExcelPath
            {
                get { return ExcelPath; }
                set
                {
                    ExcelPath = value;
                }
            }
        }

        // Create new WorkOrder object
        WorkOrder newOrder = new WorkOrder();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void UserNameData_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox != null)
            {
                newOrder.Name = textBox.Text.ToString();
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
            // TODO: Use MS Interop to update Excel and Word files to keep their existing data and formatting intact

        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();
            printDialog.UserPageRangeEnabled = true;
        }
    }
}
