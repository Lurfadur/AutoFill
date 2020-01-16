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

namespace AutoFill
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Class contains data information
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
        string EXCEL_PATH = "";

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
            if (date == null)
            {
                
            }
            else
            {
                newOrder.Date = date.Value.ToShortDateString();
            }

        }
        
    }
}
