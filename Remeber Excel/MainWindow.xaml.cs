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
using System.Reflection;
using OfficeOpenXml;
using System.IO;

namespace Remeber_Excel
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        private void ButtonClick(object sender, RoutedEventArgs e)
        {
            var excelDialog = new OpenFileDialog
            {
                InitialDirectory = AppDomain.CurrentDomain.BaseDirectory,
                Title = "Select the excel file you want to read",
                Filter = "Excel 2010 document (*.xlsx)|*.xlsx|Excel 2010 Macro-enabled document (*.xlsm)|*.xlsm|Excel 97-2003 document (*.xls)|*.xls",
                Multiselect = false
            };
            if (excelDialog.ShowDialog() == true)
                try
                {
                    using (var excelFile = new ExcelPackage(new FileInfo(excelDialog.FileName)))
                    {
                        ExcelWorksheet Sheet = excelFile.Workbook.Worksheets[1];
                        StringBuilder sBuilder = new StringBuilder();
                        var start = Sheet.Dimension.Start;
                        var end = Sheet.Dimension.End;
                        for (int i = 1; i <= end.Row; i++)
                        {
                            if (Sheet.Cells[i, 1].Value.ToString() == "Granberg, David")
                            {
                                Sheet.Cells[i, 1].Value = "Nope";
                            }
                            sBuilder.Append((string)Sheet.Cells[i, 1].Value);
                        }
                        textBox.Text = sBuilder.ToString();

                        excelFile.Save();
                    }
                }
                catch
                {
                    textBox.Text = "crap";
                }
        }

        
    }
}
