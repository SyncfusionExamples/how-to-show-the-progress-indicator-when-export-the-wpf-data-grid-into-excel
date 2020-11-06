using System;
using Syncfusion.Data;
using Syncfusion.UI.Xaml.Grid;
using Syncfusion.UI.Xaml.Grid.Helpers;
using System.Windows;
using Syncfusion.XlsIO;
using Syncfusion.UI.Xaml.Grid.Converter;
using Microsoft.Win32;
using System.IO;

namespace SfDataGridDemo
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

        private async void btnExport_Click(object sender, RoutedEventArgs e)
        {
            VisualStateManager.GoToState(this.sfDataGrid, "Busy", true);
            await this.sfDataGrid.Dispatcher.BeginInvoke(new Action(() =>
            {
                ExcelEngine excelEngine = new ExcelEngine();
                var options = new ExcelExportingOptions();
                options.ExcelVersion = ExcelVersion.Excel2013;

                excelEngine = sfDataGrid.ExportToExcel(sfDataGrid.View, options);

                var workBook = excelEngine.Excel.Workbooks[0];
                VisualStateManager.GoToState(this.sfDataGrid, "Normal", true);
                SaveFileDialog sfd = new SaveFileDialog
                {
                    FilterIndex = 2,
                    FileName = "Sample",
                    Filter = "Excel 97 to 2003 Files(*.xls)|*.xls|Excel 2007 to 2010 Files(*.xlsx)|*.xlsx|Excel 2013 File(*.xlsx)|*.xlsx"
                };

                if (sfd.ShowDialog() == true)
                {
                    using (Stream stream = sfd.OpenFile())
                    {

                        if (sfd.FilterIndex == 1)
                            workBook.Version = ExcelVersion.Excel97to2003;

                        else if (sfd.FilterIndex == 2)
                            workBook.Version = ExcelVersion.Excel2010;

                        else
                            workBook.Version = ExcelVersion.Excel2013;
                        workBook.SaveAs(stream);
                    }

                    //Message box confirmation to view the created workbook.

                    if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                                        MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    {

                        //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                        System.Diagnostics.Process.Start(sfd.FileName);
                    }
                }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }
    }
}
         
   

