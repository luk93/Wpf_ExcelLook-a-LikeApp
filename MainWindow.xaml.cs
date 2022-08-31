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
using OfficeOpenXml;
using System.Data;
using System.Reflection;
using System.IO;
using Microsoft.Win32;
using System.Timers;

namespace Wpf_ExcelLook_a_LikeApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        DataTable dataTable = new();
        FileInfo file = null;
        bool autoUpdate = false;
        Timer timerUpdate = new();

        public static int row_PopUp = 2;
        public static int col_PopUp = 3;
        public static string wsName_PopUp = "03 ES";
        public MainWindow()
        {
            InitializeComponent();

            System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
            dispatcherTimer.Tick += dispatcherTimer_Tick;
            dispatcherTimer.Interval = TimeSpan.FromSeconds(1);
            dispatcherTimer.Start();
        }

        private async void Button1_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            OpenFileDialog openFileDialog1 = new()
            {
                InitialDirectory = @"c:\Users\localadm\Desktop",
                Title = "Select some  excel file",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xlsx",
                Filter = "Excel file (*.xlsx)|*.xlsx",
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true,
            };
            if (openFileDialog1.ShowDialog() == true)
            {
                file = new FileInfo(openFileDialog1.FileName);
                if (file.Exists && !IsFileLocked(file.FullName))
                {
                    PopUp popUp = new PopUp();
                    if (popUp.ShowDialog() == true)
                    {
                        dataTable = await LoadFromExcel(file, row_PopUp, col_PopUp, wsName_PopUp);
                        await FillGridWithData(dataGrid, dataTable);
                    }
                }
            }
        }
        private static async Task<DataTable> LoadFromExcel(FileInfo file, int row, int col, string wsName)
        {
            int defaultCol = col;
            int defaultRow = row;
            DataTable output = new();
            int colAmount = 0;
            using var exPackage = new ExcelPackage(file);
            await exPackage.LoadAsync(file);
            var ws = exPackage.Workbook.Worksheets[wsName];
            //Get Columns
            while (!string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()))
            {
                output.Columns.Add(ws.Cells[row, col].Value.ToString());
                col++;
                colAmount++;
            }
            col = defaultCol;
            row = defaultRow + 1;
            //Get Rows
            while (!string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()))
            {
                output.Rows.Add(ws.Cells[row, col].Value.ToString());
                for (int i = 0; i < colAmount; i++)
                {
                    if (!string.IsNullOrWhiteSpace(ws.Cells[row, col + i].Value?.ToString()))
                    {
                        output.Rows[row - (defaultRow + 1)][i] = ws.Cells[row, col + i].Value.ToString();
                    }
                    else
                    {
                        output.Rows[row - (defaultRow + 1)][i] = "";
                    }

                }
                row++;
            }
            return output;
        }
        public static bool IsFileLocked(string filePath)
        {
            try
            {
                var stream = File.OpenRead(filePath);
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }
        private static async Task FillGridWithData(DataGrid dataGrid, DataTable dataTable)
        {
            dataGrid.ItemsSource = dataTable.DefaultView;
        }
        private async void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            if (file != null && !IsFileLocked(file.FullName) && autoUpdate && row_PopUp != 0)
            {
                try
                {
                    dataTable = await LoadFromExcel(file, row_PopUp, col_PopUp, wsName_PopUp);
                    await FillGridWithData(dataGrid, dataTable);
                }
                catch (Exception ex)
                {

                }
            }
        }

        private void autoUpdateButton_Click(object sender, RoutedEventArgs e)
        {
            if (autoUpdate)
            {
                autoUpdate = false;
                //autoUpdateButton.Background = Color.FromRgb(FFF,53B,955);
                autoUpdateButton.Content = "Auto Update OFF";

            }
            else
            {
                autoUpdate = true;
                //autoUpdateButton.Background = Color.Green;
                autoUpdateButton.Content = "Auto Update ON";
            }
        }
    }
}
