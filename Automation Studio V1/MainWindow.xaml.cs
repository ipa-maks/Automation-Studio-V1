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
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb; // Excel Import
using Microsoft.Win32; // OpenFileDialog

namespace Automation_Studio_V1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            FillDataGrid();
        }

        private void FillDataGrid()
        {
            string ConString = ConfigurationManager.ConnectionStrings["Automation_Studio_V1.Properties.Settings.APAV1ConnectionString"].ConnectionString;
            string CmdString = string.Empty;
            using (SqlConnection con = new SqlConnection(ConString))
            {
                CmdString = "SELECT Id, Name FROM Test_Factory";
                SqlCommand cmd = new SqlCommand(CmdString, con);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable("Test_Factory");
                sda.Fill(dt);
                ShowData.ItemsSource = dt.DefaultView;
            }
        }

        // Tab: Data Input

        private void ShowExcelButton_Click(object sender, RoutedEventArgs e)
        {
            string excelpath = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ExcelFilePath.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection excelconn = new OleDbConnection(excelpath);

            OleDbDataAdapter exceldataadapter = new OleDbDataAdapter("Select * From [" + Sheet.Text + "$]", excelconn);
            DataTable exceldt = new DataTable();

            exceldataadapter.Fill(exceldt);

            ExcelData.ItemsSource = exceldt.DefaultView;
        }

        private void ChooseExcelButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfiledialogexcel = new OpenFileDialog();

            if (openfiledialogexcel.ShowDialog() == true)
            {
                this.ExcelFilePath.Text = openfiledialogexcel.FileName;
            }
        }

        private void CreateTabelButton_Click(object sender, RoutedEventArgs e)
        {
            string ConString1 = ConfigurationManager.ConnectionStrings["Automation_Studio_V1.Properties.Settings.APAV1ConnectionString"].ConnectionString;
            string CmdString1 = string.Empty;
            using (SqlConnection con1 = new SqlConnection(ConString1))
            {
                CmdString1 = "CREATE TABLE Region (RegionID varchar(50), RegionDescription varchar(50))";
                SqlCommand cmd1 = new SqlCommand(CmdString1, con1);
                SqlDataAdapter sda = new SqlDataAdapter(cmd1);


            }
        }
    }
}
