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
        // Variables defined at startup
        string ConString = ConfigurationManager.ConnectionStrings["Automation_Studio_V1.Properties.Settings.APAV1ConnectionString"].ConnectionString; // Shows the path to the used Database
        string SelectedDataTable;

        private SqlDataAdapter PrivatAdapter;
        private DataTable PrivatTable;
        private SqlCommandBuilder cmbl;


        public MainWindow()
        {
            // Methods called at startup
            InitializeComponent();
            FillDataGrid(ConString);






        }

        private void FillDataGrid(string ConString)
        {
            string CmdString = "SELECT Id, Name FROM Test_Factory";

            using (SqlConnection con = new SqlConnection(ConString))
            {
                SqlDataAdapter sda = new SqlDataAdapter(CmdString, con);
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

// Create New Table

        private void CreateTabelButton_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection con = new SqlConnection(ConString))
            {
                try
                {
                    con.Open();
                    using (SqlCommand command = new SqlCommand("CREATE TABLE Factory_2(Id char(50),Name char(50));", con))
                        command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        // Show available tables

        private void ShowTablesButton_Click(object sender, RoutedEventArgs e)
        {
            GetTables();
        }

        public List<string> GetTables()
        {
            using (SqlConnection connection = new SqlConnection(ConString))
            {
                connection.Open();
                DataTable schema = connection.GetSchema("Tables");
                List<string> TableNames = new List<string>();
                foreach (DataRow row in schema.Rows)
                {
                    TableNames.Add(row[2].ToString());
                }
                ComboBox1.ItemsSource = TableNames;
                return TableNames;
            }
        }

        // Get ComboBox string

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // ... Get the ComboBox.
            var ComboBox1 = sender as ComboBox;

            // ... Set SelectedItem as Window Title.
            SelectedDataTable = ComboBox1.SelectedItem as string;
            this.Title = "Selected: " + SelectedDataTable;
        }

        // Fill Data Grid: Second Tab

        private void ShowTableInGridButton_Click(object sender, RoutedEventArgs e)
        {
            FillDataGrid2();
        }

        private void FillDataGrid2()
        {
            String SelectedDataTable = ComboBox1.Text;


            ExcelFilePath.Text = SelectedDataTable;
            // string CmdString = "SELECT Id, Name FROM @Table";
            string CmdString = String.Format("SELECT Id, Name FROM {0}", SelectedDataTable);
            // string CmdString = "SELECT Id, Name FROM Test_Factory";
            using (SqlConnection con = new SqlConnection(ConString))
            {
                SqlCommand cmd = new SqlCommand(CmdString, con);
                // cmd.Parameters.AddWithValue("@table", ComboBox1.Text);
             PrivatAdapter = new SqlDataAdapter(cmd);
            // PrivatAdapter = new SqlDataAdapter(CmdString, con);
            PrivatTable = new DataTable();

            PrivatAdapter.Fill(PrivatTable);
            ExcelData.ItemsSource = PrivatTable.DefaultView;
            }
        }

        //private void UpdateButton_Click(object sender, RoutedEventArgs e)
        //{
        //        PrivatAdapter.Update(APAV1DataSet.Tables[SelectedDataTable]);
        //}
    }
}
        
    


    







    
