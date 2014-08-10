using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Configuration;

namespace DATA_UPLOAD
{
    public partial class Form1 : Form
    {
        ExcelFile ef = new ExcelFile();
        SqlConnection SqlConn =null;
        SqlCommand SqlCmd = null;
        SqlTransaction transaction = null;
        int MaxID = 1;
        public Form1()
        {
            InitializeComponent();
        }


        private void cmdBrowse_Click(object sender, EventArgs e)
        {
            progressBar1.Style = ProgressBarStyle.Marquee;

            LoadNewFile();
            cmbSheetName.DataSource = ef.GetExcelSheetNames(txtPath.Text);        //Read All Sheet Name
            progressBar1.Style = ProgressBarStyle.Blocks;
        }

        private void LoadNewFile()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string ExcelFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            ofd.InitialDirectory = ExcelFolder;
            //ofd.Filter = "Excel|*.xls;*.csv;*.xlsx";
            ofd.Filter = "Excel|*.xlsx";
            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                userSelectedFilePath = ofd.FileName;
            }
        }

        public String userSelectedFilePath
        {
            get {
                    return txtPath.Text;
                }
            
            set {
                    txtPath.Text = value;
                }
        }

        private void cmdUpload_Click(object sender, EventArgs e)
        {
            if (txtCircleName.Text == String.Empty)
            {
                MessageBox.Show("Please Enter the Circle Name, and try again");
                return;
            }
            importDataFromExcel(txtPath.Text);

        }

        public void importDataFromExcel(string excelFilePath)
        {

            string sSQLTable = "MIS_TRANS_3RD_PARTY_SURVEY_PROCESS";

            // make sure your sheet name is correct, here sheet name is Sheet1, so you can change your sheet name if have different
            string SheetName = cmbSheetName.Text;
            string myExcelDataQuery = "";

            try
            {
                progressBar1.Style = ProgressBarStyle.Marquee;
                //Create our connection strings

                string sExcelConnectionString = ef.GetExcelDriver(excelFilePath);   //Office Driver
   
                string sSqlConnectionString = ConfigurationManager.ConnectionStrings["SQLCON"].ToString(); //"SERVER=192.168.1.3;USER ID=sa;PASSWORD=;DATABASE=MIS_OLL;CONNECTION RESET=FALSE";

                //Execute a query to erase any previous data from our destination table
                string sClearSQL = "DELETE FROM " + sSQLTable;
                SqlConn = new SqlConnection(sSqlConnectionString);
                SqlCmd = new SqlCommand(sClearSQL, SqlConn);
                SqlConn.Open();

                transaction = SqlConn.BeginTransaction();

                SqlCmd.Transaction = transaction;
                SqlCmd.CommandTimeout = 10800;
                SqlCmd.ExecuteNonQuery();

                //Read TranID 
                String query = @"select (MAX(SURVEY_CIRCLE_ID)+1) as TranID from [MIS_TRANS_3RD_PARTY_SURVEY_CIRCLE]";
                SqlCmd.Connection = SqlConn;
                SqlCmd.CommandText = query;
                SqlDataReader reader = SqlCmd.ExecuteReader();
                while (reader.Read())
                {
                    MaxID = Convert.ToInt32(reader["TranID"]);
                }
                reader.Close();
                reader.Dispose();

                //Validation if exit
                query = @"select count(SURVEY_CIRCLE_ID) as TranID from MIS_TRANS_3RD_PARTY_SURVEY_CIRCLE where [CIRCLE_FROM_DATE]='" + Convert.ToDateTime(dtpFromDate.Value).ToString("yyyy-MM-dd") + "' AND [CIRCLE_TO_DATE]='" + Convert.ToDateTime(dtpToDate.Value).ToString("yyyy-MM-dd") + "'";
                SqlCmd.Connection = SqlConn;
                SqlCmd.CommandText = query;
                reader = SqlCmd.ExecuteReader();
                while (reader.Read())
                {
                   if (Convert.ToInt32(reader["TranID"] ) > 0) 
                   {
                       MessageBox.Show("Sorry! this period is already exist");
                       return;
                   }
                }
                reader.Close();
                reader.Dispose();


                //SqlConn.Close();

                myExcelDataQuery = @"SELECT " + MaxID + ",F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,'A'+LEFT(F21,2),'A'+ left(F21,4),'A'+F21,F23,F24,F25,F26,F27 FROM " + "[" + SheetName + "]"; 
                //Series of commands to bulk copy data from the excel file into our SQL table
                sExcelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1;\"";
                OleDbConnection OleDbConn = new OleDbConnection(sExcelConnectionString);
                OleDbCommand OleDbCmd = new OleDbCommand(myExcelDataQuery, OleDbConn);
                OleDbConn.Open();

                //DataTable dtExcel = new DataTable();
                //OleDbDataAdapter data = new OleDbDataAdapter(myExcelDataQuery, OleDbConn);
                //data.Fill(dtExcel);
                //dataGridView1.DataSource = (dtExcel);

                OleDbDataReader dr = OleDbCmd.ExecuteReader();
                SqlBulkCopy bulkCopy = new SqlBulkCopy(SqlConn , SqlBulkCopyOptions.Default,transaction);

                while (dr.Read())
                {
                    string vQuery = "";

                    vQuery = "DELETE FROM " + sSQLTable;
                    SqlCmd = new SqlCommand(vQuery, SqlConn, transaction);
                    SqlCmd.ExecuteNonQuery();

                    bulkCopy.DestinationTableName = sSQLTable;

                    //bulkCopy.ColumnMappings.Add("CustomerID","ID");       // this method is used if you want to map columns with different names
                    bulkCopy.WriteToServer(dr);

                    SqlCmd = new SqlCommand("UPDATE MIS_TRANS_3RD_PARTY_SURVEY_PROCESS SET UNIT_PRC=ISNULL((SELECT UNIT_PRICE FROM MIS_SYS_3RD_PARTY_PRODUCT WHERE PRODUCT_CODE=MIS_TRANS_3RD_PARTY_SURVEY_PROCESS.VC2_1),0) WHERE SURVEY_CIRCLE_ID IS NULL", SqlConn, transaction);
                    SqlCmd.ExecuteNonQuery();

                    int rowsCopied = SqlBulkCopyHelper.GetRowsCopied(bulkCopy);

                    //Save Master Table [MIS_TRANS_3RD_PARTY_SURVEY_CIRCLE]
                    vQuery = "INSERT INTO MIS_TRANS_3RD_PARTY_SURVEY_CIRCLE ([SURVEY_CIRCLE_ID],[CIRCLE_NAME],[CIRCLE_FROM_DATE],[CIRCLE_TO_DATE],[IMPORT_BY],[IMPORT_DATE]) VALUES (" + MaxID + ",'" + txtCircleName.Text + "','" + Convert.ToDateTime(dtpFromDate.Value).ToString("yyyy-MM-dd") + "','" + Convert.ToDateTime(dtpToDate.Value).ToString("yyyy-MM-dd") + "','','" + Convert.ToDateTime(DateTime.Now.Date).ToString("yyyy-MM-dd") + "')";
                    SqlCmd = new SqlCommand(vQuery, SqlConn, transaction);
                    SqlCmd.ExecuteNonQuery();

                    //Transfer into Details Table 
                    vQuery = "INSERT INTO [MIS_TRANS_3RD_PARTY_SURVEY] ([SURVEY_CIRCLE_ID],[MONTH],[ROUND],[YEAR],[BOOK_ID],[SHOP_ID],[COLECT_DATE],[PSC_DATE],[PSC_TYPE],[PSC_SLNO],[PHY_CODE],[PHY_NAME],[PHY_DEGREE],[GENERIC_NAME],[PRODUCT_CODE],[PRODUCT_NAME],[QT_PRS],[QT_PURCH],[UNIT_PRICE],[DSTMR],[TERRITORY_CODE],[AREA_CODE],[REGION_CODE],[CH_ADDRSS],[CH_DIST],[CH_THANA],[PHY_SPC],[DIAGNAME])" +
                             "SELECT [SURVEY_CIRCLE_ID],[MONTH],[ROUND],[YEAR],[BOOK_ID],[SHOP_ID],[CDATE],[PDATE],[PRS_TYPE],[PSC_SLNO],[PHY_ID],[PHY_NM],[PHY_DEGR],[ING],[VC2_1],[NAME1],[QT_PRS],[QT_PURCH],[UNIT_PRC],[DSTMR],[TERRITORY_CODE],[ORPAM],[ORPRM],[CH_ADD],[CH_DIST],[CH_THANA],[PHY_SPC],[DIAGNAME] FROM MIS_TRANS_3RD_PARTY_SURVEY_PROCESS";
                    SqlCmd = new SqlCommand(vQuery, SqlConn, transaction);
                    SqlCmd.ExecuteNonQuery();

                    //Delete Temporary data
                    vQuery = "TRUNCATE TABLE MIS_TRANS_3RD_PARTY_SURVEY_PROCESS";
                    SqlCmd = new SqlCommand(vQuery, SqlConn, transaction);
                    SqlCmd.ExecuteNonQuery();

                    transaction.Commit();
                    progressBar1.Style = ProgressBarStyle.Blocks;
                    MessageBox.Show("Data Imported Successfully");
                }

                OleDbConn.Close();

            }

            catch (Exception ex)
            {
                //handle exception
                if (transaction != null) transaction.Rollback();
                MessageBox.Show(ex.Message);
                progressBar1.Style = ProgressBarStyle.Blocks;
            }

            finally
            {
                if (transaction != null) transaction.Dispose();
                SqlConn.Close();
            } 

        }

        static class SqlBulkCopyHelper
        {
            static FieldInfo rowsCopiedField = null;

            public static int GetRowsCopied(SqlBulkCopy bulkCopy)
            {
                if (rowsCopiedField == null)
                {
                    rowsCopiedField = typeof(SqlBulkCopy).GetField("_rowsCopied", BindingFlags.NonPublic | BindingFlags.GetField | BindingFlags.Instance);
                }

                return (int)rowsCopiedField.GetValue(bulkCopy);
            }
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
           string myExcelDataQuery = "SELECT * FROM [" + cmbSheetName.Text + "]";    // SheetName.Replace("$","")

            string sExcelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtPath.Text + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1;\"";
            OleDbConnection OleDbConn = new OleDbConnection(sExcelConnectionString);
            OleDbCommand OleDbCmd = new OleDbCommand(myExcelDataQuery, OleDbConn);
            OleDbConn.Open();

            //Preview in DataGrid
            DataTable dtExcel = new DataTable();
            OleDbDataAdapter data = new OleDbDataAdapter(myExcelDataQuery, OleDbConn);
            dataGridView1.DataSource = (dtExcel);
            //data.Fill(dtExcel);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cmbSheetName.SelectedIndex = 0;
        }


        /*
        private void LoadNewFileMulti() 
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string PictureFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            ofd.InitialDirectory = PictureFolder;
            ofd.Title = "Pick a picture; any picture";
            ofd.CustomPlaces.Add(@"C:\");
            ofd.CustomPlaces.Add(@"C:\Program Files\");
            ofd.CustomPlaces.Add(@"K:\Documents\Pictures\");

            ofd.Multiselect = true;

            ofd.Filter = "Pictures|*.jpg; *.bmp; *.png|Documents|*.txt; *.doc; *.log|All|*.*";
            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                //userSelectedFilePath = ofd.FileName;
                foreach (string fileName in ofd.FileNames)
                {
                    userSelectedFilePath += fileName + Environment.NewLine;
                }
            }
        } 
        */

        /*public static int maxValue
        {
            SqlConn = new SqlConnection(cons);
            SqlConn.Open();
            String query = "select max(studentNo) from studentInfo;";
            cmd.Connection = con;
            cmd.CommandText = query;
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            { 
                String x=reader["studentNo"].ToString();

            }
        }*/
    }
}
