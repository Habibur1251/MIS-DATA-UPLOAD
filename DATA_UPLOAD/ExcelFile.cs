using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Reflection;
using System.IO;

namespace DATA_UPLOAD
{
    public class ExcelFile
    {

        public List<string> GetExcelSheetNames(string excelFile)
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            try
            {
                // Connection String. Change the excel file to the file you
                ExcelFile ef = new ExcelFile();
                String connString = ef.GetExcelDriver(excelFile); 

                // Create connection object by using the preceding connection string.
                objConn = new OleDbConnection(connString);
                // Open connection with the database.
                objConn.Open();
                // Get the data table containg the schema guid.
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;
                //DataTable dtN = new DataTable();
               List<string> list=new List<string>();
                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    //excelSheets[i] = row["TABLE_NAME"].ToString();
                    list.Add(row["TABLE_NAME"].ToString());//.Rows.Add(row["TABLE_NAME"].ToString());
                    i++;
                }

                return list;
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        public string GetExcelDriver(string ExcelFile)
        {
            string ExcelDriver = "";

            if (Path.GetExtension(ExcelFile) == ".xls")
            {
                //for Office XP
                ExcelDriver = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ExcelFile + ";Extended Properties=Excel 8.0;HDR=Yes"; 
            }
            else if (Path.GetExtension(ExcelFile) == ".xlsx")
            {
                //for Office 2007
                ExcelDriver = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFile + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            }

            return ExcelDriver;
        }
    }


}
