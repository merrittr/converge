/*
              Direct   2nd       3rd       Youngs
              in Deg   Direc     Direc     Modulus
              28.00    152.00     0.00     75.00
 20-Jul-92     0.00      0.00     0.00
 22-Jul-92   -80.58     -7.27   -35.09
 27-Jul-92  -142.84      3.63   -39.22
 31-Jul-92   -76.91     69.03    49.54
 05-Aug-92  -402.89   -108.99  -103.20
 10-Aug-92  -168.48    159.85    80.50
 17-Aug-92  -162.99     56.31     1.03
 19-Aug-92  -161.15     76.29    12.38
 03-Sep-92   263.71    570.38   334.37
 29-Sep-92   304.00    519.52   470.60
 20-Oct-92   651.94    156.22   460.28
 09-Nov-92   688.57    319.70   532.52
 12-Dec-92  3083.91  -2317.85 -1120.77
 07-Jan-93  1640.85  -1732.94 -1159.98
 04-Feb-93  1871.59  -1849.19 -1133.15
 09-Feb-93  2153.61  -2332.38 -1376.71
 03-Mar-93  2021.76  -2336.02 -1440.69

The first column consists of the dates at which measurments were taken.  The
first column adjacent to the dates are the measurments in microstrain.  The
value at the top of the second column, 28, is the corresponding measuring
direction, in degrees, measured counter-clockwise from horizontal.  152 is
the measuring direction for the second column of microstrain measurements and
0 (horizontal) is the third measuring direction for the final column of
microstrain measurements.  The single value in the fifth column, 75.00, is
the input value of Young's Modulus, given in GPa.  
 */


using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.OleDb;
using Microsoft.SqlServer.MessageBox;
//Excel = 
namespace converge
{
    




    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public class gbl
        {
            public static string connectionString;
            public static string direction1;
            public static string direction2;
            public static string direction3;
            public static string youngsMod;
            public static string cellVal;
            public static string fileName;
            public static string av = "0";
            //public static OleDbConnection oConn;
        }



       string[] ConvertToStringArray(System.Array values)
       {
           string[] theArray = new string[values.Length];
           for (int i = 1; i <= values.Length; i++)
           {
               if (values.GetValue(1, i) == null)
                   theArray[i - 1] = "";
               else
                   theArray[i - 1] = (string)values.GetValue(1, i).ToString();
           }
           return theArray;
       }

//////////////////////////////////////////////////////////////////////////

       private void button2_Click(object sender, EventArgs e)
       {
           gbl.av = "0";
           this.openFileDialog1.FileName = "*.xls;*.xlsx;*.csv";
           if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
           {
               dataGridView1.DataSource = null;
               gbl.fileName = openFileDialog1.FileName;

               string _Extension = Path.GetExtension(gbl.fileName);
               if (!(_Extension.Equals(".csv", StringComparison.CurrentCultureIgnoreCase)))
               {
                   dataGridView1.DataSource = ImportExcel(gbl.fileName);
                   dataGridView1.Columns[0].Visible = false;
                   dataGridView1.Columns[1].Visible = false;
                   dataGridView1.Columns[2].Visible = false;
                   dataGridView1.Columns[3].Visible = false;
                   dataGridView1.Columns[4].Visible = false;
                   dataGridView1.Columns[5].Visible = false;
                   dataGridView1.Columns[6].Visible = false;
                   dataGridView1.Columns[8].Visible = false;
                   dataGridView1.Columns[9].Visible = false;
                   //dataGridView1.AutoGenerateColumns = false;
               }
               else
               {
                   gbl.fileName = this.openFileDialog1.FileName;
                   //string _Extension = Path.GetExtension(gbl.fileName);
                   this.dataGridView1.DataSource = ImportExcel(gbl.fileName);
                   string what = Path.GetFileName(gbl.fileName) + " ------ " + Path.GetDirectoryName(gbl.fileName);
                   OleDbConnection oConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.GetDirectoryName(gbl.fileName) + ";Extended Properties=\"text;HDR=No;FMT=Delimited\";");
                   OleDbDataAdapter oAr = new OleDbDataAdapter("select * from " + Path.GetFileName(gbl.fileName), oConn);
                   DataTable mDt = new DataTable("temp");
                   oAr.Fill(mDt);
                   this.dataGridView1.DataSource = mDt;
                   this.dataGridView1.Visible = true;
               }
           }
       }
//////////////////////////////////////////////////////////////////////////
       /// <summary>  

       /// Imports Data from Microsoft Excel File. 
       /// </summary>  
       /// <param name="FileName">Filename from which data need to import</param>  
       /// <returns>List of DataTables, based on the number of sheets</returns>  
       private List<DataTable> ImportExcel(string FileName)
       {
           List<DataTable> _dataTables = new List<DataTable>();
           //string _ConnectionString = string.Empty;
           string _Extension = Path.GetExtension(FileName);
           //Checking for the extentions, if XLS connect using Jet OleDB  
           if (_Extension.Equals(".xls", StringComparison.CurrentCultureIgnoreCase))
           {
               //gbl.connectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source={0};Extended Properties=Excel 8.0;";
               gbl.connectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source={0};Extended Properties=\"Excel 8.0;HDR=No;\";";
           }
           //Use ACE OleDb  
           else if (_Extension.Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase))
           {
               //gbl.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 8.0;";
               gbl.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=No;\";";
           }

          
           DataTable dataTable = null;
           //////////////////////////////////////////////////////////////////////////////////////////////////////////////
           if (!(_Extension.Equals(".csv", StringComparison.CurrentCultureIgnoreCase)))
           {
               using (OleDbConnection oleDbConnection = new OleDbConnection(string.Format(gbl.connectionString, FileName)))
               {
                   try { oleDbConnection.Open(); }
                   catch (Exception e)
                   {
                       ExceptionMessageBox box = new ExceptionMessageBox(e);
                       box.Show(this);
                       //Console.WriteLine("{0} Exception caught.", e);
                       //throw new Exception();

                   }

                   //Getting the meta data information.  
                   //This DataTable will return the details of Sheets in the Excel File.  
                   DataTable dbSchema = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables_Info, null);
                   foreach (DataRow item in dbSchema.Rows)
                   {
                       //reading data from excel to Data Table  
                       using (OleDbCommand oleDbCommand = new OleDbCommand())
                       {
                           oleDbCommand.Connection = oleDbConnection;
                           oleDbCommand.CommandText = string.Format("SELECT * FROM [{0}]", item["TABLE_NAME"].ToString());
                           //oleDbCommand.CommandText = string.Format("SELECT 'TableName' FROM [{0}]", item["TABLE_NAME"].ToString());
                           using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter())
                           {
                               oleDbDataAdapter.SelectCommand = oleDbCommand;
                               dataTable = new DataTable(item["TABLE_NAME"].ToString());
                               oleDbDataAdapter.Fill(dataTable);
                               _dataTables.Add(dataTable);
                           }
                       }
                   }
               }
           }
           else
           {
               string what = Path.GetFileName(FileName) + " ------ " + Path.GetDirectoryName(FileName);
               OleDbConnection oConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +  Path.GetDirectoryName(FileName)+ ";Extended Properties=\"text;HDR=No;FMT=Delimited\";");
               OleDbDataAdapter oAr = new OleDbDataAdapter("select * from " + Path.GetFileName(FileName),oConn);
               DataTable mDt = new DataTable("temp");
               oAr.Fill(mDt);
               _dataTables.Add(mDt);
               //dataGridView1.DataSource = mDt;
               //this.dataGridView1.Visible = true;
           }
          /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
           return _dataTables;
       }



       void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
       {
           //gbl.cellVal = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString();
           //MessageBox.Show("double " + gbl.cellVal);
           //dataGridView1.DataSource = null;
           //dataGridView1.DataSource = GetWorksheet(gbl.cellVal);
           dataGridView1_CellContentClick(sender, e);
       }

       private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
       {
           if (gbl.av == "0")
           {
               gbl.cellVal = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString();
               //MessageBox.Show("single " + gbl.cellVal);
               dataGridView1.DataSource = null;
               dataGridView1.DataSource = GetWorksheet(gbl.cellVal);
               gbl.av = "4";
           }
           /*else if (gbl.av == "1")
           {
               textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString();
           }
           else if (gbl.av == "2")
           {
               textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString();
           }
           else if (gbl.av == "3")
           {
               textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString();
           }
           else if (gbl.av == "4")
           {
                   textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString();
                   gbl.av = "5";
           }
           else if (gbl.av == "5")
           {
                   textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString();
                   gbl.av = "6";
           }
           else if (gbl.av == "6")
           {
               textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString();
               gbl.av="4";
           }*/
           try 
           {
               textBox1.Text = dataGridView1.Rows[0].Cells[1].FormattedValue.ToString();
           } 
           catch (ApplicationException ex)
           {
               string str = "The action failed.";
               ApplicationException exTop = new ApplicationException(str, ex);
               exTop.Source = this.Text;
               ExceptionMessageBox box = new ExceptionMessageBox(exTop);
               box.Show(this);
               Console.WriteLine("{0} Exception caught.", ex);
               throw new Exception();
           } 
           try 
           {
               textBox2.Text = dataGridView1.Rows[0].Cells[2].FormattedValue.ToString();
           } 
           catch (ApplicationException ex)
           {
               string str = "The action failed.";
               ApplicationException exTop = new ApplicationException(str, ex);
               exTop.Source = this.Text;
               ExceptionMessageBox box = new ExceptionMessageBox(exTop);
               box.Show(this);
               Console.WriteLine("{0} Exception caught.", ex);
               throw new Exception();
           } 
           try 
           {
               textBox3.Text = dataGridView1.Rows[0].Cells[3].FormattedValue.ToString();
           } 
           catch (ApplicationException ex)
           {
               string str = "The action failed.";
               ApplicationException exTop = new ApplicationException(str, ex);
               exTop.Source = this.Text;
               ExceptionMessageBox box = new ExceptionMessageBox(exTop);
               box.Show(this);
           } 
           try
           { 
               textBox4.Text = dataGridView1.Rows[0].Cells[4].FormattedValue.ToString(); 
           } 
           catch (ApplicationException ex)
           {
               string str = "The action failed.";
               ApplicationException exTop = new ApplicationException(str, ex);
               exTop.Source = this.Text;
               ExceptionMessageBox box = new ExceptionMessageBox(exTop);
               box.Show(this);
           } 
       }

       public System.Data.DataTable GetWorksheet(string worksheetName)
       {
           //MessageBox.Show(gbl.connectionString);
           //OleDbConnection con = new System.Data.OleDb.OleDbConnection(gbl.connectionString,gbl.fileName);
           //new OleDbConnection(string.Format(gbl.connectionString, gbl.fileName));
           //OleDbDataAdapter cmd = new System.Data.OleDb.OleDbDataAdapter("select * from [" + worksheetName.Replace("$","") + "]", con);

           OleDbConnection con = new OleDbConnection(string.Format(gbl.connectionString, gbl.fileName));



           OleDbDataAdapter cmd = new System.Data.OleDb.OleDbDataAdapter("select * from [" + worksheetName + "]", con);
           con.Open();
           System.Data.DataSet excelDataSet = new DataSet();
           cmd.Fill(excelDataSet);
           con.Close();

           return excelDataSet.Tables[0];
       }

      /*
        public void textBox1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
       {
           //msgbox("select the cell with the angle value1");
           gbl.av = "1";
       }

       private void textBox2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
       {
           //msgbox("select the cell with the angle value2");
           gbl.av = "2";
       }

       private void textBox3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
       {
           //msgbox("select the cell with the angle value3");
           gbl.av = "3";
       }
       */



//////////////////////////////////////////////////////////////////////////
    }
}



