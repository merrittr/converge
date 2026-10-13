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
using System.Drawing.Design;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.OleDb;
using Microsoft.SqlServer.MessageBox;
using System.Diagnostics;
//Excel = 
namespace converge
{


    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tabControl1.SelectedIndex = 0;
            tabControl1.SelectedTab.Text = "Data";
            tabControl1.SelectedIndex = 1;
            tabControl1.SelectedTab.Text = "Matrix And Inverted Matrix";
            tabControl1.SelectedIndex = 2;
            tabControl1.SelectedTab.Text = "Answer Matrix";
            tabControl1.SelectedIndex = 3;
            //tabControl1.SelectedTab.Text = "Reformat Data";
            //tabControl1.TabPages.Remove(tabPage4);
            tabControl1.SelectedIndex = 0;
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
           //robm 1/2/2012
           tabControl1.SelectedIndex = 0;
           button1.Enabled = false;
           //robm 1/2/2012
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
                   if (!(dataGridView1.DataSource == null))
                   {
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
                   textBox1.Text = dataGridView1.Rows[0].Cells[1].FormattedValue.ToString();
                   textBox2.Text = dataGridView1.Rows[0].Cells[2].FormattedValue.ToString();
                   textBox3.Text = dataGridView1.Rows[0].Cells[3].FormattedValue.ToString();
                   textBox4.Text = dataGridView1.Rows[0].Cells[4].FormattedValue.ToString();
                   button1.Enabled = true;
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
                       //ExceptionMessageBox box = new ExceptionMessageBox(e);                       
                       //box.Show(this);
                       MessageBox.Show(e.ToString(), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                       return null;
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
                               if (!(item["TABLE_NAME"].ToString() == "SAMPLE"))
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
           if (gbl.fileName.IndexOf("csv") < 1)
           {
               dataGridView1_CellContentClick(sender, e);
           }
       }

       private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
       {
           if (gbl.av == "0" && (gbl.fileName.IndexOf("csv") < 1) )
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
               if (dataGridView1.Rows.Count > 1)
               {
                   textBox1.Text = dataGridView1.Rows[0].Cells[1].FormattedValue.ToString();
               }
               else
               {
                   textBox1.Text = "";
               }
           } 
           catch (ApplicationException ex)
           {
               //string str = "The action failed.";
               //ApplicationException exTop = new ApplicationException(str, ex);
               //exTop.Source = this.Text;
               //ExceptionMessageBox box = new ExceptionMessageBox(ex);
               //box.Show(this);
               MessageBox.Show(ex.ToString(), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);

               return;
           } 
           try 
           {
               if (dataGridView1.Rows.Count > 1)
               {
                   textBox2.Text = dataGridView1.Rows[0].Cells[2].FormattedValue.ToString();
               }
               else
               {
                   textBox2.Text = "";
               }
           } 
           catch (ApplicationException ex)
           {
               //string str = "The action failed.";
               //ApplicationException exTop = new ApplicationException(str, ex);
               //exTop.Source = this.Text;
               //ExceptionMessageBox box = new ExceptionMessageBox(ex);
               //box.Show(this);
               MessageBox.Show(ex.ToString(), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
              //Console.WriteLine("{0} Exception caught.", ex);
              //throw new Exception();
               return;
           } 
           try 
           {
               if (dataGridView1.Rows.Count > 1)
               {
                   textBox3.Text = dataGridView1.Rows[0].Cells[3].FormattedValue.ToString();
               }
               else
               {
                   textBox3.Text = "";
               }
           } 
           catch (ApplicationException ex)
           {
               //string str = "The action failed.";
               //ApplicationException exTop = new ApplicationException(str, ex);
               //exTop.Source = this.Text;
               //ExceptionMessageBox box = new ExceptionMessageBox(ex);
               //box.Show(this);
               MessageBox.Show(ex.ToString(), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
               return;
           } 
           try
           {
               if (dataGridView1.Rows.Count > 1)
               {
                   try
                   {
                       textBox4.Text = dataGridView1.Rows[0].Cells[4].FormattedValue.ToString();
                   }
                   catch
                   {
                       textBox4.Text = "";
                   }
               }
               else
               {
                   textBox4.Text = "";
               }
           } 
           catch (ApplicationException ex)
           {
               string str = "The action failed.";
               ApplicationException exTop = new ApplicationException(str, ex);
               exTop.Source = this.Text;
               //ExceptionMessageBox box = new ExceptionMessageBox(exTop);
               //box.Show(this);
               MessageBox.Show(exTop.ToString(), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
               return;
           }


           //put code to check & convert raw
           bool GridValid = true;
           bool NoParams = false;
           bool NoDates = false;
           string GridMsg = "no message";
           int Num;
           ///////////////////////////////////////////////////////////////////////////////////////////////
           if ((dataGridView1.Rows[0].Cells[1].FormattedValue.ToString().Trim() == "") || (dataGridView1.Rows[0].Cells[2].FormattedValue.ToString().Trim() == "") || (dataGridView1.Rows[0].Cells[3].FormattedValue.ToString().Trim() == ""))
           {
               GridValid = false;
               NoParams = true;
           }
           if (!(dataGridView1.Rows[0].Cells[0].FormattedValue.ToString().Trim() == "")) 
           {
               GridValid = false;
               NoParams = true;
           }
           if (!(dataGridView1.Rows[0].Cells[0].FormattedValue.ToString().Trim() == ""))
           {
               GridValid = false;
           }
           if (!(Int32.TryParse(dataGridView1.Rows[0].Cells[1].FormattedValue.ToString(), out Num)))
           {
               GridValid = false;
           }
           if (!(Int32.TryParse(dataGridView1.Rows[0].Cells[2].FormattedValue.ToString(), out Num)))
           {
               GridValid = false;
           }
           if (!(Int32.TryParse(dataGridView1.Rows[0].Cells[3].FormattedValue.ToString(), out Num)))
           {
               GridValid = false;
           }

           try
           {
               System.DateTime dt = System.DateTime.Parse(dataGridView1.Rows[1].Cells[0].Value.ToString());
           }
           catch
           {
               GridValid = false;
               NoDates = true;
           }

           if (GridValid)
           {
              // MessageBox.Show("The grid is valid", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
               button1.Enabled = true;
           }
           else
           {
               if ((GridMsg.IndexOf("no message") > -1))
               {
                   GridMsg = "The grid is not valid";
               }

               if (NoParams)
               {
                   GridMsg += "\n\n**Parameters missing or malformed \n**Angles 1 2 3 Young's modulus or Poissons's Ratio";
                   NoParams = false;
               }
               if (NoDates)
               {
                   GridMsg += "\n\n**Dates for reading should exist in the lefmost column\n**The first cell top left most should however be blank";
                   NoDates = false;
               }
               GridMsg += "\n\nplease examine the example grids ...\\SampleFiles\\converge.csv";
               MessageBox.Show(GridMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
               //tabControl1.TabPages.Add(tabPage4);
           }
           ///////////////////////////////////////////////////////////////////////////////////////////////







       }

       public System.Data.DataTable GetWorksheet(string worksheetName)
       {
  
           //robm MessageBox.Show(gbl.connectionString);
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

       private void button1_Click(object sender, EventArgs e)
       {
          double young, poisson = Convert.ToDouble(textBox5.Text);
          double[,] mat = new double[3,3];
          double[,] omat = new double[3, 3];
          double[,] inmat = new double[3,3];
          double[,] dg1data = new double[dataGridView1.Rows.Count, 3];
          double[] angle = new double[3];
          DateTime[] dg1date = new DateTime[dataGridView1.Rows.Count];
          double radians,tmp;
          int h, i, j, l, k, pivot;
          double rad_deg = 0.017453292519943;	/* radians / degree */
      ////////////////////////////////////////////////////////////////////
         char[] closure_station = new char[20];
         char[] looking = new char[20];
         double[] sta_dist = new double[3];			/* initial station distances */
         double[] strainin = new double[3];			/* input ea eb ec  */
         double epsilon1;
         double epsilon2;
         double alpha = 0;
         double sigma1, sigma2;
         double K;
         double[] strain = new double[3];		      	/* results, ex ey and gamma */
         double p, den;
      ////////////////////////////////////////////////////////////////////

           
           l = 0;
           foreach (DataGridViewRow dgvr in dataGridView1.Rows)
           {
               if (isDate(dgvr.Cells[0].Value.ToString()))
               {
                   //dgvr.DefaultCellStyle.ForeColor = Color.Red;
                   Debug.Print(dgvr.Cells[1].Value.ToString() + "," + dgvr.Cells[2].Value.ToString() + "," + dgvr.Cells[3].Value.ToString());
                   dg1date[l] = Convert.ToDateTime(dgvr.Cells[0].Value.ToString());
                   dg1data[l, 0] = Convert.ToDouble(dgvr.Cells[1].Value.ToString());
                   dg1data[l, 1] = Convert.ToDouble(dgvr.Cells[2].Value.ToString());
                   dg1data[l, 2] = Convert.ToDouble(dgvr.Cells[3].Value.ToString());
                   l++;
               }
               if (dgvr.Cells[0].Value.ToString() == "" && Convert.ToDouble(dgvr.Cells[1].Value.ToString()) > 1.0)
               {
                   angle[0] = Convert.ToDouble(dgvr.Cells[1].Value.ToString());
                   angle[1] = Convert.ToDouble(dgvr.Cells[2].Value.ToString());
                   angle[2] = Convert.ToDouble(dgvr.Cells[3].Value.ToString());
               }

           }
           Debug.Print(this.textBox4.ToString().Replace("System.Windows.Forms.TextBox, Text: ", "") + "," + this.textBox5.ToString().Replace("System.Windows.Forms.TextBox, Text: ", ""));
           young = Convert.ToDouble(this.textBox4.ToString().Replace("System.Windows.Forms.TextBox, Text: ", ""));
           poisson = Convert.ToDouble(this.textBox5.ToString().Replace("System.Windows.Forms.TextBox, Text: ", ""));
           double shear_modulus = 500.00 * young / (poisson + 1.0);
           
/////////////////////////////////////////////////////////////////////////
  for ( i=0 ; i<3 ; i++ )
  {
	 radians = rad_deg * angle[i];
	 tmp = Math.Cos( radians );
	 mat[i,0] = tmp * tmp ;
	 tmp = Math.Sin( radians );
	 mat[i,1] = tmp * tmp ;
	 mat[i,2] = 0.5 * Math.Sin( 2 * radians ) ;
  }
  for (i = 0; i < 3; i++)
  {
      radians = rad_deg * angle[i];
      tmp = Math.Cos(radians);
      omat[i, 0] = tmp * tmp;
      tmp = Math.Sin(radians);
      omat[i, 1] = tmp * tmp;
      omat[i, 2] = 0.5 * Math.Sin(2 * radians);
  }

/* Initialize the future inverse matrix "inmat" to the identity matrix */

  for ( i=0 ; i<3 ; i++ )
  {
	 for ( j=0 ; j<3 ; j++ ) inmat[i,j] = 0 ;
	 inmat[i,i] = 1 ;
  }

/* Make the inverse matrix .  First do something like a pivot,
   to avoid getting a zero on the diagonal.  */
    pivot = 0;
    if ( Math.Abs( mat[0,0] ) < .000001) pivot = 1;
    else
    {
	   if ( Math.Abs( mat[0,1] ) > .000001)
	   {
	    tmp = mat[1,0] / mat[0,0];
	    if ( Math.Abs(tmp - mat[1,1]/mat[0,1]) < .000001) pivot =1;
     }
    }
    if ( pivot == 1 )
    {
     for ( j=0 ; j<3 ; j++ ) mat[0,j] += mat[2,j];
    }

/* Now the matrix inversion */

    for ( k=0 ; k<3 ; k++ )
    {
	   if (( tmp =  mat[k,k] ) == 0 )
	    {
		   Debug.Print("\n****divide by zero\a\a" );
           break;
	    }
	    for ( j=0 ; j<3 ; j++ )
	    {
            mat[k,j] /= tmp;
            inmat[k,j] /= tmp;
        }
	    for ( i=0 ; i<3 ; i++ )
	    {
	     if ( i != k )
	     {
		    tmp = mat[i,k];
		    for ( j=0 ; j<3 ; j++ )
		    {
		     mat[i,j]   = mat[i,j] -  mat[k,j] * tmp;
		     inmat[i,j] = inmat[i,j] - inmat[k,j] * tmp;
	    	}
	     }
	    }
    }

/* Now compensate for the  sort of pivot. */
    if ( pivot == 1 )
    {
	   for ( i=0 ; i<3 ; i++ ) inmat[i,2] += inmat[i,0];
    }
  
    for (i=0; i<3; i++)
    {
     for(j=0;j<3;j++)
     {
        omat[i, j]  = RoundToSigFigs(Convert.ToDouble(omat[i, j].ToString("#,0.00000")), 3);
        inmat[i, j] = RoundToSigFigs(Convert.ToDouble(inmat[i, j].ToString("#,0.00000")), 3);
     }
    }

    DataTable dt2 = new DataTable();
    dt2.Columns.Add(new DataColumn("MTXColumn1", typeof(Double)));
    dt2.Columns.Add(new DataColumn("MTXColumn2", typeof(Double)));
    dt2.Columns.Add(new DataColumn("MTXColumn3", typeof(Double)));
    DataTable dt3 = new DataTable();
    dt3.Columns.Add(new DataColumn("INVColumn1", typeof(Double)));
    dt3.Columns.Add(new DataColumn("INVColumn2", typeof(Double)));
    dt3.Columns.Add(new DataColumn("INVColumn3", typeof(Double)));
    DataTable dt4 = new DataTable();
    dt4.Columns.Add(new DataColumn("Date", typeof(DateTime)));
    dt4.Columns.Add(new DataColumn("ea", typeof(Double)));
    dt4.Columns.Add(new DataColumn("eb", typeof(Double)));
    dt4.Columns.Add(new DataColumn("ec", typeof(Double)));
    dt4.Columns.Add(new DataColumn("e1", typeof(Double)));
    dt4.Columns.Add(new DataColumn("e2", typeof(Double)));
    dt4.Columns.Add(new DataColumn("alpha", typeof(Double)));
    dt4.Columns.Add(new DataColumn("sigma1", typeof(Double)));
    dt4.Columns.Add(new DataColumn("sigma2", typeof(Double)));
    dt4.Columns.Add(new DataColumn("circle", typeof(System.Drawing.Bitmap)));
    DataRow dr ;


    for (i = 0; i < 3; i++)
    {
     dr = dt2.NewRow();
     dr["MTXColumn1"] = omat[i, 0];
     dr["MTXColumn2"] = omat[i, 1];
     dr["MTXColumn3"] = omat[i, 2];
     dt2.Rows.Add(dr);
    }
    dataGridView2.DataSource = dt2;

    for (i = 0; i < 3; i++)
    {
     dr = dt3.NewRow();
     dr["INVColumn1"] = inmat[i, 0];
     dr["INVColumn2"] = inmat[i, 1];
     dr["INVColumn3"] = inmat[i, 2];
     dt3.Rows.Add(dr);
    }

    dataGridView3.DataSource = dt3;
    tabControl1.SelectedIndex = 2;

/////////////////////////////////////////////////////////////////////////
// Determine epsilon1 and epsilon2 using;
//epsilon x  =strain[0]
//epsilon y  =strain[1]
//gamma xy   =strain[2] 
/////////////////////////////////////////////////////////////////////////
//epsilon1
/////////////////////////////////////////////////////////////////////////
    for (i = 0; i < 3; i++) { strain[i] = 0; }

    Debug.Print("----------------------------------------------------------------------------------");
    for (h = 0; h < dataGridView1.Rows.Count; h++)
    {
        strainin[0] = dg1data[h, 0];
        strainin[1] = dg1data[h, 1];
        strainin[2] = dg1data[h, 2];
        for (i = 0; i < 3; i++)
        {
            for (j = 0; j < 3; j++)
            {
                strain[i] += .000001 * (strainin[j] * inmat[i, j]);
            }
        }


        tmp = strain[0] - strain[1];
        tmp = Math.Sqrt(strain[2] * strain[2] + tmp * tmp);
        epsilon1 = 0.5 * (strain[0] + strain[1] + tmp);
        epsilon2 = 0.5 * (strain[0] + strain[1] - tmp);



/////////////////////////////////////////////////////////////////////////
//alpha
/////////////////////////////////////////////////////////////////////////

	   tmp = strain[0] - strain[1];
	   if ( Math.Abs( tmp ) > .0000001)
	   {
	     alpha = 0.5 * Math.Atan( strain[2] / tmp) / rad_deg ;
	   }
	   if ( tmp < 0.0 )		/* left side of plane */
	   {
	     if ( strain[2] < 0.0 )	/* third quadrant */
	     {
	       if (tmp > -.00000011 ) alpha = -45.0;
	       else alpha = alpha - 90.0;
	     }
	     else			/* second quadrant */
	     {
	       if (tmp > -.00000011 ) alpha = 45.0;
	       else alpha = alpha + 90.0;
	     }
	   }
	   else				/* right side of plane */
	   {
	     if ( strain[2] < 0.0 )	/* forth quadrant */
	     {
	       if ( tmp < .00000011 ) alpha = -45.0;
	     }
	     else			/* first quadrant */
	     {
	       if ( tmp < .00000011 ) alpha = 45.0;
	       if ( strain[2] == 0.0 ) alpha = 0.0;	/* special case */
	     }
	   }




/////////////////////////////////////////////////////////////////////////
// determine K, epsilon1 and epsilon2 
////////////////////////////////////////////////////////////////////////
       p = 4 * (1 - poisson) - 1;
       den = epsilon1 * (1 - p) - epsilon2 * (1 + p);
       K = (epsilon2 * (1 - p) - epsilon1 * (1 + p)) / den;
       sigma2 = (epsilon1 * 4 * shear_modulus) / ((1 + K) - p * (1 - K));
       sigma1 = K * sigma2;


       dr = dt4.NewRow();
       dr["date"] = dg1date[h];
       dr["ea"] = RoundToSigFigs(Convert.ToDouble(strainin[0].ToString("#,0.00000")), 3);
       dr["eb"] = RoundToSigFigs(Convert.ToDouble(strainin[1].ToString("#,0.00000")), 3);
       dr["ec"] = RoundToSigFigs(Convert.ToDouble(strainin[2].ToString("#,0.00000")), 3);
       dr["e1"] = RoundToSigFigs(Convert.ToDouble(((1000000.0) * epsilon1).ToString("#,0.00000")), 3);
       dr["e2"] = RoundToSigFigs(Convert.ToDouble(((1000000.0) * epsilon2).ToString("#,0.00000")), 3);
       dr["alpha"] = RoundToSigFigs(Convert.ToDouble(alpha.ToString("#,0.00000")), 3);
       dr["sigma1"] = RoundToSigFigs(Convert.ToDouble(sigma1.ToString("#,0.00000")), 3);
       dr["sigma2"] = RoundToSigFigs(Convert.ToDouble(sigma2.ToString("#,0.00000")), 3);

       Icon CircleIcon = new Icon(this.GetType(), "Circle.ico");
       DataGridViewImageColumn iconColumn = new DataGridViewImageColumn();
       iconColumn.Image = CircleIcon.ToBitmap();
       iconColumn.Name = "Circle";
       iconColumn.HeaderText = "Alpha";

    /////////////////////////////////////////////////////////////////////////////////////
       Graphics newGraphics = Graphics.FromImage(iconColumn.Image);

       // Alter image.
     //  newGraphics.FillRectangle(new SolidBrush(Color.Black), 100, 50, 100, 100);
    Pen redPen = new Pen(Color.Red, 1);
    Pen blackPen = new Pen(Color.Black, 1);
    Pen greenPen = new Pen(Color.Green, 1);
    Pen bluePen = new Pen(Color.Blue, 1);
    Pen whitePen = new Pen(Color.White, 1);
    double angletan = 0.0174532925 * ((double)alpha);
    double anglerad = Math.Tan((double)angletan);

    double x2 = Math.Abs(20.1/anglerad);
    double y2 = 20.5;
    //if ( (x2 > 0) && (x2 < 10) ){ newGraphics.DrawLine(redPen, (float)50.0, (float)50.0, Math.Abs((float)x2)*10, (float)50.5); }
    //Graphics.DrawLine(blackPen, x1, y1, x2, y2);
    //Y line
    newGraphics.DrawLine(whitePen, (float)-15.0, (float)15.0, (float)50.0, (float)15.0);
    //X line
    newGraphics.DrawLine(whitePen, (float)15.0, (float)-15.0, (float)15.0, (float)50.5);

    //Console.Write("angle:" + alpha + " x2:" + x2 + " y2:" + y2 + "\n");
    //Debug.Print("angle:" + alpha + " x2:" + x2 + " y2:" + y2);

    //newGraphics.DrawLine(redPen, (float)15.0, (float)15.0, (float)20.0, (float)20.0);
    if ( (x2 != 0) && (x2 < 1000) )
    {
        if (x2 < 20)
        {
            x2 = 2 * x2;
            y2 = 2 * y2;
        }

        if (alpha >= 0 && alpha <= 90) 
        {
            x2 = (double)15.0 + x2;
            y2 = (double)15 - y2;
            newGraphics.DrawLine(redPen, (float)15.0, (float)15.0, (float)x2, (float)y2);
            Console.Write("angle:" + alpha + " x2:" + x2 + " y2:" + y2 + "\n");
        }
        if (alpha >= 91 && alpha <= 180) 
        {
            x2 = (double)15.0 + x2;
            y2 = (double)15 - y2;
            newGraphics.DrawLine(redPen, (float)15.0, (float)15.0, (float)x2, (float)y2);
            Console.Write("angle:" + alpha + " x2:" + x2 + " y2:" + y2 + "\n");
        }
        if (alpha <= -1 && alpha >= -90) 
        {
            x2 = (double)15.0 + x2;
            y2 = (double)15 + y2;
            newGraphics.DrawLine(redPen, (float)15.0, (float)15.0, (float)x2, (float)y2);
            Console.Write("angle:" + alpha + " x2:" + x2 + " y2:" + y2 + "\n");
        }
        if (alpha <= -91 && alpha >= -179) 
        {
            y2 = 20.5;
            newGraphics.DrawLine(redPen, (float)15.0, (float)15.0, -1 * (float)x2, (float)y2);
            Console.Write("angle:" + alpha + " x2:" + x2 + " y2:" + y2 + "\n");
        }
    }

    /////////////////////////////////////////////////////////////////////////////////////

       //dr["circle"] = iconColumn.Image;
       dr["circle"] = iconColumn.Image;
       
       dt4.Rows.Add(dr);
       

       Debug.Print("date:" + dg1date[l] + "s1:" + strainin[0].ToString() + " s2:" + strainin[1].ToString() + " s3:" + strainin[2].ToString() + " eps1:" + ((1000000.0) * epsilon1).ToString() + " eps2:" + ((1000000.0) * epsilon2).ToString() + " alpha:" + alpha.ToString() + " sig1:" + sigma1.ToString() + " sig2:" + sigma2.ToString());
  
       strain[0] = 0;
       strain[1] = 0;
       strain[2] = 0;
    }
    dataGridView4.DataSource = dt4;
/////////////////////////////////////////////////////////////////////////
       }
       public bool isDate(string value)
       {
           DateTime date;
           return DateTime.TryParse(value, out date);
       }

       public double RoundToSigFigs(double value, int sigFigures)
       {
           // this method will return a rounded double value at a number of signifigant figures.
           // the sigFigures parameter must be between 0 and 15, exclusive.

           int roundingPosition;   // The rounding position of the value at a number of sig figures.
           double scale;           // Optionally used scaling value, for rounding whole numbers or decimals past 15 places

           // handle exceptional cases
           if (value == 0.0d) { return value; }
           if (double.IsNaN(value)) { return 0.0d; } //{ return double.NaN; }
           if (double.IsPositiveInfinity(value)) { return double.PositiveInfinity; }
           if (double.IsNegativeInfinity(value)) { return 0.0d; } //{ return double.NegativeInfinity; }
           if (sigFigures < 1 || sigFigures > 14) { throw new ArgumentOutOfRangeException("The sigFigures argument must be between 0 and 15 exclusive."); }


           string strValue = value.ToString().Replace("-","");
           int mantissaIdx = strValue.IndexOf(".");
           sigFigures = sigFigures + mantissaIdx;


           // The resulting rounding position will be negative for rounding at whole numbers, and positive for decimal places.
           if (value < 1 && value > -1 )
           {
           roundingPosition = sigFigures - 2 - (int)(Math.Floor(Math.Log10(Math.Abs(value))));
           }
           else
           {
           roundingPosition = sigFigures - 1 - (int)(Math.Floor(Math.Log10(Math.Abs(value))));
           }
           // try to use a rounding position directly, if no scale is needed.
           // this is because the scale mutliplication after the rounding can introduce error, although 
           // this only happens when you're dealing with really tiny numbers, i.e 9.9e-14.
           if (roundingPosition > 0 && roundingPosition < 15)
           {
               return Math.Round(value, roundingPosition, MidpointRounding.AwayFromZero);
           }
           else
           {
               scale = Math.Pow(10, Math.Ceiling(Math.Log10(Math.Abs(value))));
               return Math.Round(value / scale, sigFigures, MidpointRounding.AwayFromZero) * scale;
           }
       }

       private void button3_Click(object sender, EventArgs e)
       {
       string outFile = gbl.fileName.ToString();
       outFile = Path.GetFileName(Path.ChangeExtension(outFile,"csv"));
       outFile = outFile.ToLower();


       SaveFileDialog saveFileDialog1 = new SaveFileDialog();
       saveFileDialog1.Filter = "CSV|*.csv";
       saveFileDialog1.Title = "Save an CSV File";
       saveFileDialog1.ShowDialog();



        //using (StreamWriter MyFile = new StreamWriter(Application.StartupPath + "\\output_" + outFile))
       using (StreamWriter MyFile = new StreamWriter(saveFileDialog1.FileName))
        {
         MyFile.WriteLine("date,ea,eb,ec,e1,e2,alpha,sigma1,sigma2");
         for (int i = 0; i < dataGridView4.RowCount - 1; i++)
         {
          string stLine = "";
          for (int j = 0; j < 9; j++)
          {
           //stLine = stLine + grdResults.Rows.Item(i).Cells.Item(j).Value + ",";
              Debug.Print("i:" + i + " j:" + j + " data:" + dataGridView1.Rows[i].Cells[j].FormattedValue.ToString());
              stLine = stLine + dataGridView4.Rows[i].Cells[j].FormattedValue.ToString() + ",";
          }
          MyFile.WriteLine(stLine);
         }
         MessageBox.Show("Information of all the files are successfully saved in the following location: \n" + saveFileDialog1.FileName, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
////////////////////////////////////////////////////////////////
System.Drawing.Graphics GraphicsObject = Graphics.FromHwnd(this.Handle);
GraphicsObject.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
System.Drawing.Pen redpen = new System.Drawing.Pen(Color.Red, 1);
GraphicsObject.PageUnit = GraphicsUnit.Pixel;
GraphicsObject.DrawLine(redpen, new Point(0, 0), new Point(50, 50));

////////////////////////////////////////////////////////////////
       }
      }
    }
}




