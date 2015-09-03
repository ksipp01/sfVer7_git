using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using SqlCeBulkCopy = ErikEJ.SqlCe.SqlCeBulkCopy;
using OleDbDataReader = System.Data.IDataReader;
using System.Data.SqlServerCe;
using DataTable = System.Data.DataTable;
using System.Windows.Forms;
using System.Drawing;
using System.Data.SqlClient;
using System.Data.Common;
using ErikEJ.SqlCe;




namespace Pololu.Usc.ScopeFocus
{
    class Data
    {
        public DataGridView dataGridView1 = new DataGridView();
        private static int _rows;
        //private static int _apexHFR;
        //private static int _enteredPID;
        //private static double _enteredSlopeDWN;
        //private static double _enteredSlopeUP;
        //private static string _equip;
        //private static int _pID;
        //private static string _importPath;
        //private static int _intersectPos;

        //private static int _hFRarraymin;
        //private static int _posminHFR;
        //private static float _slopeHFRup;
        //private static float _slopeHFRdwn;
        //private int rows;
        //public static int intersectPos
        //{
        //    get { return _intersectPos; }
        //    set { _intersectPos = value; }
        //}


        //public static string ImportPath
        //{
        //    get { return _importPath; }
        //    set { _importPath = value; }
        //}
        //public static int PID
        //{
        //    get { return _pID; }
        //    set { _pID = value; }
        //}
        //public static int PosminHFR
        //{
        //    get { return _posminHFR; }
        //    set { _posminHFR = value; }
        //}
        //public static float SlopeHFRup
        //{
        //    get { return _slopeHFRup; }
        //    set { _slopeHFRup = value; }
        //}
        //public static float SlopeHFRdwn
        //{
        //    get { return _slopeHFRdwn; }
        //    set { _slopeHFRdwn = value; }
        //}
        //public static int HFRarraymin
        //{
        //    get { return _hFRarraymin; }
        //    set { _hFRarraymin = value; }
        //}

        //public static int ApexHFR
        //{
        //    get { return _apexHFR; }
        //    set { _apexHFR = value; }
        //}
        //public static int EnteredPID
        //{
        //    get { return _enteredPID; }
        //    set { _enteredPID = value; }
        //}
        //public static double EnteredSlopeDWN
        //{
        //    get { return _enteredSlopeDWN; }
        //    set { _enteredSlopeDWN = value; }
        //}
        //public static double EnteredSlopeUP
        //{
        //    get { return _enteredSlopeUP; }
        //    set { _enteredSlopeUP = value; }
        //}
        //public static string Eqiup
        //{
        //    get { return _equip; }
        //    set { _equip = value; }
        //}
        string conString = WindowsFormsApplication1.Properties.Settings.Default.MyDatabase_2ConnectionString;

        string _filename = "";
        public string Filename
        {
            get { return this._filename; }
            set { _filename = value; }
        }

        public void FillData()
        {
            try
            {
                using (SqlCeConnection con = new SqlCeConnection(conString))
                {
                    con.Open();
                    using (SqlCeDataAdapter a = new SqlCeDataAdapter("SELECT * FROM table1", con))
                    {
                       
                        
                        DataTable t = new DataTable();
                        a.Fill(t);
                        dataGridView1.DataSource = t;
                        a.Update(t);
                    }
                    con.Close();
                }
                MainWindow m = new MainWindow();
                ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + m.TS3 + "'";
              //  ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + m.toolStripStatusLabel3.Text.ToString() + "'";
                //   ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + textBox13.Text.ToString() + "'";
            }
            catch (Exception ex)
            {
                MainWindow m = new MainWindow();
              m.Log("FillData Error" + ex.ToString());
            }
        }





        //analyze SQL data  

        public void GetAvg()
        {
            try
            {
                using (SqlCeConnection con = new SqlCeConnection(conString))
                {
                    con.Open();
                    using (SqlCeCommand com = new SqlCeCommand("SELECT AVG(PID) FROM table1 WHERE Equip = @equip", con))
                    {
                        com.Parameters.AddWithValue("@equip", MainWindow.Eqiup);
                        SqlCeDataReader reader = com.ExecuteReader();
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0))
                            {
                                int numb5 = reader.GetInt32(0);
                               MainWindow.EnteredPID = numb5;
                            }
                        }
                        reader.Close();
                    }
                    using (SqlCeCommand com1 = new SqlCeCommand("SELECT AVG(SlopeDWN) FROM table1 WHERE Equip = @equip", con))
                    {
                        com1.Parameters.AddWithValue("@equip",MainWindow.Eqiup);
                        SqlCeDataReader reader1 = com1.ExecuteReader();
                        while (reader1.Read())
                        {
                            if (!reader1.IsDBNull(0))
                            {
                                MainWindow main = new MainWindow();
                               
                                double numb5 = reader1.GetDouble(0);
                                double numb5Rnd = Math.Round(numb5, 5);
                                main.TB3 = numb5Rnd.ToString();
                              // textBox3.Text = numb5Rnd.ToString();
                               MainWindow.EnteredSlopeDWN = numb5Rnd;
                            }

                        }
                        reader1.Close();
                    }

                    using (SqlCeCommand com2 = new SqlCeCommand("SELECT AVG(SlopeUP) FROM table1 WHERE Equip = @equip", con))
                    {
                        com2.Parameters.AddWithValue("@equip", MainWindow.Eqiup);
                        SqlCeDataReader reader2 = com2.ExecuteReader();
                        while (reader2.Read())
                        {
                            if (!reader2.IsDBNull(0))
                            {
                                double numb5 = reader2.GetDouble(0);
                                double numb5Rnd = Math.Round(numb5, 5);
                                MainWindow m = new MainWindow();
                                m.TB10 = numb5Rnd.ToString();
                               // textBox10.Text = numb5Rnd.ToString();
                               MainWindow.EnteredSlopeUP = numb5Rnd;
                            }
                        }
                        reader2.Close();
                    }
                    using (SqlCeCommand com3 = new SqlCeCommand("SELECT AVG(BestHFR) FROM table1 WHERE Equip = @equip", con))
                    {
                        com3.Parameters.AddWithValue("@equip", MainWindow.Eqiup);
                        SqlCeDataReader reader3 = com3.ExecuteReader();
                        while (reader3.Read())
                        {
                            if (!reader3.IsDBNull(0))
                            {
                                int numb6 = reader3.GetInt32(0);
                                MainWindow m = new MainWindow();
                                m.TB15 = numb6.ToString();
                               // textBox15.Text = numb6.ToString();
                            }
                        }
                        reader3.Close();
                    }
                    /*
                    //   rem'd ****moved to mainWindow load so it puts focus position value in at startup
                    using (SqlCeCommand com4 = new SqlCeCommand("SELECT AVG(FocusPos) FROM table1 WHERE Equip = @equip", con))
                    {
                        com4.Parameters.AddWithValue("@equip", equip);
                        SqlCeDataReader reader4 = com4.ExecuteReader();
                        while (reader4.Read())
                        {
                            if (!reader4.IsDBNull(0))
                            {
                                int numb7 = reader4.GetInt32(0);
                              //     textBox4.Text = numb7.ToString();******88888rem'd 4-10
                            }
                        }
                        reader4.Close();
                    }

*/
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MainWindow m = new MainWindow();
                m.Log("GetAvg Error" + ex.ToString());
            }
        }
        public void WriteSQLdata()
        {
            try
            {
                using (SqlCeConnection con = new SqlCeConnection(conString))
                {

                    con.Open();
                    int num =MainWindow.HFRarraymin;
                    int num2 =MainWindow.apexHFR;
                    int num4 = MainWindow.PosminHFR;
                    float up = (float)MainWindow.SlopeHFRup;
                    float down = (float)MainWindow.SlopeHFRdwn;
                    //row numbering  adds 1 to max value, allows for deletion of rows without number dulpication
                    //can always modify or re-number in excel then import
                    using (SqlCeCommand com1 = new SqlCeCommand("SELECT MAX (Number) FROM table1", con))
                    {
                        SqlCeDataReader reader = com1.ExecuteReader();
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0))
                            {
                                _rows = 0;
                            }
                            else
                            {
                                _rows = reader.GetInt32(0);
                            }
                        }
                    }

                    using (SqlCeCommand com = new SqlCeCommand("INSERT INTO table1 (Date, PID, SlopeDWN, SlopeUP, Number, Equip, BestHFR, FocusPos) VALUES (@Date, @PID, @SlopeDWN, @SlopeUP, @Number, @equip, @BestHFR, @FocusPos)", con))
                    {

                        com.Parameters.AddWithValue("@Date", DateTime.Now);
                        com.Parameters.AddWithValue("@PID",MainWindow.PID);
                        com.Parameters.AddWithValue("@SlopeDWN", down);
                        com.Parameters.AddWithValue("@SlopeUP", up);
                        com.Parameters.AddWithValue("@Number", _rows + 1);
                        com.Parameters.AddWithValue("@equip",MainWindow.Eqiup);
                        com.Parameters.AddWithValue("@BestHFR",MainWindow.HFRarraymin);
                        com.Parameters.AddWithValue("@FocusPos",MainWindow.intersectPos);
                        com.ExecuteNonQuery();
                        _rows++;
                    }
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MainWindow m = new MainWindow();
               m.Log("WriteSQLData Error" + ex.ToString());
            }
        }
        public void importDataFromExcel()
        {

            string sSQLTable = "table1";
            //this works
            string myExcelDataQuery = "SELECT Date, PID, SlopeDWN, SlopeUP, Number, Equip, BestHFR, FocusPos FROM [Sheet1$]";
            try
            {
                if (MainWindow.ImportPath == null)
                {

                    MainWindow m = new MainWindow();
                 //   DialogResult result = openFileDialog1.ShowDialog();
                //    ImportPath = openFileDialog1.FileName.ToString();
                  //  ImportPath = Filename;
                    m.TB34 = Filename;
                   // textBox34.Text = ImportPath.ToString();
                }
                string sSqlConnectionString = conString;
                string sExcelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +MainWindow.ImportPath + @";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""";
                //  string sExcelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\Users\kevin\Documents\scopefocusData.xls;Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""";

                string sClearSQL = "DELETE FROM " + sSQLTable;

                SqlCeConnection SqlConn = new SqlCeConnection(conString);
                SqlCeCommand SqlCmd = new SqlCeCommand(sClearSQL, SqlConn);
                SqlConn.Open();
                SqlCmd.ExecuteNonQuery();
                SqlConn.Close();
                OleDbConnection OleDbConn = new OleDbConnection(sExcelConnectionString);
                OleDbCommand OleDbCmd = new OleDbCommand(myExcelDataQuery, OleDbConn);
                OleDbConn.Open();
                OleDbDataReader dr = OleDbCmd.ExecuteReader();
                using (SqlCeBulkCopy bc = new SqlCeBulkCopy(conString))
                {
                    bc.DestinationTableName = "table1";
                    bc.WriteToServer(dr);
                }

                OleDbConn.Close();
                MessageBox.Show("Data Import Successful", "scopefocus");
            }
            catch
            {
                MessageBox.Show("Import Failed", "scopefocus");
            }
        }


        public void Update()
        {
            try
            {
                MainWindow m = new MainWindow();
                if (MainWindow.Eqiup == null)
                {
                    MessageBox.Show("Must select equipment first", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
             //   Data d = new Data();
                GetAvg();
                FillData();

                //try adding std dev and display in textbox16
                // Std Dev UP
                if (this.dataGridView1.RowCount > 2)
                {
                    List<double> StdDevCalcUP = new List<double>();
                    for (int i = 0; i < (dataGridView1.Rows.Count - 1); i++)
                    {
                        StdDevCalcUP.Add(Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value));
                    }

                    double standardDeviationUP = m.CalculateSD(StdDevCalcUP);
                    double sdUP = Math.Round(standardDeviationUP, 5);
                    m.TB16 = sdUP.ToString();
                }
                /*  make sure there is enough data   *****remd 6-28.  not needed annoying  
                if (dataGridView1.RowCount <= 2)
                {
                    DialogResult result;
                    result = MessageBox.Show("Need More Data for Calculation", "scopefocus",
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.OK)
                    {
                        return;
                    }
                }
                */

                //Std Dev DWN
                if (this.dataGridView1.RowCount > 2)
                {
                    List<double> StdDevCalcDWN = new List<double>();
                    for (int i = 0; i < (dataGridView1.Rows.Count - 1); i++)
                    {
                        StdDevCalcDWN.Add(Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value));
                    }

                    double standardDeviationDWN = m.CalculateSD(StdDevCalcDWN);
                    double sdDWN = Math.Round(standardDeviationDWN, 5);
                    m.TB14 = sdDWN.ToString();
                }
                /*    ****remd 6-29 not needed
                if (dataGridView1.RowCount <= 2)
                {
                    DialogResult result;
                    result = MessageBox.Show("Need More Data for Calculation", "scopefocus",
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.OK)
                    {
                        return;
                    }
                }
                 */
            }
            catch (Exception ex)
            {
              //  Log("Update Error - Line 1817" + ex.ToString());
            }
        }


    }

}
