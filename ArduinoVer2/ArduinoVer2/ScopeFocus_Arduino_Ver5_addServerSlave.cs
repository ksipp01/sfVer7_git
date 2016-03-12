//Arduino version
//added standby to abort v-curve 9-23-11
//this is latest as of 10-12-11
//10-13-11 added gets Max form filename.  still need to use if 2 minHFRs are equally low
//1-7-12 fixed fine v repeat reset
//*******need to change default path in setting to c:windows\temp  prior to publish.  
//******this is temporary just for installation and can be changed on first use 
//Ver4_filter_metricHFR had the monitor PHD dx dy background stuff..that was omitted 
//5-17-12 added different socket close method all rem'd w/ date
//5-23-12 flat for every filter bin matches that fitler
//5-23-12 added abililty to connect arduino after form_load.  also noted default port setting not being saved or used. 

//***************************this slave version uses scopefocus slave buttons to control the 2nd neb
//***********Lite version pending......

using System;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Linq;
using System.Data;
using System.IO.Ports;
using System.Data.SqlServerCe;
using System.Configuration;
using WindowsFormsApplication1.Properties;
using System.Net.Sockets;
using System.Media;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Point = System.Drawing.Point;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.Common;
using ErikEJ.SqlCe;
using SqlCeBulkCopy = ErikEJ.SqlCe.SqlCeBulkCopy;
using OleDbDataReader = System.Data.IDataReader;
using System.Net.Mail;
using System.Diagnostics;








namespace Pololu.Usc.ScopeFocus
{
    public partial class MainWindow : Form
    {

        TcpClient clientSocket2 = new TcpClient();//added for neb slave
        TcpClient clientSocket = new TcpClient();//added for neb communication
        TcpClient phdsocket = new TcpClient();//test for phd socket
        NetworkStream serverStream;
        SerialPort port;
        SerialPort port2;
        
        public MainWindow()
        {
            InitializeComponent();
            foreach
                (var page in tabControl1.TabPages.Cast<TabPage>())
            {
                page.CausesValidation = true;
                page.Validating += new CancelEventHandler(OnTabPageValidating);
            }
            timer1.Start();
            
            string[] portlist = SerialPort.GetPortNames();
            foreach (String s in portlist)
            {
                comboBox1.Items.Add(s);
                comboBox6.Items.Add(s);
            }
            if (comboBox1.Items.Count == 1)
            {
              //  portautoselect = true;***** rem'd 5-24 for slave

            }

            fileSystemWatcher1.EnableRaisingEvents = false;
            fileSystemWatcher2.EnableRaisingEvents = false;
            fileSystemWatcher3.EnableRaisingEvents = false;
            fileSystemWatcher4.EnableRaisingEvents = false;
            fileSystemWatcher5.EnableRaisingEvents = false;//added to test metricHFR
            fileSystemWatcher6.EnableRaisingEvents = true;
            textBox11.Focus();

        }
        //*******************remove this if the second form FindCenter is not used ***********
        //try adding to make this avail in antoher form;
        public class Variables
        {
            private static int NebhWnd;

            public static int NebhWndl
            {
                get { return NebhWnd; }
                set { NebhWnd = value; }

            }
        }
        
            int SocketPort = 4301;
            string ScriptName = "\\listenPort.neb";
            //begin handle stuff
            private const int WM_COMMAND = 0x0112;
            private const int WM_CLOSE = 0xF060;
            private const int BN_CLICKED = 245;
            private const int GWL_ID = -12;

            private const int WM_RBUTTONDOWN = 0x204; //Right mousebutton down
            private const int WM_RBUTTONUP = 0x205;   //Right mousebutton up
            private const int WM_LBUTTONDOWN = 0x201; //Left mousebutton down
            private const int WM_LBUTTONUP = 0x202;  //Left mousebutton up
            //  private const int WM_GETTEXT = 0x0D;

            public const int WM_SETTEXT = 0x000C;
            public const int WM_GETTEXT = 0x000D;
            public const int WM_GETTEXTLENGTH = 0x000E;
            public const int WM_USER = 0x00400;

            //added to try SB_GET stuff
            public const int SB_GETTEXT = WM_USER + 13;
            public const int SB_GETTEXTLENGTH = WM_USER + 12;
            public const int SB_GETPARTS = WM_USER + 6;
            public const int SB_SETPARTS = WM_USER + 4;
            //********************change this so just in find handle method and remove textbox37**************
            //   string ScriptName = "\\listenPort.neb";
            //   string send; //temp to display script location
            string NebVersion;
            int NebVNumber;
            string NebPath;
            string ImportPath;

            public static string ExcelFilename;
            //en
            //   int hWnd;  goes w/ enumerator at botton
            bool setupWindowFound = false;
            int Advancedhwnd;
            int NebSlavehwnd;
            int Advhwnd;
            bool FlatsOn = false;
            int EnteredPID;
            double EnteredSlopeUP;
            double EnteredSlopeDWN;
            int editfound = 0; //counts "edit" handles for focus duration...need second one
            int hwndDuration;
            int hwndChild;
            int hwndChild1;
            int hwndChild2;
            bool FilterFocusOn = false;
            float FocusTime;
            bool startup = true;//used to ensure tab change only changes populates focuspos on startup
            int CaptureBin;
            bool FocusLocObtained = false;
            bool TargetLocObtained = false;
            int Aborthwnd;
            int Framehwnd;
            int Finehwnd;
            string NebCamera;
            public int NebhWnd;
            int PHDhwnd;
            int LoadScripthwnd;

            int panelhwnd;

            bool MountMoving = false;
            string GotoDoneCommand;
            string FocusLoc = "";
            string TargetLoc = "";

            string Filtertext;
            int metricHFR;
            int metricN = 0;
            int currentmetricN = 0;
            int[] AvgMetricHFR = null;
            bool MetricSample = false;
            int AvgMetric = 0;
            int testMetricHFR = 0;
            bool DarksOn = false;
            bool filtersynced = false;
            bool filterMoving = false;
            int CaptureTime;
            string Nebname;
            int filterCountCurrent = 0;
            int totalsubs;
            int filternumber = 0;
            int subsperfilter;
            int subCountCurrent = 0;
            int currentfilter = 0;
            int filter1used = 0;
            int filter2used = 0;
            int filter3used = 0;
            int filter4used = 0;
            int filter5used = 0;
            double BestPos;
            int selectedrowcount;
            string selectedcell;
            bool roughvdone = false;
            int arraycountright = 0;
            int arraycountleft = 0;
            int enteredMaxHFR;
            int enteredMinHFR;
            bool portautoselect = false;
            string path2;
            string portselected;
            string port2selected;
            int fineVrepeatDone;
            bool fineVrepeatOn = false;
            int fineVrepeat;
            string equip;
            string equipPrefix;
            int rows;
            bool GotoFocusOn = false;
            int intersectPos;
            double slopeHFRdwn = 0;
            double slopeHFRup = 0;
            double XintUP;
            double XintDWN;
            int PID;
            int HFRarraymin = 999;
            double maxarrayMax = 1;
            int apexHFR;
            int backlashN = 10;
            float backlashSum = 0;
            float backlashAvg = 0;
            int backlashCount;
            bool backlashDetermOn;
            bool backlashOUT;
            int backlashPosIN;
            int backlashPosOUT;
            float backlash = 0;
            double maxMax = 1;
            int posmaxMax;
            //abitrarily set min = 999 to start with
            int total = 0;
            int count = 0;
            int min = 999;
            int posMin = 0;
            int repeatProgress = 0;
            int repeatDone = 0;
            int vProgress = 0;//progess of one individual Vcurve
            int vDone = 0;
            int arraysize1;
            int arraysize2;
            int sum = 0;
            int avg = 0;
            long sumMax = 0;
            long avgMax = 0;
            int ffenable = 0;
            int repeatTotal = 0;
            int[] list = null;
            double[] listMax = null;
            double[] abc = null;
            int[] peat = null;
            double[] peatMax = null;
            int[] minHFRpos = null;
            double[] maxMaxPos = null;
            int posminHFR;
            int tempon = 0;
            int tempsum = 0;
            int tempavg = 0;
            int vv = 0;//total aquisitions for a given v-curve, N * repeat
            int travel;//for tsa-120 direct rack and pinion, quarter step
            int connect = 0;
            int connect2 = 0;
            int temp = 100;
            int portopen = 0;
            int vcurveenable = 0;

            int closing = 0;
            int syncval;
            int templog;
            int MoveDelay; //helps ensure no focus movement during capture
            int roundto;
            int vN;
            string conString = WindowsFormsApplication1.Properties.Settings.Default.MyDatabase_2ConnectionString;
            int FocusPerSub;
            int FocusPerSubCurrent = 0;
            int FilterFocusGroupCurrent;

            double CaptureTime3;

        
        void setarraysize()
        {
            arraysize1 = (int)numericUpDown5.Value;//course v N
            arraysize2 = (int)numericUpDown3.Value;//fine v N
        }

        void OnTabPageValidating(object sender, CancelEventArgs e)
        {

        }

        void PortOpen()
        {
            try
            {
                if (port != null)
                {
                    return;
                   // port.Close();*****rem'd 5-24 for slave
                  //  port.Dispose();
                }
                
                    port = new SerialPort(comboBox1.SelectedItem.ToString(), 9600, Parity.None, 8, StopBits.One);
                    port.Open();
                    watchforOpenPort();
                
                if (portopen == 1)
                {
                    Log("Connected to Arduino on " + comboBox1.SelectedItem.ToString());
                    portselected = comboBox1.SelectedItem.ToString();
                    button8.BackColor = System.Drawing.Color.Lime;
                    this.button8.Text = "Connected";
                }
                comboBox6.Items.Remove(portselected);
                // *****try add computer connect mode arduino 4-24
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
                Thread.Sleep(10);
                port.Write("C");
                Thread.Sleep(50);
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
            }
            catch(Exception ex)
            {
                Log("PortOpen Error" + ex.ToString());
                Send("PortOpen Error" + ex.ToString());
                FileLog("PortOpen Error" + ex.ToString());

            }
        }


        void Port2Open()
        {
            try
            {
                if (port2 != null)
                {
                    port2.Close();
                    port2.Dispose();
                }

                port2 = new SerialPort(comboBox6.SelectedItem.ToString(), 9600, Parity.None, 8, StopBits.One);
                port2.Open();
                port2selected = comboBox6.SelectedItem.ToString();
            }
            catch(Exception ex) 
            {
                Log("Port2Open Error" + ex.ToString());
                Send("Port2Open Error" + ex.ToString());
                FileLog("Port2Open Error" + ex.ToString());

            }
        }

        void watchforOpenPort()
        {
            try
            {
                if (port == null)
                {
                    portopen = 0;
                    DialogResult result2 = MessageBox.Show("Arduino Not Connected", "Arduino scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    portopen = 1;
                }
            }
            catch(Exception ex) 
            {
                Log("WatchOpenPort Error" + ex.ToString());
                Send("WatchOpenPort Error" + ex.ToString());
                FileLog("WatchOpenPort Error" + ex.ToString());

            }
        }

        void positionbar()
        {
            progressBar2.Maximum = travel;
            progressBar2.Minimum = 0;
            progressBar2.Increment(10);
            progressBar2.Value = count;
        }
        #region logging
        public void Log(Exception e)
        {
            Log(e.Message);
        }

        public void Log(string text)
        {
            
            if (LogTextBox.Text != "")
                LogTextBox.Text += Environment.NewLine;
            LogTextBox.Text += DateTime.Now.ToLongTimeString() + "  " + text;
            LogTextBox.SelectionStart = LogTextBox.Text.Length;
            LogTextBox.ScrollToCaret();

        }
        #endregion

        void FillData()
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
                ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";
                //   ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + textBox13.Text.ToString() + "'";
            }
            catch(Exception ex) 
            {
                Log("FillData Error" + ex.ToString());
            }
            }





        //analyze SQL data  

        void GetAvg()
        {
            try
            {
                using (SqlCeConnection con = new SqlCeConnection(conString))
                {
                    con.Open();
                    using (SqlCeCommand com = new SqlCeCommand("SELECT AVG(PID) FROM table1 WHERE Equip = @equip", con))
                    {
                        com.Parameters.AddWithValue("@equip", equip);
                        SqlCeDataReader reader = com.ExecuteReader();
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0))
                            {
                                int numb5 = reader.GetInt32(0);
                                EnteredPID = numb5;
                            }
                        }
                        reader.Close();
                    }
                    using (SqlCeCommand com1 = new SqlCeCommand("SELECT AVG(SlopeDWN) FROM table1 WHERE Equip = @equip", con))
                    {
                        com1.Parameters.AddWithValue("@equip", equip);
                        SqlCeDataReader reader1 = com1.ExecuteReader();
                        while (reader1.Read())
                        {
                            if (!reader1.IsDBNull(0))
                            {
                                double numb5 = reader1.GetDouble(0);
                                double numb5Rnd = Math.Round(numb5, 5);
                                textBox3.Text = numb5Rnd.ToString();
                                EnteredSlopeDWN = numb5Rnd;
                            }

                        }
                        reader1.Close();
                    }

                    using (SqlCeCommand com2 = new SqlCeCommand("SELECT AVG(SlopeUP) FROM table1 WHERE Equip = @equip", con))
                    {
                        com2.Parameters.AddWithValue("@equip", equip);
                        SqlCeDataReader reader2 = com2.ExecuteReader();
                        while (reader2.Read())
                        {
                            if (!reader2.IsDBNull(0))
                            {
                                double numb5 = reader2.GetDouble(0);
                                double numb5Rnd = Math.Round(numb5, 5);
                                textBox10.Text = numb5Rnd.ToString();
                                EnteredSlopeUP = numb5Rnd;
                            }
                        }
                        reader2.Close();
                    }
                    using (SqlCeCommand com3 = new SqlCeCommand("SELECT AVG(BestHFR) FROM table1 WHERE Equip = @equip", con))
                    {
                        com3.Parameters.AddWithValue("@equip", equip);
                        SqlCeDataReader reader3 = com3.ExecuteReader();
                        while (reader3.Read())
                        {
                            if (!reader3.IsDBNull(0))
                            {
                                int numb6 = reader3.GetInt32(0);
                                textBox15.Text = numb6.ToString();
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
            catch(Exception ex)
            {
                Log("GetAvg Error" + ex.ToString());
            }
        }
        private void WriteSQLdata()
        {
            try
            {
                using (SqlCeConnection con = new SqlCeConnection(conString))
                {

                    con.Open();
                    int num = HFRarraymin;
                    int num2 = apexHFR;
                    int num4 = posminHFR;
                    float up = (float)slopeHFRup;
                    float down = (float)slopeHFRdwn;
                    //row numbering  adds 1 to max value, allows for deletion of rows without number dulpication
                    //can always modify or re-number in excel then import
                    using (SqlCeCommand com1 = new SqlCeCommand("SELECT MAX (Number) FROM table1", con))
                    {
                        SqlCeDataReader reader = com1.ExecuteReader();
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0))
                            {
                                rows = 0;
                            }
                            else
                            {
                                rows = reader.GetInt32(0);
                            }
                        }
                    }

                    using (SqlCeCommand com = new SqlCeCommand("INSERT INTO table1 (Date, PID, SlopeDWN, SlopeUP, Number, Equip, BestHFR, FocusPos) VALUES (@Date, @PID, @SlopeDWN, @SlopeUP, @Number, @equip, @BestHFR, @FocusPos)", con))
                    {

                        com.Parameters.AddWithValue("@Date", DateTime.Now);
                        com.Parameters.AddWithValue("@PID", PID);
                        com.Parameters.AddWithValue("@SlopeDWN", down);
                        com.Parameters.AddWithValue("@SlopeUP", up);
                        com.Parameters.AddWithValue("@Number", rows + 1);
                        com.Parameters.AddWithValue("@equip", equip);
                        com.Parameters.AddWithValue("@BestHFR", HFRarraymin);
                        com.Parameters.AddWithValue("@FocusPos", intersectPos);
                        com.ExecuteNonQuery();
                        rows++;
                    }
                    con.Close();
                }
            }
            catch(Exception ex) 
            {
                Log("WriteSQLData Error" + ex.ToString());
            }
        }


        private void fileSystemWatcher2_Changed(object sender, FileSystemEventArgs e)
        {

            vcurve();

        }
        bool FineFocusAbort = false;
        bool NebListenOn = false;
        //std dev, avg and gotofocus
       // bool focusing = false;
        private void fileSystemWatcher3_Changed(object sender, FileSystemEventArgs e)
        {

                textBox41.Refresh();
                textBox41.Clear();

                if (currentfilter == 1)
                    Filtertext = comboBox2.Text;
                if (currentfilter == 2)
                    Filtertext = comboBox3.Text;
                if (currentfilter == 3)
                    Filtertext = comboBox4.Text;
                if (currentfilter == 4)
                    Filtertext = comboBox5.Text;
                watchforOpenPort();
                if (portopen == 1)
                {
                    
                    // int nn = 10;
                    int current;
                    roundto = (int)numericUpDown1.Value;
                    int nn = (int)numericUpDown5.Value;
                    if (vProgress < nn)
                    {
                        int[] list = new int[nn];
                        if (checkBox22.Checked == true)
                        {
                            metricpath = Directory.GetFiles(path2.ToString(), "metric*.fit");
                            current = GetMetric(metricpath, roundto);
                        }
                        else
                        {
                            fileSystemWatcher3.Filter = "*.bmp";
                            string[] filePaths = Directory.GetFiles(path2.ToString(), "*.bmp");
                            //remd to use path2 settins    string[] filePaths = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.bmp");         
                            current = GetFileHFR(filePaths, roundto);
                            if (current > 500)
                            {
                                //***** 6-28  still seems to goto focus position twice after abort and before first metric.
                                FineFocusAbort = true;
                                Array.Clear(list, 0, arraysize1);
                                Log("Focus star lost, using full frame metric");
                                SetForegroundWindow(NebhWnd);
                                Thread.Sleep(1000);
                                PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                                Thread.Sleep(2000);
                                checkBox22.Checked = true;
                                fileSystemWatcher3.EnableRaisingEvents = false;
                                clientSocket.Client.Disconnect(true);
                               // gotoFocus();
                              //  return;
                                /*
                                NebListenStart(NebhWnd, SocketPort);
                                
                                Thread.Sleep(2000);
                                fileSystemWatcher3.Filter = "*.fit";
                                MetricCapture();
                                
                                while (working == true)
                                {
                                    WaitForSequenceDone("Sequnce done", NebhWnd);
                                    Log("waiting for seq done");
                                    Thread.Sleep(100);
                                }
                                working = true;
                                
                              //  Thread.Sleep(4000);
                                metricpath = Directory.GetFiles(path2.ToString(), "metric*.fit");
                                current = GetMetric(metricpath, roundto);
                                 */



                            }

                        }
                        if (FineFocusAbort == false)
                        {
                            list[vProgress] = current;
                            abc[vProgress] = current;
                            sum = sum + list[vProgress];
                            avg = sum / nn;
                            string stdev = Math.Round(GetStandardDeviation(abc), 2).ToString();
                            textBox9.Text = stdev.ToString();
                            //     textBox14.Text = avg.ToString();
                            Log("Goto Focus :\t  N " + (vProgress + 1).ToString() + "\tHFR" + abc[vProgress].ToString() + "\t  Avg " + avg.ToString() + "\t  StdDev " + stdev.ToString());
                            string strLogText = "Goto Focus" + "\t  " + abc[vProgress].ToString() + "\t  " + avg.ToString() + "\t" + (vProgress + 1).ToString() + "\t" + stdev.ToString();
                            string path = textBox11.Text.ToString();
                            string fullpath = path + @"\log.txt";
                            StreamWriter log;
                            if (!File.Exists(fullpath))
                            {
                                log = new StreamWriter(fullpath);
                            }
                            else
                            {
                                log = File.AppendText(fullpath);
                            }
                            if (vProgress == 0)
                            {
                                log.WriteLine(DateTime.Now);
                                log.WriteLine("Type" + "\tHFR" + "\tAvg" + "\tN" + "\tStdDev");
                            }
                            log.WriteLine(strLogText);
                            log.Close();
                            vProgress++;
                            if ((checkBox22.Checked == true) && (vProgress != nn))
                                MetricCapture();
                            toolStripStatusLabel1.Text = "Finding Focus " + vProgress.ToString();//**************added 2_29 for testing

                        }

                    }
                        if (vProgress == nn)
                        {
                            //?  may need FSW3 enalbing event = false here for std dev use
                            if (GotoFocusOn == true)
                            {

                                if (radioButton2.Checked == true)//use upslope
                                {
                                    FillData();
                                    GetAvg();

                                    //  EnteredPID = Convert.ToInt32(textBox12.Text);
                                    //    EnteredSlopeUP = Convert.ToDouble(textBox10.Text);
                                    //      textBox14.Text = avg.ToString();
                                    BestPos = count - (avg / EnteredSlopeUP) + (EnteredPID / 2);
                                    gotopos(Convert.ToInt16(BestPos));
                                    Thread.Sleep(1000);
                                    fileSystemWatcher3.EnableRaisingEvents = false;
                                    GotoFocusOn = false;
                                    Log("Goto Focus Position: " + Convert.ToInt16(BestPos).ToString());
                                    textBox4.Text = Convert.ToInt16(BestPos).ToString();

                                }
                                if (radioButton3.Checked == true)//use downslope
                                {
                                    FillData();
                                    GetAvg();

                                    //    EnteredPID = Convert.ToInt32(textBox12.Text);
                                    //    EnteredSlopeDWN = Convert.ToDouble(textBox3.Text);
                                    // textBox14.Text = avg.ToString();
                                    BestPos = count - (avg / EnteredSlopeDWN) - (EnteredPID / 2);
                                    gotopos(Convert.ToInt16(BestPos));
                                    Thread.Sleep(1000);
                                    fileSystemWatcher3.EnableRaisingEvents = false;
                                    GotoFocusOn = false;
                                    //   Log("Goto Focus Position" + Convert.ToInt16(BestPos).ToString());
                                    textBox4.Text = Convert.ToInt16(BestPos).ToString();
                                }


                                string strLogText = "Goto Focus Position " + Convert.ToInt16(BestPos).ToString() + "Current Filter_";

                                string path = textBox11.Text.ToString();
                                string fullpath = path + @"\log.txt";
                                StreamWriter log;
                                if (!File.Exists(fullpath))
                                {
                                    log = new StreamWriter(fullpath);
                                }
                                else
                                {
                                    log = File.AppendText(fullpath);
                                }
                                if (Filtertext != null)
                                    log.WriteLine(strLogText + Filtertext.ToString());
                                else
                                    log.WriteLine(strLogText);

                                log.Close();
                                if (FilterFocusOn == true)
                                {
                                    if (checkBox22.Checked == false)
                                    {
                                        //  FilterFocusOn = false;move to while belwo
                                        ShowWindow(NebhWnd, SW_SHOW);
                                        //  ShowWindow(NebhWnd, SW_RESTORE);
                                        SetForegroundWindow(NebhWnd);//? may not need 3-4
                                        Thread.Sleep(500);//may not need 3-4 both (below too) were 1000 6-1
                                        SetForegroundWindow(Aborthwnd);
                                        PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                                        Thread.Sleep(500);//was 1000
                                        NebListenOn = false;
                                    }
                                    if (checkBox10.Checked == false)
                                    {
                                        // toolStripStatusLabel1.Text = "Slewing to Target Location";
                                        //   this.Refresh();
                                        //    FocusDoneTargetReturn();
                                        //   Thread.Sleep(1000);
                                        //  MessageBox.Show("simulate Slew");
                                        toolStripStatusLabel1.Text = "Slewing to Target";
                                        this.Refresh();
                                        GotoTargetLocation();

                                        while (MountMoving == true)
                                        {
                                            Thread.Sleep(50); // pause for 1/20 second
                                            System.Windows.Forms.Application.DoEvents();
                                        }
                                        if (MountMoving == false)
                                        {
                                            button35.Text = "At Target";
                                            button35.BackColor = System.Drawing.Color.Lime;
                                            button33.Text = "Goto";
                                            button33.UseVisualStyleBackColor = true;
                                        }
                                        //    Thread.Sleep(SlewDelay);
                                        ResumePHD();
                                        //    GotoTargetLocation(); //   *****rem for debugging*****
                                        //   Thread.Sleep(SlewDelay);
                                        //    ResumePHD();
                                    }
                                    //   SetForegroundWindow(NebhWnd);
                                    //   Thread.Sleep(1000);

                                    //    PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                                    /*
                                    if (clientSocket != null)
                                    {

                                        clientSocket.GetStream().Close();//added 5-17-12
                                        clientSocket.Client.Disconnect(true);//added 5-17-12
                                    }
                                    */
                                    if (checkBox19.Checked == true)
                                        ResumePHD();

                                    if (!IsSlave() && (checkBox22.Checked == false))
                                    {

                                        clientSocket.GetStream().Close();
                                        clientSocket.Close();// rem'd 5-17-12                                    

                                    }

                                    if (IsSlave())
                                    {

                                        clientSocket2.GetStream().Close();
                                        clientSocket2.Close();
                                    }


                                    Thread.Sleep(1000);
                                    //    SlavePaused = true;



                                    if (IsServer())
                                    {
                                        Log("Waiting for slave focus");
                                        toolStripStatusLabel1.Text = "Waiting for slave focus";
                                        this.Refresh();
                                        while (working == true)
                                        {
                                            WaitForSequenceDone("Fine focus done", NebSlavehwnd);
                                            //  Thread.Sleep(50);
                                        }
                                        working = true;
                                        Log("Slave focus done");
                                    }

                                    /* // this works below  ********* 6-6 
                                while (FilterFocusOn == true)
                                {
                                    int StatusstripHandle = FindWindowEx(NebSlavehwnd, 0, "msctls_statusbar32", null);

                                    //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                                    IntPtr statusHandle = new IntPtr(StatusstripHandle);
                                    StatusHelper sh = new StatusHelper(statusHandle);
                                    string[] captions = sh.Captions;

                                    if (captions[0] == "Fine focus done")
                                        FilterFocusOn = false;
                                    Thread.Sleep(50);


                                }
                                
                                     */

                                    /*
                                    if (IsSlave())
                                    {
                                        while (FilterFocusOn == true)
                                        {
                                            int StatusstripHandle = FindWindowEx(NebhWnd, 0, "msctls_statusbar32", null);

                                            //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                                            IntPtr statusHandle = new IntPtr(StatusstripHandle);
                                            StatusHelper sh = new StatusHelper(statusHandle);
                                            string[] captions = sh.Captions;

                                            while (captions[0] != "Fine focus done")
                                            {
                                                Thread.Sleep(50); // pause for 1/20 second
                                                System.Windows.Forms.Application.DoEvents();

                                            }
                                            FilterFocusOn = false;
                                        }
                                    }
                                
                   */





                                    /*      ******change to wait for neb status to say "Fine focue done"   *******
                                                                    if (IsServer())
                                                                    {
                                                                        Log("Waiting for Slave focus");
                                    
                                                                        while (textBox41.Text != "Focus Done")
                                                                        {
                                                                            textBox41.Refresh();
                                                                            Thread.Sleep(50); // pause for 1/20 second
                                                                            System.Windows.Forms.Application.DoEvents();
                                                                        }
                                                                    }
                                                                    if (IsSlave())
                                                                    {
                                                                        Log("Waiting for Server focus");
                                                                        while (textBox41.Text != "Focus Done")
                                                                        {
                                                                            textBox41.Refresh();
                                                                            Thread.Sleep(50); // pause for 1/20 second
                                                                            System.Windows.Forms.Application.DoEvents();
                                                                        }
                                                                    }
                                    */
                                    /*
                                    if (IsSlave())
                                    {
                                        while (textBox41.Text != "Capturing")
                                        {
                                            textBox41.Refresh();
                                            Thread.Sleep(50); // pause for 1/20 second
                                            System.Windows.Forms.Application.DoEvents();
                                        }
                                    }
                                    if (IsServer())
                                        SendtoSlave("Capturing");
                                     */

                                    FilterFocusOn = false;
                                    if (!IsSlave())
                                    {

                                        fileSystemWatcher4.EnableRaisingEvents = true;//move this to last 3-4 from under abort/sleep
                                        NebCapture();//*************un remd 3_4

                                    }



                                }

                            }
                            if (checkBox22.Checked == true)
                            {
                                /*
                                 SetForegroundWindow(NebhWnd);
                                 Thread.Sleep(1000);
                                 PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                                 Thread.Sleep(1000);
                                 SendKeys.SendWait("~");
                                 SendKeys.Flush();
                                 Thread.Sleep(100);
                                 */
                                if (metricpath[0] != null)
                                    File.Delete(metricpath[0]);
                                /*
                    //    serverStream = clientSocket.GetStream();
                        byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                        serverStream.Write(outStream, 0, outStream.Length);
                        Thread.Sleep(1000);
                        serverStream.Flush();
                      //  serverStream.Close();
                    //    SetForegroundWindow(NebhWnd);
                   //     Thread.Sleep(1000);
                   //     PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                   //     Thread.Sleep(1000);
                      //  NebListenOn = false;
                     //   clientSocket.GetStream().Close();//added 5-17-12
                                 */
                                //    clientSocket.Client.Disconnect(true);//added 5-17-12  ***use this one
                                //    clientSocket.Close();

                            }

                        }
                        if (FineFocusAbort == true)
                        {
                          //  FineFocusAbort = false;
                            gotoFocus();
                        }
                }
           // }
           /*     
            catch(Exception ex) 
            {
                Log("Unknown Error FilesystemWatcher3 - StdDev/GotoFocus" + ex.ToString());
                Send("Unknown Error FilesystemWatcher3 - StdDev/GotoFocus" + ex.ToString());
                FileLog("Unknown Error FilesystemWatcher3 - StdDev/GotoFocus" + ex.ToString());

            }
            */
              //  
        }

        void standby()
        {
            fileSystemWatcher2.EnableRaisingEvents = false;
            fileSystemWatcher3.EnableRaisingEvents = false;
            fileSystemWatcher1.EnableRaisingEvents = false;
            fileSystemWatcher4.EnableRaisingEvents = false;//added 3_13
            fileSystemWatcher5.EnableRaisingEvents = false; //added to test  metricHFR
            if (closing != 1)
            {
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
                positionbar();
                progressBar1.Value = 0;
            }
            //************* added 3_13  *****************
            /*  *****remd 6-28
            if (clientSocket.Connected == true)
            {
                NebListenStop();
             //   clientSocket.GetStream().Close();//added 5-17-12
             //   clientSocket.Client.Disconnect(true);// 5-17-12
                clientSocket.Close();  //rem'd 5-17-12  ****currently how it was previously
            }
            //end add
             */

                
        }

        void gotopos(Int16 value)
        {
            try
            {
                if (value > travel)
                {
                    DialogResult result1 = MessageBox.Show("Goto Exceeds Max Travel", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (result1 == DialogResult.OK)
                    {
                        return;
                    }
                }
                if (value < 0)
                {
                    DialogResult result1 = MessageBox.Show("Goto Exceeds Full In", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (result1 == DialogResult.OK)
                    {
                        return;
                    }
                }
                port.DiscardInBuffer();
                Thread.Sleep(20);
                port.Write(value.ToString());
                Thread.Sleep(20);
                port.DiscardOutBuffer();
                Thread.Sleep(20);

                //goto progress bar
                int diff = Math.Abs(value - count);
                if (closing == 0) // allows for no progress bar during closure while stepper returns to zero
                {
                    for (int zz = 0; zz < diff; zz++)
                    {
                        if (zz % 8 == 0)
                        {
                           // Thread.Sleep(5);  ************rem'd 6-5 to speep up
                        }
                        if (vcurveenable != 1)
                        {

                            progressBar1.Maximum = diff;
                            progressBar1.Minimum = 0;
                            progressBar1.Increment(10);
                            progressBar1.Value = zz;
                        }
                        count = count + (value - count);
                        textBox1.Text = count.ToString();
                        textBox4.Text = posMin.ToString();

                    }
                    if (vcurveenable != 1)
                    {
                        Log("Goto: " + count.ToString());

                        progressBar1.Value = 0;
                    }
                    positionbar();
                }

                port.DiscardInBuffer();
            }
            catch (Exception e)
            {
                Log("GotoPos Error"+ e.ToString());
                Send("GotoPos Error" + e.ToString());
                FileLog("GotoPos Error" + e.ToString());

            }
        }

        public static double GetStandardDeviation(double[] numb)
        {
            

                double Sumdbl = 0.0, SumOfSqrs = 0.0;
                for (int count = 0; count < numb.Length; count++)
                {
                    Sumdbl += numb[count];
                    SumOfSqrs += Math.Pow(numb[count], 2);
                }
                double topSum = (numb.Length * SumOfSqrs) - (Math.Pow(Sumdbl, 2));
                double val = (double)numb.Length;
                return Math.Sqrt(topSum / (val * (val - 1)));
            
           
        }
        

        public static double GetSlope(double[] xArray, double[] yArray)
        {
            if (xArray == null)
                throw new ArgumentNullException("xArray");
            if (yArray == null)
                throw new ArgumentNullException("yArray");
            if (xArray.Length != yArray.Length)
                throw new ArgumentException("Array Length Mismatch");
            if (xArray.Length < 2)
                throw new ArgumentException("Arrays too short.");

            double n = xArray.Length;
            double sumxy = 0, sumx = 0, sumy = 0, sumx2 = 0;
            for (int i = 0; i < xArray.Length; i++)
            {
                sumxy += xArray[i] * yArray[i];
                sumx += xArray[i];
                sumy += yArray[i];
                sumx2 += xArray[i] * xArray[i];
            }
            return ((sumxy - sumx * sumy / n) / (sumx2 - sumx * sumx / n));
        }


        public static int GetMetric(string[] filePaths, int round)
        {
            
            
                int current2;
                //   long param3;
                foreach (string filename2 in filePaths)
                {
                    int first = filename2.IndexOf(@"_");
                    // int second = filename.IndexOf(".", first + 1);
                    string size = filename2.Substring(first + 1, 3);
                    current2 = Convert.ToInt16(size);
                    return current2;
                    //  param3 = Convert.ToInt32(current2);
                    //   return  ((int)((param3 + (0.5 * round)) / round) * round);
                }
                return 999999999;          

        }


        public static int GetFileHFR(string[] filePaths, int round)
        {
            int current;
            long param2;
            foreach (string filename in filePaths)
            {
                int first = filename.IndexOf(@".");
                int second = filename.IndexOf(".", first + 1);
                string size = filename.Substring(first - 2, 1);
                int num;
                bool isNumeric = int.TryParse(size, out num);
                if (isNumeric)
                {
                    string current2 = filename.Substring(first - 2, second - first);
                    string replace = current2.Replace(".", "");
                    current = Convert.ToInt32(replace);
                }

                else
                {
                    string current2 = filename.Substring(first - 1, second - first - 1);
                    string replace = current2.Replace(".", "");
                    current = Convert.ToInt32(replace);
                }
                param2 = Convert.ToInt32(current);
                return ((int)((param2 + (0.5 * round)) / round) * round);
            }
            return 999999999;
        }

        private static long GetFileMax(string[] filePaths, int round)
        {
            foreach (string filenameMax in filePaths)
            {
                int firstMax = filenameMax.IndexOf(@"_");
                int secondMax = filenameMax.IndexOf("_", firstMax + 1);
                string sizeMax = filenameMax.Substring(firstMax + 1, ((secondMax - firstMax) - 1));

                long paramMax = Convert.ToInt32(sizeMax);

                long roundtoMax = (round * 10);
                return (long)((paramMax + (0.5 * roundtoMax)) / roundtoMax) * roundtoMax;
            }
            return 999999;
        }
        string[] metricpath;
//vcurvehere
        void vcurve()
        {
          //  try
          //  {
                if (currentfilter == 1)
                    Filtertext = comboBox2.Text;
                if (currentfilter == 2)
                    Filtertext = comboBox3.Text;
                if (currentfilter == 3)
                    Filtertext = comboBox4.Text;
                if (currentfilter == 4)
                    Filtertext = comboBox5.Text;
                if (enteredMaxHFR < enteredMinHFR)
                {
                    DialogResult result = MessageBox.Show("Minimum HFR cannot be greater than Maximum HFR ", "scopefocus",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (result == DialogResult.OK)
                    {
                        return;
                    }
                }
                roundto = (int)numericUpDown1.Value;
                MoveDelay = (int)numericUpDown9.Value;

                vcurveenable = 1;
                if (tempon == 1)
                {
                    fileSystemWatcher1.EnableRaisingEvents = true;
                }
                if ((vProgress == 0) & (tempon == 1))
                {
                    int finegoto = posMin - ((((int)numericUpDown3.Value) / 2) * ((int)numericUpDown8.Value));
                    //fine v-curve goes to N/2 * step size in from the focus position -- V should be centered
                    gotopos(Convert.ToInt16(finegoto));
                }

                if (ffenable == 1)
                {
                    textBox17.Text = (fineVrepeatDone + 1).ToString();//exposure repeat
                    repeatTotal = (int)numericUpDown7.Value;

                }
                else
                {
                    repeatTotal = (int)numericUpDown4.Value;
                }
                if (repeatProgress == (repeatTotal - 1))
                {
                    repeatDone = 1;
                }
                if (ffenable == 1)
                {
                    //define Fine-V total sample number (should be twice goto position[line 914 temp cal or 187 fine-v]/step size [line 1135])
                    vN = (int)numericUpDown3.Value;
                }

                else
                {
                    vN = (int)numericUpDown5.Value;
                }
                if (vProgress < vN)
                {
                    progressBar1.Maximum = (vN);
                    progressBar1.Minimum = 0;
                    progressBar1.Increment(1);
                    progressBar1.Value = vProgress + 1;
                }
                if (vProgress == vN)//****** changed to -1 6-28
                {
                    vDone = 1; 
                }
                int avg = 0;
                int current = 0;
                int closestHFR = 1;
                long closestMax = 1;
                if (vDone != 1)
                {
                    if (checkBox22.Checked == true)
                    {
                        //begin test metriHFR
                        metricpath = Directory.GetFiles(path2.ToString(), "metric*.fit");
                        try
                        {
                            testMetricHFR = GetMetric(metricpath, roundto);
                            textBox25.Text = testMetricHFR.ToString();
                            Log("MetricHFR" + testMetricHFR.ToString());
                        }
                        catch (Exception e)
                        {
                            standby();
                            MessageBox.Show("Error parsing filename, ensure there are not any files w/ aaaa_bbbb_cccc_dddd_eeee.fit structure in the capture directory", "scoprefocus");
                            Log("GetHFR error" + e.ToString());
                            Send("GetHFR error" + e.ToString());
                            FileLog("GetHFR error" + e.ToString());

                            return;


                        }

                            closestHFR = testMetricHFR;
                        
                    }
                    int[] templist = new int[((vN * repeatTotal) + 1)];
                    string[] filePaths = Directory.GetFiles(path2.ToString(), "*.bmp");
                    if (checkBox22.Checked == false)//added for metrichfr
                    {
                        closestHFR = GetFileHFR(filePaths, roundto);
                    }
                    if (checkBox22.Checked == false)//add for metricHFR
                    {
                        closestMax = GetFileMax(filePaths, roundto);
                        if (repeatProgress == 0)
                        {
                            peatMax[repeatProgress] = closestMax;
                            avgMax = closestMax;
                            sumMax = closestMax;
                        }
                        if (repeatProgress != 0)
                        {
                            peatMax[repeatProgress] = closestMax;
                            sumMax = sumMax + (int)peatMax[repeatProgress];

                            avgMax = sumMax / (repeatProgress + 1);
                        }
                        //addition for max

                        if (repeatDone == 1)
                        {
                            listMax[vProgress] = avgMax;

                        }
                    }//end metricHFR addtion "if" 
                    //setup positions array for HFR
                    minHFRpos[vProgress] = count;
                    if (checkBox22.Checked == false)//add for metriHFR
                    {
                        maxMaxPos[vProgress] = count;
                    }
                    if (checkBox22.Checked == true)//added for metricHFR
                    {
                        maxMax = 1;//added for metricHFR

                    }
                    //determine minHFR and build array
                    if (repeatProgress == 0)
                    {
                        peat[repeatProgress] = closestHFR;
                        avg = closestHFR;
                        sum = closestHFR;
                    }

                    if (repeatProgress != 0)
                    {
                        peat[repeatProgress] = closestHFR;
                        sum = sum + peat[repeatProgress];

                        avg = sum / (repeatProgress + 1);
                    }
                    //repeatDone=1 when repeat done

                    if (repeatDone == 1)
                    {
                        list[vProgress] = avg;
                    }
                    if (repeatDone != 1)
                    {
                        repeatProgress++;
                        vv++;
                    }
                    if (repeatDone == 1)
                    {

                        templist[vv] = temp;
                        tempsum = tempsum + templist[vv];
                        tempavg = tempsum / (vv + 1);
                        total = total + 1;
                        textBox1.Text = count.ToString();
                        textBox4.Text = posMin.ToString();
                        //progress current is textbx 7 of total N is textbox6
                        textBox6.Text = vN.ToString();
                        textBox7.Text = (vProgress + 1).ToString();
                        Log("V-Curve: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t\tPos" + count.ToString() + " \t  HFRavg" + avg.ToString() + "\tMaxavg" + avgMax.ToString() + "\t" + Filtertext);
                        string strLogTextA = "V-Curve: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + count.ToString() + " \t" + avg.ToString() + "\t" + avgMax.ToString();
                        string strLogText = "V-curve" + "\t" + temp.ToString() + "\t" + count.ToString() + "\t" + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + current.ToString() + "\t" + min.ToString() + "\t" + posMin.ToString();
                        string strLogText2 = "TempCal" + "\t " + tempavg.ToString() + "\t " + posMin.ToString();
                        string strLogText3 = "Fine-V: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + count.ToString() + " \t" + avg.ToString() + "\t" + avgMax.ToString() + "\t" + min.ToString() + "\t" + maxMax.ToString() + "\t" + posMin.ToString() + "\t" + Filtertext;
                        positionbar();
                        if ((tempon != 1) & (ffenable != 1) & (repeatDone == 1))
                        {
                            chart1.Series[0].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg));
                        }
                        if ((ffenable == 1) & (tempon == 0) & (repeatDone == 1) & (backlashDetermOn == false))
                        {
                            if ((avg < enteredMinHFR) || (avg > enteredMaxHFR) & (radioButton1.Checked == true))
                            {
                                chart1.Series[2].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg));
                            }
                            else
                            {
                                chart1.Series[1].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg));
                            }
                        }

                        string path = textBox11.Text.ToString();
                        string fullpath = path + @"\log.txt";

                        // Create a writer and open the file:
                        StreamWriter log;

                        if (!File.Exists(fullpath))
                        {
                            log = new StreamWriter(fullpath);
                        }
                        else
                        {
                            log = File.AppendText(fullpath);
                        }
                        // Write to the file:
                        if (backlashDetermOn == false)
                        {
                            if ((ffenable == 1) & (vProgress == 0))
                            {
                                log.WriteLine(DateTime.Now);
                                log.WriteLine("Fine-V" + "\tTemp" + "\t Pos" + "\tN" + "\tHFR" + "\tminHFR" + "\tposmin");
                            }
                            if ((vProgress == 0) & (tempon == 0))
                            {
                                log.WriteLine(DateTime.Now);
                                log.WriteLine("Type  N:" + "\t  Pos" + "\tHFRAvg" + "\tMaxAvg" + "\tminHFR" + " \tmaxMax" + maxMax.ToString() + "\tposmin");
                            }
                            if ((tempon == 1) & (vProgress == (vN - 1)))
                            {
                                if (templog == 0)
                                {
                                    log.WriteLine(DateTime.Now);
                                    log.WriteLine("Type" + "\tAvgTemp" + "\tposMin");
                                    // log.WriteLine(strLogText2);
                                    templog = 1;
                                    // log.Close();
                                }

                                log.WriteLine(strLogText2);

                            }
                            if ((tempon == 0) & (ffenable != 1))
                            {
                                log.WriteLine(strLogTextA);
                            }
                            if ((tempon == 0) & (ffenable == 1))
                            {
                                log.WriteLine(strLogText3);
                            }
                        }

                        log.Close();
                        repeatDone = 0;
                        //reset repeat
                        repeatProgress = 0;
                        sum = 0;
                        if (vProgress < vN)
                        {
                            int step = (int)numericUpDown2.Value;
                            if (ffenable == 1)
                            {
                                //defines Fine-V step size
                                step = (int)numericUpDown8.Value;
                            }
                            if (backlashDetermOn != true)
                            {
                                gotopos(Convert.ToInt16(count + step));
                                vProgress++;
                                vv++;
                                Thread.Sleep(MoveDelay);//helps prevent focus mvmnt during capture
                            }
                            if (backlashDetermOn == true)
                            {
                                if (backlashOUT == false)
                                {
                                    if (vProgress < (vN - 1))
                                    {
                                        gotopos(Convert.ToInt16(count - step));
                                    }
                                    vProgress++;
                                    vv++;
                                    Thread.Sleep(MoveDelay);
                                }
                                if (backlashOUT == true)
                                {
                                    if (vProgress < (vN - 1))
                                    {
                                        gotopos(Convert.ToInt16(count + step));
                                    }
                                    vProgress++;
                                    vv++;
                                    Thread.Sleep(MoveDelay);
                                }
                            }

                        }
                    }


                }
                //V=1...skips to here.     
                if (vDone == 1)
                {
                    fileSystemWatcher2.EnableRaisingEvents = false;
                    fileSystemWatcher5.EnableRaisingEvents = false; //added to test metritHFR     

                    for (int arraycount = 0; arraycount < vProgress; arraycount++)
                    {
                        if (list[arraycount] < HFRarraymin)
                        {
                            HFRarraymin = list[arraycount];
                            posminHFR = minHFRpos[arraycount];
                            min = HFRarraymin;
                            apexHFR = arraycount;
                            textBox4.Text = posMin.ToString();
                        }

                    }
                    for (int arraycount = 0; arraycount < vProgress; arraycount++)
                    {
                        bool MaxtooHi = false;
                        if (listMax[arraycount] > 35000)
                        {
                            MaxtooHi = true;
                        }
                        if (listMax[arraycount] > maxarrayMax)
                        {
                            maxarrayMax = listMax[arraycount];
                            posmaxMax = (int)maxMaxPos[arraycount];
                            if ((posmaxMax == posminHFR) & (MaxtooHi == false))
                            {
                                posMin = posmaxMax;

                            }

                        }
                    }

                    //make sure v-curve is symetric
                    //                                                                                   
                    if (radioButton1.Checked == true)
                    {                                //below was vN/2 +4 and -4 changed to +/- vN/4 11-24
                        if (((apexHFR > (vN / 2 + vN / 4)) || (apexHFR < (vN / 2 - vN / 4))) & (ffenable == 1) & (backlashDetermOn == false))
                        {
                            fileSystemWatcher2.EnableRaisingEvents = false;
                            fileSystemWatcher5.EnableRaisingEvents = false; //added to test metricHFR
                            DialogResult result2 = MessageBox.Show("V-Curve Not Symmetric", "Aborted - Repeat Rough V-Curve", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (result2 == DialogResult.OK)
                            {

                                standby();
                            }
                        }
                    }
                    if ((ffenable == 1) & (radioButton1.Checked == true))
                    {


                        //first figure out how many points on each side of v-curve based on entered min/max
                        for (int arraycount = 0; arraycount < apexHFR; arraycount++)
                        {
                            if ((list[arraycount] < enteredMaxHFR) & (list[arraycount] > enteredMinHFR))
                            {
                                arraycountright++;
                            }
                        }
                        for (int arraycount2 = apexHFR; arraycount2 < vN; arraycount2++)
                        {
                            if ((list[arraycount2] < enteredMaxHFR) & (list[arraycount2] > enteredMinHFR))
                            {
                                arraycountleft++;
                            }
                        }
                        double[] HFRdwn = new double[arraycountright];
                        double[] HFRup = new double[arraycountleft];
                        double[] HFRposdwn = new double[arraycountright];
                        double[] HFRposup = new double[arraycountleft];
                        int x = 0;
                        for (int arraycount = 0; arraycount < apexHFR; arraycount++)
                        {

                            //define right side < entered max and < entered min
                            if ((list[arraycount] < enteredMaxHFR) & (list[arraycount] > enteredMinHFR))
                            {
                                HFRdwn[x] = list[arraycount];
                                HFRposdwn[x] = minHFRpos[arraycount];
                                x++;
                            }


                        }
                        int y = 0;
                        for (int arraycount2 = apexHFR; arraycount2 < vN; arraycount2++)
                        {
                            //define left side < entered max > min

                            if ((list[arraycount2] < enteredMaxHFR) & (list[arraycount2] > enteredMinHFR))
                            {
                                HFRup[y] = list[arraycount2];
                                HFRposup[y] = minHFRpos[arraycount2];
                                y++;
                            }

                        }



                        //Begin slope calc based on entered min max array*****************


                        slopeHFRdwn = (1 / GetSlope(HFRdwn, HFRposdwn));
                        slopeHFRup = (1 / GetSlope(HFRup, HFRposup));
                        //use point 1/2 the way 
                        XintDWN = minHFRpos[apexHFR - vN / 4] - list[apexHFR - vN / 4] / slopeHFRdwn;
                        XintUP = minHFRpos[apexHFR + vN / 4] - list[apexHFR + vN / 4] / slopeHFRup;
                        PID = (int)XintDWN - (int)XintUP;
                        posMin = posminHFR;
                        //use point half way up each side to calc intersect position can be used as relative focus position
                        intersectPos = (int)GetIntersectPos((double)minHFRpos[apexHFR - vN / 4], (double)minHFRpos[apexHFR + vN / 4], (double)list[apexHFR - vN / 4], (double)list[apexHFR + vN / 4], slopeHFRdwn, slopeHFRup);
                        //all vN/2 above were 5

                        textBox4.Text = posMin.ToString();
                        textBox15.Text = HFRarraymin.ToString();
                        WriteSQLdata(); ;
                        FillData();
                        Log("Slope: N " + (vProgress + 1).ToString() + "\tslopeUP" + slopeHFRup.ToString() + " \tSlopeDWN" + slopeHFRdwn.ToString() + "\tIntersect" + intersectPos.ToString() + "\tPID" + PID.ToString() + "\t" + Filtertext);

                    }
                    else
                    {
                        posMin = posminHFR;
                    }

                    textBox4.Text = posMin.ToString();
                    textBox1.Text = current.ToString();

                    string path = textBox11.Text.ToString();
                    string fullpath = path + @"\log.txt";
                    string fullpathData = path + @"\data.txt";
                    StreamWriter log;
                    StreamWriter logData;

                    if (!File.Exists(fullpath))
                    {
                        log = new StreamWriter(fullpath);
                        logData = new StreamWriter(fullpathData);
                    }
                    else
                    {
                        log = File.AppendText(fullpath);
                        logData = File.AppendText(fullpathData);

                    }
                    log.WriteLine("Posmin: " + posMin.ToString() + "\tminHFR " + min.ToString() + "\tminHFRpos " + posminHFR.ToString() + "\tmaxMAx " + maxarrayMax.ToString() + "\tmaxMaxPos " + posmaxMax.ToString());
                    //removed & radiobutton1,checked from below
                    if (ffenable == 1)
                    {
                        for (int xx = 0; xx < vN; xx++)
                        {
                            if (xx == 0)
                            {
                                log.WriteLine(DateTime.Now);
                                log.WriteLine(Filtertext);
                            }
                            log.WriteLine("HFR " + list[xx].ToString() + "\tPosition " + minHFRpos[xx].ToString());
                        }
                        logData.WriteLine(DateTime.Now + "\t" + vN.ToString() + "\t" + slopeHFRdwn.ToString() + "\t" + slopeHFRup.ToString() + "\t" + XintDWN.ToString() + "\t" + XintUP.ToString() + "\t" + PID.ToString() + "\t" + apexHFR.ToString() + "\t" + Filtertext);
                    }
                    log.Close();
                    logData.Close();
                    if ((vDone == 1) & (ffenable == 1))
                    {

                        fileSystemWatcher1.EnableRaisingEvents = false;
                        fileSystemWatcher2.EnableRaisingEvents = false;
                        fileSystemWatcher5.EnableRaisingEvents = false; // added to test metric HFR
                        //handle repeated fine v curves
                        if (fineVrepeat > 1)
                        {
                            button3.PerformClick();

                        }
                        else
                        {
                            if (checkBox22.Checked == true)
                            {
                                /*
                                serverStream = clientSocket.GetStream();
                                byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                                serverStream.Write(outStream, 0, outStream.Length);
                                serverStream.Flush();
                                if (metricpath[0] != null)
                             //       File.Delete(metricpath[0]);
                                Thread.Sleep(3000);
                                serverStream.Close();
                                Thread.Sleep(3000);
                                 */
                                SetForegroundWindow(NebhWnd);
                                Thread.Sleep(1000);
                                PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                                Thread.Sleep(1000);
                                NebListenOn = false;
                              //  File.Delete(metricpath[0]);
                                /*
                                // clientSocket.GetStream().Close();//added 5-17-12
                                //  clientSocket.Client.Disconnect(true);//added 5-17-12
                                clientSocket.Close();
                                */
                                if (metricpath != null)
                                    File.Delete(metricpath[0]);
                                currentmetricN = 0;
                            }
                            fineVrepeatDone = 0;
                            fineVrepeatOn = false;
                            standby();
                            
                        }

                    }
                    //end course v-curve and goes to rough focus point
                    if ((vDone == 1) & (tempon == 0) & (backlashDetermOn == false) & (ffenable != 1))
                    {
                        fileSystemWatcher1.EnableRaisingEvents = false;
                        gotopos(Convert.ToInt16(posMin));
                        standby();
                    }
                    //reset for more v -curves for tempcal
                    if ((tempon == 1) & (vDone != 1))
                    {
                        repeatDone = 0;
                        repeatTotal = 0;
                        fileSystemWatcher1.EnableRaisingEvents = true;
                    }
                    //tempcal done
                    if ((tempon == 1) & (vDone == 1))
                    {
                        tempsum = 0;
                        vDone = 0;
                        vv = 0;
                        min = 500;
                        vProgress = 0;
                        fileSystemWatcher1.EnableRaisingEvents = false;
                        int finegoto = posMin - ((((int)numericUpDown3.Value) / 2) * ((int)numericUpDown8.Value));
                        if ((finegoto - 100) < 0)
                        {
                            DialogResult result1 = MessageBox.Show("Goto Exceeds Full In", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (result1 == DialogResult.OK)
                            {
                                return;
                            }
                        }
                        chart1.Series[3].Points.AddXY(Convert.ToDouble(posMin), Convert.ToDouble(tempavg));
                        gotopos(Convert.ToInt16(finegoto - 100));//take up backlash                            
                        Thread.Sleep(2000);
                        gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 4));
                        Thread.Sleep(1000);
                        gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 3));
                        Thread.Sleep(1000);
                        gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 2));
                        Thread.Sleep(1000);
                        gotopos(Convert.ToInt16(finegoto - (int)numericUpDown8.Value));
                        Thread.Sleep(1000);
                        gotopos(Convert.ToInt16(posMin));

                        repeatProgress = 0;
                        vcurveenable = 0;
                        Array.Clear(list, 0, arraysize2);
                        Array.Clear(listMax, 0, arraysize2);
                        HFRarraymin = 999;
                        maxarrayMax = 1;
                        tempcal();

                    }
                    vcurveenable = 0;//added 9-12
                }

                if ((backlashDetermOn == true) & (vDone == 1))
                {
                    backlashCount++;
                    if (backlashCount == backlashN - 1)
                    {
                        standby();
                    }
                    if (backlashOUT == true)
                    {
                        backlashOUT = false;
                        backlashPosOUT = posMin;
                    }
                    else
                    {

                        backlashOUT = true;
                        backlashPosIN = posMin;
                    }

                    if ((backlashCount == 2) || (backlashCount == 4) || (backlashCount == 6) || (backlashCount == 8) || (backlashCount == 10))
                    {

                        backlash = Math.Abs(backlashPosIN - backlashPosOUT);
                        textBox8.Text = backlash.ToString();
                        chart1.Series[2].Points.AddXY(Convert.ToDouble(backlashCount), Convert.ToDouble(backlash));
                        backlashSum = backlash + backlashSum;
                        backlashAvg = backlashSum / (backlashCount / 2);
                        Log("Backlash: N " + backlashCount.ToString() + "\tPosOUT " + backlashPosOUT.ToString() + "\tPosIN " + backlashPosIN.ToString() + "\tBacklash " + backlash.ToString() + "\tAvg " + backlashAvg.ToString());
                        string path = textBox11.Text.ToString();
                        string fullpath = path + @"\log.txt";

                        StreamWriter log;
                        if (!File.Exists(fullpath))
                        {
                            log = new StreamWriter(fullpath);
                        }
                        else
                        {
                            log = File.AppendText(fullpath);
                        }
                        if (backlashCount == 2)
                        {
                            log.WriteLine(DateTime.Now);
                            log.WriteLine("Backlash N" + "\tPosOUT" + "\tPosIN" + "\tBacklash" + "\tAvg");
                        }
                        log.WriteLine("         " + backlashCount.ToString() + "\t  " + backlashPosOUT.ToString() + "\t " + backlashPosIN.ToString() + "\t   " + backlash.ToString() + "\t\t" + backlashAvg.ToString());
                        if (backlashCount == backlashN)
                        {
                            log.WriteLine("Avg Backlash: " + backlashAvg.ToString());
                            textBox8.Text = "Avg " + backlashAvg.ToString();

                        }
                        log.Close();

                    }

                    tempsum = 0;
                    vDone = 0;
                    vv = 0;
                    min = 500;

                    maxMax = 1;
                    fileSystemWatcher1.EnableRaisingEvents = true;

                    vProgress = 0;
                    repeatProgress = 0;
                    //?????not sure why below???
                    posMin = count;
                    Array.Clear(listMax, 0, arraysize2);
                    Array.Clear(list, 0, arraysize2);
                    HFRarraymin = 999;
                    maxarrayMax = 1;




                }

           // }
           /* 
            catch (Exception e)
            {
                Log("Unknwon Error - Vcurve  line 920-1530" + e.ToString());
                Send("Unknwon Error - Vcurve  line 920-1530" + e.ToString());
                FileLog("Unknwon Error - Vcurve  line 920-1530" + e.ToString());

            }
         */    
        }
            
        private void chart1_Click(object sender, EventArgs e)
        {
            chart1.Series[1].Points.AddXY(Convert.ToDouble(min), Convert.ToDouble(count));

        }
        
        
        //added to sync for manual changes
        void sync()
        {
            try
            {
                int syncfailure = 0;
                string numbers = "";
                int numb = 0;
                textBox1.Clear();
                string str = "";
                Thread.Sleep(10);
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
                Thread.Sleep(10);
                port.Write("P");
                Thread.Sleep(50);
                str = port.ReadExisting();
                Thread.Sleep(50);
                port.DiscardInBuffer();
                numbers = string.Join(null, System.Text.RegularExpressions.Regex.Split(str, "[^\\d]"));
                if (numbers == "")
                {
                    DialogResult result;
                    syncfailure = 1;
                    port.Write("E");
                    numbers = count.ToString();
                    Log("Sync Failed -- Retry");
                    result = MessageBox.Show("Sync Failure -- Retry ", "Sync Failure",
                                MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.Retry)
                    {
                        button9.PerformClick();
                    }
                }
                numb = Int32.Parse(numbers);
                if (numb > travel)
                {
                    DialogResult result;
                    syncfailure = 1;
                    port.Write("E");
                    numbers = count.ToString();
                    Log("Sync Exceeds Maximum Travel");
                    result = MessageBox.Show("Sync Exceed Maximum Travel", "Sync Failure",
                                MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.Retry)
                    {
                        button9.PerformClick();

                    }
                }
                syncval = numb;
                if (syncfailure != 1)
                {
                    Log("Synced to position " + syncval.ToString());
                    textBox1.Text = numb.ToString();
                    string strLogText4 = "Synced to position " + syncval.ToString();
                    positionbar();
                    string path = textBox11.Text.ToString();
                    string fullpath = path + @"\log.txt";
                    // Create a writer and open the file:
                    StreamWriter log;

                    if (!File.Exists(fullpath))
                    {
                        log = new StreamWriter(fullpath);
                    }
                    else
                    {
                        log = File.AppendText(fullpath);
                    }
                    log.WriteLine(strLogText4);
                    log.Close();
                }
            }
            catch(Exception e)
            {
                Log("Sync Error line 1547-1632" + e.ToString());
                Send("Sync Error line 1547-1632" + e.ToString());
                FileLog("Sync Error line 1547-1632" + e.ToString());

            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String strVersion = System.Reflection.Assembly.GetCallingAssembly().GetName().Version.ToString();

            MessageBox.Show("           Arduino Version" + strVersion.ToString() + "\r" + "           Kevin Sipprell MD\r           www.scopefocus.info\r      ", "scopefocus for nebulosity",
            MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        //tempcal, backlash and mult v-curve
        private void fileSystemWatcher1_Changed(object sender, FileSystemEventArgs e)
        {
            if (tempon == 1)
            {
                tempcal();
            }
            if (backlashDetermOn == true)
            {
                backlashDeterm();
            }
        }

        void backlashDeterm()
        {
            backlashDetermOn = true;
            fileSystemWatcher1.EnableRaisingEvents = true;
            port.DiscardInBuffer();
            ffenable = 1;
            tempon = 0;
            vcurve();
        }

        void tempcal()
        {

            fileSystemWatcher1.EnableRaisingEvents = true;
            port.DiscardInBuffer();
            ffenable = 1;
            tempon = 1;
            vcurve();
        }

        public static int GetBestHFRPos(int Xa, int Ya, int Yb, int ma)
        {
            return (Xa + (Yb - Ya) / ma);
        }


        public static double GetIntersectPos(double upX, double downX, double upY, double downY, double upSlope, double downSlope)
        {
            return (((upSlope * upX) - (downSlope * downX) + downY - upY) / (upSlope - downSlope));
        }

        public static double GetPositionHFR(int Xa, int Ya, int Xb, int Yb, int Xc, int Yc, int Xd, int Yd)
        {
            double m1 = (Ya - Yb) / (Xa - Xb);
            double m2 = (Yc - Yd) / (Xc - Xd);
            return (((m1 * Xa) - (m2 * Xc) + Yc - Ya) / (m1 - m2));
        }

        public static double GetPositionMax(double Xa, double Ya, double Xb, double Yb, double Xc, double Yc, double Xd, double Yd)
        {
            double m1 = (Ya - Yb) / (Xa - Xb);
            double m2 = (Yc - Yd) / (Xc - Xd);
            return (((m1 * Xa) - (m2 * Xc) + Yc - Ya) / (m1 - m2));
        }

       //gotofocushere
        //goes to near focus point uses std dev routine, takes 10 exposures, calculates best focus and goes there
        public void gotoFocus()
        {
            try
            {

                GetAvg();
                FillData();
                toolStripStatusLabel1.Text = "Taking up Backlash";
                this.Refresh();
                if (radioButton2.Checked == true)//using upslope
                {//this was original 
                    gotopos(Convert.ToInt16(posMin + 50));
                    Thread.Sleep(3000);
                    gotopos(Convert.ToInt16(posMin + 20));
                    Thread.Sleep(3000);
                }
                if (radioButton3.Checked == true)//using downslope added 11-21
                {
                    gotopos(Convert.ToInt16(posMin - 50));
                    Thread.Sleep(3000);
                    gotopos(Convert.ToInt16(posMin - 30));
                    Thread.Sleep(3000);
                }
                port.DiscardInBuffer();
                port.DiscardOutBuffer();
                fileSystemWatcher2.EnableRaisingEvents = false;
                fileSystemWatcher5.EnableRaisingEvents = false; //added to test metricHFR
                fileSystemWatcher3.EnableRaisingEvents = true;
                fileSystemWatcher1.EnableRaisingEvents = false;
                fileSystemWatcher4.EnableRaisingEvents = false;

                GotoFocusOn = true;//~line 440
                min = 1;
                sum = 0;
                avg = 0;
                vDone = 0;
                vProgress = 0;
                vv = 0;
                list = new int[10];
                abc = new double[10];
                if (checkBox22.Checked == true)
                {
                    fileSystemWatcher3.Filter = "*.fit";
  
                 //   if (clientSocket.Connected == false)
                 //   {
                        //clientSocket.Connect("127.0.0.1", SocketPort);//connects to neb
                        NebListenStart(NebhWnd, SocketPort);
                 //   }

                        
                    MetricCapture();
                }
                else
                    fileSystemWatcher3.Filter = "*.bmp";
            }
            catch(Exception e)
            {
                Log("GotoFocus Error" + e.ToString());
                Send("GotoFocus Error" + e.ToString());
                FileLog("GotoFocus Error" + e.ToString());

            }
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        //calc slope stddev from gridview  
        public double CalculateSD(List<double> values)
        {
            List<double> tmpList = new List<double>();
            double average = values.Average();
            values.ForEach(d => tmpList.Add((d - average) * (d - average)));

            return Math.Sqrt(tmpList.Average());
        }



        //update button
        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                if (equip == null)
                {
                    MessageBox.Show("Must select equipment first", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                GetAvg();
                FillData();

                //try adding std dev and display in textbox16
                // Std Dev UP
                if (dataGridView1.RowCount > 2)
                {
                    List<double> StdDevCalcUP = new List<double>();
                    for (int i = 0; i < (dataGridView1.Rows.Count - 1); i++)
                    {
                        StdDevCalcUP.Add(Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value));
                    }

                    double standardDeviationUP = CalculateSD(StdDevCalcUP);
                    double sdUP = Math.Round(standardDeviationUP, 5);
                    textBox16.Text = sdUP.ToString();
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
                if (dataGridView1.RowCount > 2)
                {
                    List<double> StdDevCalcDWN = new List<double>();
                    for (int i = 0; i < (dataGridView1.Rows.Count - 1); i++)
                    {
                        StdDevCalcDWN.Add(Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value));
                    }

                    double standardDeviationDWN = CalculateSD(StdDevCalcDWN);
                    double sdDWN = Math.Round(standardDeviationDWN, 5);
                    textBox14.Text = sdDWN.ToString();
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
            catch(Exception ex)
            {
                Log("Update Error - Line 1817" + ex.ToString());
            }
        }

        private void GetHandles()
        {
            try
            {
                
                Callback myCallBack = new Callback(EnumChildGetValue);

                //    NebhWnd = FindWindow(null, "Nebulosity v3.0-a6");
                //    LoadScripthwnd = FindWindow(null, "Load script");
                //    Log("Load script " + LoadScripthwnd.ToString());
                //  SetForegroundWindow(NebhWnd);
                //   if (NebhWnd == 0)
                //    {
                //        MessageBox.Show("Please Start Calling Window Application");
                //    }
                

                EnumChildWindows(LoadScripthwnd, myCallBack, 0);
            }
            catch(Exception e)
            {
                Log("GetHandles Error" + e.ToString());
                Send("GetHandles Error" + e.ToString());
                FileLog("GetHandles Error" + e.ToString());

            }
       
        }

        public string MainWindowName;
        private bool FindHandlesDone = false;
        public void FindHandles()
        {
            //******5-30 this might be problem, maybe this should just be done once...not sure if 
            //finding the loadscript though. consider doing seperate for the 2 neb windows.  one 
            //with (spawned)
               try
                {

                    MainWindowName = "Nebulosity";
                    IntPtr hWnd2 = SearchForWindow("wxWindow", MainWindowName);
                    Log("Neb Handle Found -- " + hWnd2.ToInt32());
                    NebhWnd = hWnd2.ToInt32();

                    if (NebhWnd != 0)
                    {
                        //finds the Neb version (the number after the "v")
                        StringBuilder sb = new StringBuilder(1024);
                        SendMessage(NebhWnd, WM_GETTEXT, 1024, sb);
                        //  GetWindowText(camera, sb, sb.Capacity);
                        Log(sb.ToString());
                        NebVersion = sb.ToString();
                        int NebVposNumber = NebVersion.IndexOf("v");
                        string NebVNumberAfter = NebVersion.Substring(NebVposNumber + 1, 1);
                        NebVNumber = Convert.ToInt16(NebVNumberAfter);
                    }





                    if (checkBox14.Checked == false)//don't need for slave mode
                    {
                        IntPtr PHDhwnd2 = SearchForWindow("wxWindow", "PHD");
                        Log("PHD Handle Found --  " + PHDhwnd2.ToInt32());
                        PHDhwnd = PHDhwnd2.ToInt32();
                    }
                    // PHDhwnd = FindWindow(null, "PHD Guiding 1.13.0b  -  www.stark-labs.com (Log active)");
                    LoadScripthwnd = FindWindow(null, "Load script");
                    //   Log("PHD " + PHDhwnd.ToString());
                    //   Log("Load script " + LoadScripthwnd.ToString());


                  
                }
                   
                catch (Exception e)
                {
                    Log("FindHandles Error" + e.ToString());
                    Send("FindHandles Error" + e.ToString());
                    FileLog("FindHandles Error" + e.ToString());

                }
            
        }

        private void MainWindow_Load_1(object sender, EventArgs e)
        {
           try
            {

                textBox24.Text = PublishVersion.ToString();//version number on about tab
                //************ try addition for new equipcomboxes
                if (WindowsFormsApplication1.Properties.Settings.Default.ComboItems7 == null)
                {
                    WindowsFormsApplication1.Properties.Settings.Default.ComboItems7 = new System.Collections.Specialized.StringCollection();
                }

                foreach (string item in WindowsFormsApplication1.Properties.Settings.Default.ComboItems7)
                {
                    comboBox7.Items.Add(item);
                }
                if (WindowsFormsApplication1.Properties.Settings.Default.ComboItems8 == null)
                {
                    WindowsFormsApplication1.Properties.Settings.Default.ComboItems8 = new System.Collections.Specialized.StringCollection();
                }

                foreach (string item in WindowsFormsApplication1.Properties.Settings.Default.ComboItems8)
                {
                    comboBox8.Items.Add(item);
                }
               
                Callback myCallBack = new Callback(EnumChildGetValue);
                FindHandles();
                //   int hWnd;
                if (PHDhwnd == 0)
                {
                    DialogResult result1;
                    result1 = MessageBox.Show("PHD Not Found - Open and 'Retry', 'Ignore' or 'Abort' to Close", "scopefocus",
                         MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                    if (result1 == DialogResult.Ignore)
                        this.Refresh();
                    if (result1 == DialogResult.Retry)
                    {
                        FindHandles();
                        this.Refresh();
                    }
                    if (result1 == DialogResult.Abort)
                        System.Environment.Exit(0);
                }


                if (NebhWnd == 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Nebulosity Not Found - Open or Close & Restart and hit 'Retry' or 'Ignore' to continue",
                       "scopefocus", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);//change so ok moves focus
                    if (result == DialogResult.Ignore)
                        this.Refresh();
                    if (result == DialogResult.Retry)
                    {
                        FindHandles();
                        this.Refresh();
                    }
                    if (result == DialogResult.Abort)
                        System.Environment.Exit(0);

                }
                else
                {

                    EnumChildWindows(NebhWnd, myCallBack, 0);
                }
                FindNebCamera();
                //5 lines below for path2 settings
                path2 = WindowsFormsApplication1.Properties.Settings.Default.path.ToString();
                textBox11.Text = path2.ToString();
                NebPath = WindowsFormsApplication1.Properties.Settings.Default.NebPath.ToString();
                textBox35.Text = NebPath.ToString();
                // send = NebPath + ScriptName;
                // textBox37.Text = send.ToString();
                toolStripStatusLabel5.Text = "Path: " + path2.ToString();

                textBox27.Text = user.ToString();
                textBox13.Text = server.ToString();
                textBox28.Text = pswd.ToString();
                textBox36.Text = to.ToString();
                fileSystemWatcher1.Path = path2.ToString();
                fileSystemWatcher2.Path = path2.ToString();
                fileSystemWatcher5.Path = path2.ToString();//added to test metricHFR
                fileSystemWatcher3.Path = path2.ToString();
                fileSystemWatcher4.Path = path2.ToString();
             //   travel = WindowsFormsApplication1.Properties.Settings.Default.travel;
                textBox2.Text = travel.ToString();

                //   camera = WindowsFormsApplication1.Properties.Settings.Default.camera;
                //   textBox22.Text = camera;

                setarraysize();
 
               
                FillData();
              
                toolStripStatusLabel1.Text = "Startup Complete";
            
            }
            catch(Exception ex)
            {
                Log("MainWindow_Load Error" + ex.ToString());
            }
            
        }

        private void ButtonDisable_Click_1(object sender, EventArgs e)
        {
            //*****add for new equip comboboxes
          WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Clear();
                foreach (string item in comboBox7.Items)
                {
               WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Add(item);
                }
                WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Clear();
                foreach (string item in comboBox8.Items)
                {
                    WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Add(item);
                }
            //******end addt.

        //        WindowsFormsApplication1.Properties.Settings.Default.port = portselected.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.path = path2.ToString();
          //  WindowsFormsApplication1.Properties.Settings.Default.camera = textBox22.Text.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.NebPath = NebPath.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.user = this.user;
            WindowsFormsApplication1.Properties.Settings.Default.pswd = this.pswd;
            WindowsFormsApplication1.Properties.Settings.Default.server = this.server;
            WindowsFormsApplication1.Properties.Settings.Default.to = this.to;
            int i = comboBox7.SelectedIndex;
            if (i == 0)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel1 = travel;
            if (i == 1)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel2 = travel;
            if (i == 2)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel3 = travel;
            if (i == 3)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel4 = travel;
            if (i == 4)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel5 = travel;
            if (i == 5)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel6 = travel;
            
            WindowsFormsApplication1.Properties.Settings.Default.Save();
            closing = 1;
            standby();
            if (port == null)
            {
                System.Environment.Exit(0);
            }
            else
            {

                gotopos(0);
                port.Close();
            }

            System.Environment.Exit(0);
        }
        //Abort
        private void button4_Click_2(object sender, EventArgs e)
        {
            fineVrepeatOn = false;//added 1-7-12
            fineVrepeatDone = 0;
            standby();
        }
        //reverse
        private void button2_Click_1(object sender, EventArgs e)
        {
            watchforOpenPort();
            if (portopen == 1)
            {
                int step = (int)numericUpDown2.Value;
                gotopos(Convert.ToInt16(count + step));

                textBox1.Text = count.ToString();
                textBox4.Text = posMin.ToString();
                positionbar();
            }
        }
        //v-curve
        private void button6_Click_1(object sender, EventArgs e)
        {
            if (NebCamera == "No camera")
                NoCameraSelected();
            setarraysize();
            watchforOpenPort();
            if (portopen == 1)
            {
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
                backlashDetermOn = false;
                repeatProgress = 0;
                repeatDone = 0;
                vProgress = 0;
                vDone = 0;
                repeatTotal = 0;
                vN = 0;
                sum = 0;
                min = 500;
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                chart1.Series[2].Points.Clear();
                tempon = 0;
                ffenable = 0;
                posMin = count;
                vv = 0;
                tempsum = 0;
                maxMax = 0;
                roughvdone = true;
                fileSystemWatcher2.EnableRaisingEvents = true;
                fileSystemWatcher5.EnableRaisingEvents = true;//added to test metricHFR
                fileSystemWatcher3.EnableRaisingEvents = false;
                fileSystemWatcher1.EnableRaisingEvents = false;
                avgMax = 0;
                sumMax = 0;
                //added with new variable array stuff to fix std dev 11-7
                peatMax = new double[arraysize1];
                minHFRpos = new int[arraysize1];
                maxMaxPos = new double[arraysize1];
                peat = new int[arraysize1];
                listMax = new double[arraysize1];
                list = new int[arraysize1];
                peat = new int[arraysize1];
            }
        }
        //fine V

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (NebCamera == "No camera")
                    NoCameraSelected();
                vcurveenable = 1;//added 1-7-12
                /* ***********remd for debugging 6-28
                if (roughvdone == false)//make sure rough v curve done to establish rough focus point
                {
                    DialogResult result2;
                    result2 = MessageBox.Show("Must perform a rough V-curve first", "scopefocus",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    if (result2 == DialogResult.OK)
                    {
                        return;
                    }
                }
                 */
                setarraysize();
                watchforOpenPort();
                if (portopen == 1)
                {
                    if ((textBox20.Text == "" || textBox18.Text == "") & (radioButton1.Checked == true))
                    {
                        DialogResult result;
                        result = MessageBox.Show("HFR Min/Max cannot be blank", "scopefocus",
                                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        if (result == DialogResult.OK)
                        {
                            textBox18.Focus();
                            return;
                        }
                    }
                    if (radioButton1.Checked == true)
                    {
                        enteredMaxHFR = Convert.ToInt16(textBox20.Text.ToString());
                        enteredMinHFR = Convert.ToInt16(textBox18.Text.ToString());
                    }
                    backlashDetermOn = false;
                    port.DiscardInBuffer();
                    fineVrepeatcounter();
                    chart1.Series[0].Points.Clear();
                    chart1.Series[1].Points.Clear();
                    chart1.Series[2].Points.Clear();

                    int finegoto = posMin - ((((int)numericUpDown3.Value) / 2) * ((int)numericUpDown8.Value));
                    //fine v-curve goes to N/2 * step size in from the focus position -- V should be centered
                    if ((finegoto - 100) < 0)
                    {
                        DialogResult result1 = MessageBox.Show("Goto Exceeds Full In\rAdditional backfocus needed to take-up backlash and center V-curve", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (result1 == DialogResult.OK)
                        {
                            return;
                        }
                    }
                    Log("Clearing backlash...");//added 1-7-12
                    gotopos(Convert.ToInt16(finegoto - 100));//take up backlash     
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 4));
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 3));
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 2));
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt16(finegoto - (int)numericUpDown8.Value));
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt16(finegoto));
                    Thread.Sleep(1000);
                    ffenable = 1;
                    repeatProgress = 0;
                    repeatDone = 0;
                    vProgress = 0;
                    vDone = 0;
                    repeatTotal = 0;
                    vN = 0;
                    sum = 0;
                    min = 500;
                    HFRarraymin = 999;
                    tempon = 0;
                    sumMax = 0;
                    avgMax = 0;
                    maxMax = 0;
                    apexHFR = 0;
                    vv = 0;
                    peatMax = new double[arraysize2];
                    minHFRpos = new int[arraysize2];
                    maxMaxPos = new double[arraysize2];
                    peat = new int[arraysize2];
                    listMax = new double[arraysize2];
                    list = new int[arraysize2];
                    peat = new int[arraysize2];
                    Array.Clear(listMax, 0, arraysize2);
                    //****added 11-14
                    arraycountleft = 0;
                    arraycountright = 0;
                    if (fineVrepeatOn == false)//added 1-7-12
                    {
                        fineVrepeatDone = 0;
                    }
                    fileSystemWatcher2.EnableRaisingEvents = true;
                    fileSystemWatcher5.EnableRaisingEvents = true;//added to test metric HFR
                    fileSystemWatcher3.EnableRaisingEvents = false;
                    fileSystemWatcher1.EnableRaisingEvents = false;
                }
            }
            catch(Exception ex)
            {
                Log("FineV Error - Line 2191" + ex.ToString());
                Send("FineV Error - Line 2191" + ex.ToString());
                FileLog("FineV Error - Line 2191" + ex.ToString());

            }
        }

        void fineVrepeatcounter()
        {
            if ((fineVrepeatOn == true) & (fineVrepeat > 1))
            {
                fineVrepeatDone++;
                fineVrepeat--;
            }
            if (fineVrepeatOn == false)
            {
                fineVrepeat = (int)numericUpDown10.Value;
                if (fineVrepeat > 1)
                {
                    fineVrepeatOn = true;
                }
            }
        }
//fwd button
        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                watchforOpenPort();
                if (portopen == 1)
                {
                    int step = (int)numericUpDown2.Value;
                    if ((count - step) < 0)
                    {
                        DialogResult result1 = MessageBox.Show("Goto Exceeds Full In", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (result1 == DialogResult.OK)
                        {
                            count = 0;
                            return;
                        }
                    }
                    gotopos(Convert.ToInt16(count - step));
                    textBox1.Text = count.ToString();
                    textBox4.Text = posMin.ToString();
                    positionbar();
                }
            }
            catch(Exception ex)
            {
                Log("Fwd Button Error" + ex.ToString());
            }
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            sync();
            count = syncval;
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (portopen == 1)
                {
                    fileSystemWatcher2.EnableRaisingEvents = false;
                    fileSystemWatcher5.EnableRaisingEvents = false; // added to test metricHFR
                    fileSystemWatcher3.EnableRaisingEvents = false;
                    int go2 = (int)numericUpDown6.Value;

                    gotopos(Convert.ToInt16(go2));
                    Thread.Sleep(20);
                }
            }
            catch(Exception ex)
            {
                Log("goto button error" + ex.ToString());
            }
        }
        //backlash button
        private void button10_Click_1(object sender, EventArgs e)
        {

        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            GetAvg();
            FillData();
            gotoFocus();
        }
       
 
        private void MainWindow_Shown(object sender, EventArgs e)
        {
            if (portautoselect == true)
            {
                comboBox1.SelectedIndex = 0;
                portselected = comboBox1.SelectedItem.ToString();
                button8.PerformClick();
                if (port != null)
                {
                    button8.BackColor = System.Drawing.Color.Lime;
                    this.button8.Text = "Connected";
                    playsound();
                }
            }
            else
            {
                comboBox1.Focus();
            }
        }


        private void button8_Click_2(object sender, EventArgs e)
        {
            try
            {
                if (portselected == null)
                {
                    MessageBox.Show("Port not selected", "scopefocus");
                    return;
                }
                if (port == null)
                {
                    PortOpen();
                    Thread.Sleep(500);
                }
                if (connect == 1)
                {
                    DialogResult result2 = MessageBox.Show("Already Connected - Resetting Arduino to Computer Mode", "Arduino scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

                connect = 1;
             //   WindowsFormsApplication1.Properties.Settings.Default.port = port.PortName.ToString();
            //    WindowsFormsApplication1.Properties.Settings.Default.Save();

                //*****try enter computer connected mode for arduino
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
                Thread.Sleep(10);
                port.Write("C");
                Thread.Sleep(50);
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
            }
            catch(Exception ex)
            {
                Log("Connect Button Error" + ex.ToString());
            }
        }
 
          
        //makes sure equipment is selected prior to page change
        private void tabControl1_SelectedIndexChanged_1(object sender, EventArgs e)
        {


        }

        //this works w/ sim.  need one extra file change since using delete
        private void fileSystemWatcher2_Deleted(object sender, FileSystemEventArgs e)
        {
            vcurve();
        }

        private void textBox11_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            path2 = folderBrowserDialog1.SelectedPath.ToString();
            textBox11.Text = path2.ToString();
            toolStripStatusLabel5.Text = path2.ToString();

        }

        //reset defaults 
        private void button14_Click(object sender, EventArgs e)
        {
            try
            {

                DialogResult result = MessageBox.Show("Reset to Defaults? ", "scopefocus",
                                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.Cancel)
                {
                    return;
                }
                else
                {
                    WindowsFormsApplication1.Properties.Settings.Default.path = @"c:\windows\temp";
                //    WindowsFormsApplication1.Properties.Settings.Default.port = @"COM";
                    WindowsFormsApplication1.Properties.Settings.Default.travel = 5750;
                    WindowsFormsApplication1.Properties.Settings.Default.NebPath = @"C:\Program Files (x86)\Nebulosity3";
                    //    WindowsFormsApplication1.Properties.Settings.Default.camera = "Simulator";
                    WindowsFormsApplication1.Properties.Settings.Default.user = "Enter user";
                    WindowsFormsApplication1.Properties.Settings.Default.pswd = "Enter pswd";
                    WindowsFormsApplication1.Properties.Settings.Default.server = "Enter server";
                    WindowsFormsApplication1.Properties.Settings.Default.to = "Enter to";
                    WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Clear();
                    WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Clear();
                    WindowsFormsApplication1.Properties.Settings.Default.Save();
                }
            }
            catch(Exception ex)
            {
                Log("Restore defaults - Error" + ex.ToString());
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            portselected = comboBox1.SelectedItem.ToString();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            //*****add for new equip comboboxes
            WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Clear();
            foreach (string Item in comboBox7.Items)
            {
                WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Add(Item);
            }
            WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Clear();
            foreach (string Item in comboBox8.Items)
            {
                WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Add(Item);
            }
            //******end addt.
         //   WindowsFormsApplication1.Properties.Settings.Default.port = portselected.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.path = path2.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.travel = travel;
            WindowsFormsApplication1.Properties.Settings.Default.NebPath = NebPath.ToString();
         //   WindowsFormsApplication1.Properties.Settings.Default.camera = camera;
            WindowsFormsApplication1.Properties.Settings.Default.user = this.user;
            WindowsFormsApplication1.Properties.Settings.Default.pswd = this.pswd;
            WindowsFormsApplication1.Properties.Settings.Default.server = this.server;
            WindowsFormsApplication1.Properties.Settings.Default.to = this.to;
            int i = comboBox7.SelectedIndex;
            if (i == 0)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel1 = travel;
            if (i == 1)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel2 = travel;
            if (i == 2)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel3 = travel;
            if (i == 3)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel4 = travel;
            if (i == 4)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel5 = travel;
            if (i == 5)
                WindowsFormsApplication1.Properties.Settings.Default.MaxTravel6 = travel;
            WindowsFormsApplication1.Properties.Settings.Default.Save();

           closing = 1;
            standby();
            //  gotopos(0);
            if (port == null)
            {
                System.Environment.Exit(0);
            }
            else
            {

                gotopos(0);
                port.Close();
            }
            if (phdsocket.Connected == true)
            phdsocket.Close();
            if (clientSocket != null)
            {
             //   clientSocket.GetStream().Close();//added 5-17-12
             //   clientSocket.Client.Disconnect(false);//added 5-17-12
                clientSocket.Close();
            }
            System.Environment.Exit(0);
        }
        //paypal for about tab
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            string url = "https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=KBV3Q26GZUTNL&lc=US&item_name=scopefocus%2einfo&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donateCC_LG%2egif%3aNonHosted";

            string business = "info@scopefocus.info";  //paypal email
            string description = "Donation";
            string country = "US";
            string currency = "USD";

            url += "https://www.paypal.com/cgi-bin/webscr" +
                "?cmd=" + "_donations" +
                "&business=" + business +
                "&lc=" + country +
                "&item_name=" + description +
                "&currency_code=" + currency +
                "&bn=" + "PP%2dDonationsBF";
            System.Diagnostics.Process.Start(url);
        }
        //changes cusor to hand making look more like a link for paypal link
        private void pictureBox1_MouseHover(object sender, EventArgs e)
        {
            this.pictureBox1.Cursor = Cursors.Hand;
        }
        //this makes red x in upper righ corner the same as close buttons
        private void MainWindow_FormClosing(object sender, FormClosingEventArgs e)
        {

            //*****add for new equip comboboxes
            WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Clear();
            foreach (string Item in comboBox7.Items)
            {
                WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Add(Item);
            }
            WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Clear();
            foreach (string Item in comboBox8.Items)
            {
                WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Add(Item);
            }
            //******end addt.

        //    WindowsFormsApplication1.Properties.Settings.Default.port = portselected.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.path = path2.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.travel = travel;
            WindowsFormsApplication1.Properties.Settings.Default.NebPath = NebPath.ToString();
           WindowsFormsApplication1.Properties.Settings.Default.user = this.user;
           WindowsFormsApplication1.Properties.Settings.Default.pswd = this.pswd;
           WindowsFormsApplication1.Properties.Settings.Default.server = this.server;
           WindowsFormsApplication1.Properties.Settings.Default.to = this.to;
         //   WindowsFormsApplication1.Properties.Settings.Default.camera = camera;
           int i = comboBox7.SelectedIndex;
           if (i == 0)
               WindowsFormsApplication1.Properties.Settings.Default.MaxTravel1 = travel;
           if (i == 1)
               WindowsFormsApplication1.Properties.Settings.Default.MaxTravel2 = travel;
           if (i == 2)
               WindowsFormsApplication1.Properties.Settings.Default.MaxTravel3 = travel;
           if (i == 3)
               WindowsFormsApplication1.Properties.Settings.Default.MaxTravel4 = travel;
           if (i == 4)
               WindowsFormsApplication1.Properties.Settings.Default.MaxTravel5 = travel;
           if (i == 5)
               WindowsFormsApplication1.Properties.Settings.Default.MaxTravel6 = travel;
            WindowsFormsApplication1.Properties.Settings.Default.Save();
            closing = 1;
            standby();
            if (port == null)
            {
                System.Environment.Exit(0);
            }
            else
            {
                gotopos(0);
                port.Close();
            }
            if (phdsocket != null)
                phdsocket.Close();
            if (clientSocket != null)
            {
                //clientSocket.GetStream().Close();//added 5-17-12
              //  clientSocket.Client.Disconnect(false);//added 5-17-12
                clientSocket.Close();
            }
        }
        private bool checkboxChanged = false;//allows for only 1 time fill of all checkboxes at 
        //startup, once on is changed--tab change no longer fills the boxes
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                EquipRefresh();
                
            }
            catch(Exception ex)
            {
                Log("TabContorl Changed - Error" + ex.ToString());
            }
            }
        private void EquipRefresh()
        {
            equipPrefix = comboBox7.Text + "_" + comboBox8.Text;
            if (comboBox8.Text == "NB")
            {
                if (checkboxChanged == false)
                    StdNB();
                equipPrefix = comboBox7.Text + "_" + "LRGB";//always start w/ lum for NB
            }
            if (radioButton9.Checked == true)
                equip = equipPrefix + "_E";
            if (radioButton10.Checked == true)
                equip = equipPrefix + " _R";
            else
                equip = equipPrefix;

            if (comboBox8.Text == "LRGB")
            {
                if (checkboxChanged == false)
                    StdLRGB();
            }
            if (checkBox22.Checked == true)
                equip = equip + "_Metric";

            //   textBox13.Text = equip.ToString();
            toolStripStatusLabel3.Text = equip.ToString();
            //next line filters gridview by equip
            ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";
            //   ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + textBox13.Text.ToString() + "'";


            if (equip == null)
            {
                MessageBox.Show("Must Select Equipment First", "scopefocus");
                tabControl1.SelectedTab = tabPage2;//page 2 is setup (1st tab)
            }
            // }
            if ((tabControl1.SelectedTab == tabPage5) & (port2selected == null))
            {
                if (comboBox6.Items.Count == 1)
                {
                    comboBox6.SelectedIndex = 0;
                    port2selected = comboBox6.SelectedItem.ToString();
                    button36.PerformClick();

                }
            }
            button13.PerformClick();

            //below give a focus position starting point based on v-curve data.
            if (startup == true)
            {
                using (SqlCeConnection con = new SqlCeConnection(conString))
                {
                    con.Open();
                    using (SqlCeCommand com4 = new SqlCeCommand("SELECT AVG(FocusPos) FROM table1 WHERE Equip = @equip", con))
                    {
                        com4.Parameters.AddWithValue("@equip", equip);
                        SqlCeDataReader reader4 = com4.ExecuteReader();
                        while (reader4.Read())
                        {
                            if (!reader4.IsDBNull(0))
                            {
                                int numb7 = reader4.GetInt32(0);
                                posMin = numb7;
                                textBox4.Text = numb7.ToString();//sets focus position to avg from data
                                numericUpDown6.Value = numb7;//sets goto value to avg focus position
                            }
                        }
                        reader4.Close();
                    }
                    con.Close();
                }
                startup = false;
            }
        }
        //deletes all SQL data for this Equip
        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                if (equip == null)
                {
                    MessageBox.Show("Must select Equipment", "scopefocus");
                    return;
                }
                DialogResult result = MessageBox.Show("Are you sure?\rThis will delete all data for this equipment ", "scopefocus",
                                   MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.Cancel)
                {
                    return;
                }
                if (result == DialogResult.OK)
                {
                    using (SqlCeConnection con = new SqlCeConnection(conString))
                    {
                        con.Open();
                        using (SqlCeCommand com = new SqlCeCommand("DELETE FROM table1 WHERE Equip =@Equip", con))
                        {

                            com.Parameters.Add(new SqlCeParameter("@equip", equip));
                            com.ExecuteNonQuery();
                        }
                        con.Close();
                    }
                    FillData();
                }
            }
            catch(Exception ex)
            {
                Log("Error Deleting SQL Data"+ ex.ToString());
            }
        }
        //View All button
        private void button17_Click(object sender, EventArgs e)
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
            }
            catch(Exception ex)
            {
                Log("View All SQL data Error" + ex.ToString());
            }

        }
        //this allows toggling of radio button must have autocheck property false
        private void radioButton1_Click(object sender, EventArgs e)
        {

            if (radioButton1.Checked == true)
            {
                radioButton1.Checked = false;
            }
            else
                radioButton1.Checked = true;

        }
        //next 2 ensure UP/DWN radio buttons are toggled
        private void radioButton3_Click(object sender, EventArgs e)
        {

            if (radioButton3.Checked == true)
            {
                radioButton2.Checked = false;
                radioButton3.Checked = true;
            }
            else
            {
                radioButton3.Checked = true;
                radioButton2.Checked = false;
            }

        }

        private void radioButton2_Click(object sender, EventArgs e)
        {

            if (radioButton2.Checked == true)
            {
                radioButton3.Checked = false;
                radioButton2.Checked = true;
            }
            else
            {
                radioButton2.Checked = true;
                radioButton3.Checked = false;
            }

        }
        //next 2 ensure the entered HFR min < max
        private void textBox20_Leave(object sender, EventArgs e)
        {
            enteredMaxHFR = Convert.ToInt16(textBox20.Text.ToString());
            if (textBox18.Text.ToString() != "")
            {
                if (enteredMaxHFR < enteredMinHFR)
                {
                    DialogResult result = MessageBox.Show("Maximum value must be greater than Minimum ", "scopefocus",
                                   MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (result == DialogResult.OK)
                    {
                        textBox20.Clear();
                        textBox20.Focus();
                        return;
                    }
                }
            }
        }

        private void textBox18_Leave(object sender, EventArgs e)
        {
            enteredMinHFR = Convert.ToInt16(textBox18.Text.ToString());
            if (textBox20.Text.ToString() != "")
            {
                if (enteredMinHFR > enteredMaxHFR)
                {
                    DialogResult result = MessageBox.Show("Minimum value must be less than Maximum ", "scopefocus",
                                   MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (result == DialogResult.OK)
                    {
                        textBox18.Clear();
                        textBox18.Focus();
                        return;
                    }
                }
            }
        }
        //allow for use of enter key after text input, goes to next field (TAB) 
        protected override bool ProcessDialogKey(Keys keyData)
        {
            switch (keyData)
            {
                case Keys.Enter:
                case Keys.Space:
                    return base.ProcessDialogKey(Keys.Tab);
            }
            return base.ProcessDialogKey(keyData);
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            //will try to allow for deletion of selected rows

        }
        //determine backlash
        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                backlash = 0;
                backlashSum = 0;
                backlashCount = 0;
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                chart1.Series[2].Points.Clear();

                sum = 0;
                avg = 0;
                vDone = 0;
                vProgress = 0;
                tempon = 0;
                tempsum = 0;
                vv = 0;
                templog = 0;
                backlashDetermOn = true;
                radioButton1.Checked = false;//calculations off
                ffenable = 1;
                int finegoto = posMin - ((((int)numericUpDown3.Value) / 2) * ((int)numericUpDown8.Value));
                //fine v-curve goes to N/2 * step size in from the focus position -- V should be centered
                if ((finegoto - 100) < 0)
                {
                    DialogResult result1 = MessageBox.Show("Goto Exceeds Full In/rAdditional backfocus needed to take-up backlash and center V-curve", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (result1 == DialogResult.OK)
                    {
                        return;
                    }
                }
                gotopos(Convert.ToInt16(finegoto - 100));//take up backlash  
                Thread.Sleep(1000);
                gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 4));
                Thread.Sleep(1000);
                gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 3));
                Thread.Sleep(1000);
                gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 2));
                Thread.Sleep(1000);
                gotopos(Convert.ToInt16(finegoto - (int)numericUpDown8.Value));
                Thread.Sleep(1000);
                gotopos(Convert.ToInt16(finegoto));
                backlashOUT = true;//identifies current v curve direction is going out(reverse)
                fileSystemWatcher2.EnableRaisingEvents = false;
                fileSystemWatcher5.EnableRaisingEvents = false;//added to test metricHFR
                fileSystemWatcher3.EnableRaisingEvents = false;

                fileSystemWatcher1.EnableRaisingEvents = true;
            }
            catch(Exception ex)
            {
                Log("Backlash error" + ex.ToString());
                FileLog("Backlash Error" + ex.ToString());
            }
        }
        //tempcal button
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                chart1.Series[2].Points.Clear();

                //  min = 1;
                sum = 0;
                avg = 0;
                vDone = 0;
                vProgress = 0;
                tempon = 1;
                tempsum = 0;
                vv = 0;
                templog = 0;
                peatMax = new double[arraysize2];
                minHFRpos = new int[arraysize2];
                maxMaxPos = new double[arraysize2];
                peat = new int[arraysize2];
                listMax = new double[arraysize2];
                list = new int[arraysize2];
                peat = new int[arraysize2];
                Array.Clear(listMax, 0, arraysize2);
                int finegoto = posMin - ((((int)numericUpDown3.Value) / 2) * ((int)numericUpDown8.Value));
                //fine v-curve goes to N/2 * step size in from the focus position -- V should be centered
                if ((finegoto - 100) < 0)
                {
                    DialogResult result1 = MessageBox.Show("Goto Exceeds Full In", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (result1 == DialogResult.OK)
                    {
                        return;
                    }
                }
                gotopos(Convert.ToInt16(finegoto - 100));//take up backlash  
                Thread.Sleep(1000);
                gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 4));
                Thread.Sleep(1000);
                gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 3));
                Thread.Sleep(1000);
                gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value) * 2));
                Thread.Sleep(1000);
                gotopos(Convert.ToInt16(finegoto - (int)numericUpDown8.Value));
                Thread.Sleep(1000);
                fileSystemWatcher2.EnableRaisingEvents = false;
                fileSystemWatcher5.EnableRaisingEvents = false; //added to test metricHFR
                fileSystemWatcher3.EnableRaisingEvents = false;
                fileSystemWatcher1.EnableRaisingEvents = true;
                //  tempcal();
            }
            catch(Exception ex)
            {
                Log("TempCal Error" + ex.ToString());
                Send("TempCal Error" + ex.ToString());
                FileLog("TempCal Error" + ex.ToString());

            }
        }
        //deletes selected rows
        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                selectedrowcount = dataGridView1.SelectedCells.Count;
                if (selectedrowcount == 1)
                {
                    selectedcell = dataGridView1.SelectedCells[0].Value.ToString();
                    using (SqlCeConnection con = new SqlCeConnection(conString))
                    {
                        con.Open();
                        using (SqlCeCommand com = new SqlCeCommand("DELETE FROM table1 WHERE Number =@number", con))
                        {
                            int delrow = Convert.ToInt32(selectedcell);
                            com.Parameters.Add(new SqlCeParameter("@number", delrow));
                            com.ExecuteNonQuery();
                        }
                        con.Close();
                    }
                    FillData();
                }
                else//setup array for all selected cell values
                {
                    for (int x = 0; x < selectedrowcount; x++)
                    {
                        selectedcell = dataGridView1.SelectedCells[0].Value.ToString();//because one gets deleted every x, index [0] deletes the next one
                        using (SqlCeConnection con = new SqlCeConnection(conString))
                        {
                            con.Open();
                            using (SqlCeCommand com = new SqlCeCommand("DELETE FROM table1 WHERE Number =@number", con))
                            {
                                int delrow = Convert.ToInt32(selectedcell);
                                com.Parameters.Add(new SqlCeParameter("@number", delrow));
                                com.ExecuteNonQuery();
                            }
                            con.Close();
                        }
                        FillData();
                    }
                }
            }
            catch(Exception ex)
            {
                Log("Delete Selected SQL data Error" + ex.ToString());
            }

        }
        //debugfilteradvance
        private void filteradvance()
        {
            try
            {
                FlatCalcDone = false;
                if (SequenceRunning == true)
                {
                    DisableUpDwn();
                   // fileSystemWatcher1.EnableRaisingEvents = true;
                }
                    toolStripStatusLabel1.Text = "Filter Moving";
                    this.Refresh();
                    string movedone;
                    port.DiscardOutBuffer();
                    port.DiscardInBuffer();
                    Thread.Sleep(20);
                    port.Write("A");
                    filterMoving = true;
                    Thread.Sleep(50);

                    while (filterMoving == true)
                    {
                        movedone = port.ReadExisting();
                        if (movedone == "D")
                        {

                            filterMoving = false;
                            
                            toolStripStatusLabel1.Text = "Capturing";
                            this.Refresh();
                            //   toolStripStatusLabel1.Text = " ";
                            //    toolStripStatusLabel1.Text = " ";

                        }
                    }
                    Thread.Sleep(50);
                    port.DiscardInBuffer();
                
                if (currentfilter == 0)//allows advance without count to set to starting pos. 
                {
                    return;
                }
                /* *********************rem'd 4-11 put currentfilter asignment in filtersequence()
                // if (checkBox5.Checked == false)
                //  {
                if (currentfilter != 7)//***************changed from 4 on 2_29 to ? fix dark2
                {
                    currentfilter++;
                }
                else
                    currentfilter = 1;//reset at end
*/
                //    toolStripStatusLabel1.Text = " ";
                DisplayCurrentFilter();
            }
            catch
            {
                Log("Filter Advance Error - Make Sure Arduino Connected");
                Send("Filter Advance Error - Make Sure Arduino Connected");
                FileLog("Filter Advance Error - Make Sure Arduino Connected");

            }
        }

private void DisplayCurrentFilter()
{
    try
    {
        //  if (SequenceRunning == true)
        //  DisableUpDwn();
        //    }
        /*
                    if (checkBox5.Checked == true)
                    {
                        if (currentfilter != 5)
                        {
                            currentfilter++;
                        }
                        else
                            currentfilter = 1;
                    }
         */
        if (currentfilter == 1)
            Filtertext = comboBox2.Text;
        if (currentfilter == 2)
            Filtertext = comboBox3.Text;
        if (currentfilter == 3)
            Filtertext = comboBox4.Text;
        if (currentfilter == 4)
            Filtertext = comboBox5.Text;
        if (DarksOn == true)
        {
            Filtertext = "Dark";
            textBox21.BackColor = System.Drawing.Color.Black;
            textBox21.ForeColor = System.Drawing.Color.White;
            toolStripStatusLabel4.ForeColor = System.Drawing.Color.White;
            toolStripStatusLabel4.BackColor = System.Drawing.Color.Black;
        }
        if (FlatsOn == true)//****************added 4_5 for flats
        {
            Filtertext = "Flat";
            textBox21.BackColor = System.Drawing.Color.White;
            textBox21.ForeColor = System.Drawing.Color.Blue;
            toolStripStatusLabel4.ForeColor = System.Drawing.Color.Blue;
            toolStripStatusLabel4.BackColor = System.Drawing.Color.White;
        }
        if (Filtertext == "Blue")
        {
            textBox21.BackColor = System.Drawing.Color.Blue;
            textBox21.ForeColor = System.Drawing.Color.White;
            toolStripStatusLabel4.ForeColor = System.Drawing.Color.White;
            toolStripStatusLabel4.BackColor = System.Drawing.Color.Blue;
        }
        /*
                if ((Filtertext != "Blue") || (Filtertext != "Dark"))
                    textBox21.ForeColor = System.Drawing.Color.White;
        */
        if ((DarksOn == false) && (Filtertext != "Blue"))
        {
            textBox21.ForeColor = System.Drawing.Color.Black;
            toolStripStatusLabel4.ForeColor = System.Drawing.Color.Black;
        }

        if (Filtertext == "Luminosity")
        {
            textBox21.BackColor = System.Drawing.Color.White;
            toolStripStatusLabel4.BackColor = System.Drawing.Color.White;
            
        }
        if (Filtertext == "Green")
        {
            textBox21.BackColor = System.Drawing.Color.Lime;
            toolStripStatusLabel4.BackColor = System.Drawing.Color.Lime;
        }

        if (Filtertext == "Red")
        {
            textBox21.BackColor = System.Drawing.Color.Red;
            toolStripStatusLabel4.BackColor = System.Drawing.Color.Red;
        }


        if (Filtertext == "Ha")
        {
            textBox21.BackColor = System.Drawing.Color.Magenta;
            toolStripStatusLabel4.BackColor = System.Drawing.Color.Magenta;
        }
        if (Filtertext == "OIII")
        {
            textBox21.BackColor = System.Drawing.Color.Cyan;
            toolStripStatusLabel4.BackColor = System.Drawing.Color.Cyan;
        }
        if (Filtertext == "SII")
        {
            textBox21.BackColor = System.Drawing.Color.Yellow;
            toolStripStatusLabel4.BackColor = System.Drawing.Color.Yellow;
        }
        //Added to sync filter with equip
        if ((Filtertext == "Ha") || (Filtertext == "OIII") || (Filtertext == "SII"))
            equipPrefix = comboBox7.Text + "_" + Filtertext;
        if (radioButton9.Checked == true)
            equip = equipPrefix + "_E";
        if (radioButton10.Checked == true)
            equip = equipPrefix + " _R";
        else
            equip = equipPrefix;
            if ((Filtertext == "Luminosity")) //  rem'd 4_19 || (Filtertext == "Red") || (Filtertext == "Green") || (Filtertext == "Blue"))
       //only need this to start at lum w/ NB
            {
            equipPrefix = comboBox7.Text + "_" + "LRGB";
            if (radioButton9.Checked == true)
                equip = equipPrefix + "_E";
            if (radioButton10.Checked == true)
                equip = equipPrefix + " _R";
            else
                equip = equipPrefix;
        }
            if (checkBox22.Checked == true)
                equip = equip + "_Metric";
            toolStripStatusLabel3.Text = equip.ToString();
        //end filter/equip sync addition
        toolStripStatusLabel4.Text = Filtertext.ToString();
        toolStripStatusLabel2.Text = "Sub " + subCountCurrent.ToString() + "of " + totalsubs.ToString();
        //next line filters gridview by equip
        ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";

        //  FillData();
        //   GetAvg();
        //   textBox27.Text = Filtertext;
        textBox21.Text = Filtertext.ToString();
      
   }
    catch
   {
        Log("Display Current Filter Error");
    }
        }

        private int FocusPerSubGroupCount;
        void checkfiltercount()//check current sub count and filter position, initiates first capture(subcount = 0)
        {
          //  try
         //   {
                //first 3 line duplicate w/ radio 7 enable, may be able to delete


                //  subsperfilter = totalsubs / filternumber;
                //textBox19.Text = subsperfilter.ToString();
                //   textBox27.Text = Filtertext;
                //    textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;

                textBox23.Text = totalsubs.ToString();
                if (subCountCurrent == 0)//starts for first one
                {
                    NebCapture();
                }
                if ((subCountCurrent == totalsubs) & (subCountCurrent != 0))//signifies ens of sequence
                {
                    //adioButton7.Checked = false;
                    fileSystemWatcher4.EnableRaisingEvents = false;
                    //NetworkStream serverStream = clientSocket.GetStream();
                    serverStream = clientSocket.GetStream();
                    byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                    serverStream.Write(outStream, 0, outStream.Length);
                    serverStream.Flush();
                    toolStripStatusLabel1.Text = "Sequence Done";
                    this.Refresh();
                    if (DarksOn == true)
                        DarksOn = false;
                    currentfilter = 1;
                    DisplayCurrentFilter();
                    //      currentfilter = 0;
                    subCountCurrent = 0;
                    filterCountCurrent = 0;
                    Thread.Sleep(3000);//prevent overlapping sounds
                    serverStream.Close();
                    SetForegroundWindow(NebhWnd);
                    Thread.Sleep(1000);
                    PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                    Thread.Sleep(1000);
                    NebListenOn = false;
                   // clientSocket.GetStream().Close();//added 5-17-12
                  //  clientSocket.Client.Disconnect(true);//added 5-17-12
                    clientSocket.Close();
                    toolStripStatusLabel4.BackColor = System.Drawing.Color.LightGray;
                    toolStripStatusLabel4.Text = Filtertext.ToString();
                    button12.UseVisualStyleBackColor = true;
                    button31.UseVisualStyleBackColor = true;
                    if (checkBox11.Checked == true)
                    {
                        StopPHD();
                        StopTracking();
                    }
                    button26.UseVisualStyleBackColor = true;
                    SequenceRunning = false;
                    EnableAllUpDwn();
                    //close slave
                    if (IsServer())
                    {
                        Log("Waiting for slave sequence done");
                        toolStripStatusLabel1.Text = "Waiting for Slave Sequence";
                        this.Refresh();
                        while (working == true)
                        {
                            WaitForSequenceDone("Sequence done", NebSlavehwnd);
                           // Thread.Sleep(50);
                        }
                        working = true;
                        Log("Slave sequence done");
                        SlavePause();//close slave
                      
                    }
                   
                   //      clientSocket.Close();
                    playsound();

                    return;
                }
                if (filterCountCurrent == subsperfilter)
                {
                    if (IsServer())
                    {
                        
 
                        Log("Waiting for slave sequence done");
                        toolStripStatusLabel1.Text = "Waiting for slave sequence";
                        this.Refresh();
                            while (working == true)
                            {                   
                                WaitForSequenceDone("Sequence done", NebSlavehwnd);
                               // Thread.Sleep(50);
                            }
                        working = true;
                        Log("Slave sequence done");
                       


                    }
                    //reset sub count for given filter
                    filterCountCurrent = 0;
                    FilterSequence();

                }
                
                //add to do sequence for refocus per subs
                if (checkBox17.Checked == true)
                if (FocusPerSubCurrent == FocusPerSub)
                {
                   
                    //then need to rename subsperfilter for nebcapture and repeat it d1,d2,d3,d4 times
                    //keep track of subsperfocus 1,2,3,4 may not be needed as ? same as focuspersub
                    FocusPerSubGroupCount++;
                    
                    // FilterFocus();
                }

          //  }
                /*
            catch (Exception e)
            {
                Log("CheckFlterCount Error" + e.ToString());
                Send("CheckFilterCount Error" + e.ToString());
                FileLog("CheckFilterCount Error" + e.ToString());

            }
*/
        }

     

        private void fileSystemWatcher4_Changed(object sender, FileSystemEventArgs e)
        {
            //if (FlatDone == false)
          //  {

                textBox41.Refresh();
                textBox41.Clear();
                if (!IsSlave())
                {
                    subCountCurrent++;
                    filterCountCurrent++;//the sub per this filter number
                    ToolStripProgressBar();
                    checkfiltercount();

                    if (checkBox17.Checked == true)
                        FocusPerSubCurrent++;
                }
            /*
                if (IsSlave())
                {
                    SendtoServer("Sub Complete");//will have the server wait for this.
                    Thread.Sleep(500);
                    textBox41.Refresh();
                }
            */
               // Thread.Sleep(500);
                //  }
        }
        //step filter forward
        private void FilterStepFwd()
        {
            try
            {
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
                Thread.Sleep(20);
                port.Write("F");
                Thread.Sleep(200);
            }
            catch(Exception ex)
            {
                Log("Filter Fwd Error" + ex.ToString());
            }
        }


        private void filterStepRev()
        {
            try
            {
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
                Thread.Sleep(20);
                port.Write("R");
                Thread.Sleep(200);
            }
            catch(Exception ex)
            {
                Log("Filter rev Error" + ex.ToString());
            }
        }


        private void button19_Click(object sender, EventArgs e)
        {
            FilterStepFwd();

        }
        
        
        //filter advance on filter tab
        private void button20_Click(object sender, EventArgs e)
        {
            if (currentfilter == 0)//if unsynced just advance
                filteradvance();
            if ((currentfilter != 0) && (currentfilter !=4))//below allows display of current filter while manually changing
            {
                currentfilter++;
                filteradvance();
                Update();
                return;
            }
            if (currentfilter == 4)
            {
                currentfilter = 1;
                filteradvance();
                Update();
            }
        }
      
        //set max travel on setup tab
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            travel = Convert.ToInt16(textBox2.Text.ToString());

        }
        //std LRGB

        private void StdLRGB()
        {
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 1;
            comboBox4.SelectedIndex = 2;
            comboBox5.SelectedIndex = 3;
            checkBox1.Checked = true;
            checkBox2.Checked = true;
            checkBox3.Checked = true;
            checkBox4.Checked = true;
            checkBox5.Checked = true;
        }

       
        private void StdNB()
        {
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 4;
            comboBox4.SelectedIndex = 5;
            comboBox5.SelectedIndex = 6;
            checkBox2.Checked = true;
            checkBox3.Checked = true;
            checkBox4.Checked = true;
            checkBox5.Checked = true;
        }

        //add filter
        private void button24_Click(object sender, EventArgs e)
        {
            string newitem = textBox5.Text.ToString();
            comboBox2.Items.Add(newitem);
            comboBox3.Items.Add(newitem);
            comboBox4.Items.Add(newitem);
            comboBox5.Items.Add(newitem);
            textBox5.Clear();
        }

        private void CountFilterTotal()//this may not be needed.  seems to just check for non-zero at sequencego
        {
            int filter6used;
            if (checkBox1.Checked == true)
                filter1used = 1;
            else
                filter1used = 0;
            if (checkBox2.Checked == true)
                filter2used = 1;
            else
                filter2used = 0;
            if (checkBox3.Checked == true)
                filter3used = 1;
            else
                filter3used = 0;
            if (checkBox4.Checked == true)
                filter4used = 1;
            else
                filter4used = 0;
            if (checkBox5.Checked == true)
                filter5used = 1;
            else
                filter5used = 0;
            if (checkBox9.Checked == true)
                 filter6used = 1;
            else
                filter6used = 0;

            filternumber = filter4used + filter3used + filter2used + filter1used + filter5used + filter6used;//calculate total number of filters

        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            CountFilterTotal();
            if (checkBox1.Checked == false)//put zeros in if unchecked
            {
                numericUpDown11.Value = 0;
                numericUpDown12.Value = 0;
            }
            checkboxChanged = true;//prevents re checking all boxes based on start up selection of filter set..
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            CountFilterTotal();
            if (checkBox2.Checked == false)
            {
                numericUpDown13.Value = 0;
                numericUpDown16.Value = 0;
            }
            checkboxChanged = true;
        }
      
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            CountFilterTotal();
            if (checkBox3.Checked == false)
            {
                numericUpDown14.Value = 0;
                numericUpDown17.Value = 0;
            }
            checkboxChanged = true;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            CountFilterTotal();
            if (checkBox4.Checked == false)
            {
                numericUpDown15.Value = 0;
                numericUpDown18.Value = 0;
            }
            checkboxChanged = true;
        }
        //when subsperfilter == filtercountcurrent...this tells what filter next and starts capture again. 
        //filtersequencehere
        private void FilterSequence()
        {
           // try
           // {
                DisplayCurrentFilter(); //try *****  added 5-23  
             //   textBox38.Text = currentfilter.ToString();
                if (currentfilter == 1)
                {
                    if (FilterFocusGroupCurrent < FocusGroup[0])
                    {
                        FilterFocus();
                        subsperfilter = SubsPerFocus[0];
                        FilterFocusGroupCurrent++;
                        Nebname = comboBox2.Text + "_" + FilterFocusGroupCurrent.ToString();
                        NebCapture();
                        return;//may not be needed???
                    }
                    else
                    FilterFocusGroupCurrent = 0;//must be equal to nuber of groups so reset and move to next filter
                    //should continue to cycle here until FFGC + FG[1]
                    if (checkBox18.Checked == true)
                    {
                        CaptureTime3 = FlatExp;
                        CaptureBin = (int)numericUpDown24.Value;
                        DoFlat();
                    }
                    if (checkBox2.Checked == true)
                    {
                        currentfilter = 2;
                        filteradvance();
                   //     FlatDone = false;
                   //     fileSystemWatcher1.EnableRaisingEvents = true;
                        if (checkBox8.Checked == true)
                        {

                            FilterFocus();
                        }
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown13.Value;
                            Nebname = comboBox3.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[1];
                            Nebname = comboBox3.Text + "_1";
                        }

                        // textBox21.Text = currentfilter.ToString();
                        //  textBox27.Text = currentfilter.ToString();
                      //  subsperfilter = (int)numericUpDown13.Value;
                      //  Nebname = comboBox3.Text.ToString();

                        CaptureTime = (int)numericUpDown16.Value * 1000;
                        CaptureBin = (int)numericUpDown25.Value;
                        //Thread.Sleep(4000);
                        //   if (FilterFocusOn == false)
                        if (!IsSlave())//*****added 6-17
                        {
                            if (checkBox8.Checked == false)//refocus after filter change check box
                                NebCapture();
                        }
                            return;
                        
                    }
                    if (checkBox3.Checked == true)
                    {
                        currentfilter = 3;
                        filteradvance();
                        filteradvance();
                     //   FlatDone = false;
                        //  textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = currentfilter.ToString();

                        //hread.Sleep(4000);
                        if (checkBox8.Checked == true)
                        {
                            
                            FilterFocus();
                        }
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown14.Value;
                            Nebname = comboBox4.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[2];
                            Nebname = comboBox4.Text + "_1";
                        }
                       // subsperfilter = (int)numericUpDown14.Value;
                      //  Nebname = comboBox4.Text.ToString();
                        CaptureTime = (int)numericUpDown17.Value * 1000;
                        CaptureBin = (int)numericUpDown26.Value;
                        //   if (FilterFocusOn == false)
                        if (checkBox8.Checked == false)
                            NebCapture();
                        
                        //   Thread.Sleep(4000);
                        return;
                    }
                    if (checkBox4.Checked == true)
                    {
                        currentfilter = 4;
                        filteradvance();
                        filteradvance();
                        filteradvance();
                    //    FlatDone = false;
                        //   textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = currentfilter.ToString();
                        if (checkBox8.Checked == true)
                            FilterFocus();
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown15.Value;
                            Nebname = comboBox5.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[3];
                            Nebname = comboBox5.Text + "_1";
                        }

                        // Thread.Sleep(4000);
                      //  subsperfilter = (int)numericUpDown15.Value;
                      //  Nebname = comboBox5.Text.ToString();
                        CaptureTime = (int)numericUpDown18.Value * 1000;
                        CaptureBin = (int)numericUpDown27.Value;
                        //   if (FilterFocusOn == false)
                    if (checkBox8.Checked == false)
                        NebCapture();
                      
                        return;
                    }
                    if (checkBox5.Checked == true)
                    {
                        currentfilter = 5;
                        DarksOn = true;
                       // filteradvance();//back to pos 1 
                        Thread.Sleep(1000);
                        editfound = 0;
                        Nebname = "Dark1";
                        subsperfilter = (int)numericUpDown20.Value;
                        CaptureTime = (int)numericUpDown19.Value * 1000;
                        CaptureBin = (int)numericUpDown28.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();
                        
                        return;
                    }
                    if (checkBox9.Checked == true)//Dark 2 allows for different time or bin
                    {
                        currentfilter = 6;
                        DarksOn = true;
                        //no advance already at 1
                        Thread.Sleep(1000);
                        editfound = 0;
                        Nebname = "Dark2";
                        subsperfilter = (int)numericUpDown29.Value;
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();
                       
                        return;
                    }
                    if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                    {
                        
                        currentfilter = 7;
                        if (FlatsOn == false)
                        ToggleFlat();
                        //filteradvance();//back to pos 1 
                        Thread.Sleep(1000);
                        editfound = 0;
                        Nebname = "Flat1";
                        subsperfilter = (int)numericUpDown32.Value;
                        CaptureTime = (int)numericUpDown34.Value;
                        CaptureBin = (int)numericUpDown36.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                       
                        if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                        {
                            CalculateFlatExp();
                            return;
                        }
                        
                        else
                            NebCapture();
                       
                        return;
                        
                    }
                     
                }
                
                    if (currentfilter == 2)
                    {
                        if (FilterFocusGroupCurrent < FocusGroup[1])//continues here for focusperfilter
                        {
                           
                            FilterFocus();
                            subsperfilter = SubsPerFocus[1];
                            FilterFocusGroupCurrent++;
                            Nebname = comboBox3.Text + "_" + FilterFocusGroupCurrent.ToString();
                            NebCapture();
                            return;//may not be needed???
                        }
                        else
                            FilterFocusGroupCurrent = 0;//must be equal to nuber of groups so reset and move to next filter
                        //should continue to cycle here until FFGC + FG[1]
                        if (checkBox18.Checked == true)
                        {
                            CaptureTime3 = FlatExp;
                            CaptureBin = (int)numericUpDown25.Value;
                            DoFlat();
                        }
                        if (checkBox3.Checked == true)
                        {

                            currentfilter = 3;
                            filteradvance();
                         //   FlatDone = false;
                            //   textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                            //    textBox27.Text = Filtertext;
                            if (checkBox8.Checked == true)
                                FilterFocus();
                            if (checkBox17.Checked == false)
                            {
                                subsperfilter = (int)numericUpDown14.Value;
                                Nebname = comboBox4.Text;
                            }
                            else
                            {
                                FilterFocusGroupCurrent = 1;
                                subsperfilter = SubsPerFocus[2];
                                Nebname = comboBox4.Text + "_1";
                            }
                            // subsperfilter = (int)numericUpDown14.Value;
                            CaptureTime = (int)numericUpDown17.Value * 1000;
                            CaptureBin = (int)numericUpDown26.Value;
                            //  Nebname = comboBox4.Text.ToString();
                            //    if (FilterFocusOn == false)
                            if (checkBox8.Checked == false)
                                NebCapture();
                            // Thread.Sleep(4000);
                            return;
                        }
                        if (checkBox4.Checked == true)
                        {

                            currentfilter = 4;
                            filteradvance();
                            filteradvance();
                        //    FlatDone = false;
                            // textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                            //   textBox27.Text = Filtertext;
                            if (checkBox8.Checked == true)
                                FilterFocus();
                            if (checkBox17.Checked == false)
                            {
                                subsperfilter = (int)numericUpDown15.Value;
                                Nebname = comboBox5.Text;
                            }
                            else
                            {
                                FilterFocusGroupCurrent = 1;
                                subsperfilter = SubsPerFocus[3];
                                Nebname = comboBox5.Text + "_1";
                            }
                            //  subsperfilter = (int)numericUpDown15.Value;
                            //  Nebname = comboBox5.Text.ToString();
                            CaptureTime = (int)numericUpDown18.Value * 1000;
                            CaptureBin = (int)numericUpDown27.Value;
                            //   if (FilterFocusOn == false)
                            if (checkBox8.Checked == false)
                                NebCapture();
                            //   Thread.Sleep(4000);
                            return;
                        }
                        if (checkBox5.Checked == true)
                        {
                            currentfilter = 5;
                            filteradvance();
                            filteradvance();
                            filteradvance();
                            DarksOn = true;
                            // filteradvance();//back to pos 1 
                            Thread.Sleep(1000);
                            editfound = 0;
                            Nebname = "Dark1";
                            subsperfilter = (int)numericUpDown20.Value;
                            CaptureTime = (int)numericUpDown19.Value * 1000;
                            CaptureBin = (int)numericUpDown28.Value;
                            FilterFocusOn = false;
                            //  fileSystemWatcher4.EnableRaisingEvents = true;
                            NebCapture();
                            return;
                        }
                        if (checkBox9.Checked == true)//Dark 2 allows for different time or bin
                        {
                            currentfilter = 6;
                            DarksOn = true;
                            //no advance already at 1
                            Thread.Sleep(1000);
                            editfound = 0;
                            Nebname = "Dark2";
                            subsperfilter = (int)numericUpDown29.Value;
                            CaptureTime = (int)numericUpDown30.Value * 1000;
                            CaptureBin = (int)numericUpDown31.Value;
                            FilterFocusOn = false;
                            //  fileSystemWatcher4.EnableRaisingEvents = true;
                            NebCapture();
                            return;
                        }
                        if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                        {
                            currentfilter = 7;
                            filteradvance();
                            filteradvance();
                            filteradvance();
                            if (FlatsOn == false)
                                ToggleFlat();
                            //  filteradvance();//back to pos 1 
                            Thread.Sleep(1000);
                            editfound = 0;
                            Nebname = "Flat1";
                            subsperfilter = (int)numericUpDown32.Value;
                            CaptureTime = (int)numericUpDown34.Value;
                            CaptureBin = (int)numericUpDown36.Value;
                            FilterFocusOn = false;
                            //  fileSystemWatcher4.EnableRaisingEvents = true;
                            if ((FlatCalcDone == false) & (checkBox15.Checked == true)) //15 is flat autoexp calc
                            {
                                CalculateFlatExp();
                                return;
                            }

                            else
                                NebCapture();
                            return;
                        }

                    
                }
                
                    if (currentfilter == 3)
                    {
                        if (FilterFocusGroupCurrent < FocusGroup[2])
                        {
                            FilterFocus();
                            subsperfilter = SubsPerFocus[2];
                            FilterFocusGroupCurrent++;
                            Nebname = comboBox4.Text + "_" + FilterFocusGroupCurrent.ToString();
                            NebCapture();
                            return;//may not be needed???
                        }
                        else
                            FilterFocusGroupCurrent = 0;//must be equal to nuber of groups so reset and move to next filter
                        //should continue to cycle here until FFGC + FG[1]
                        if (checkBox18.Checked == true)
                        {
                            CaptureTime3 = FlatExp;
                            CaptureBin = (int)numericUpDown26.Value;
                            DoFlat();
                        }
                        if (checkBox4.Checked == true)
                        {
                            currentfilter = 4;
                            filteradvance();
                         //   FlatDone = false;
                            if (checkBox8.Checked == true)
                                FilterFocus();

                            if (checkBox17.Checked == false)
                            {
                                subsperfilter = (int)numericUpDown15.Value;
                                Nebname = comboBox5.Text;
                            }
                            else
                            {
                                FilterFocusGroupCurrent = 1;
                                subsperfilter = SubsPerFocus[3];
                                Nebname = comboBox5.Text + "_1";
                            }
                            //   textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                            //     textBox27.Text = Filtertext;
                            //  subsperfilter = (int)numericUpDown15.Value;
                            //   Nebname = comboBox5.Text.ToString();
                            CaptureTime = (int)numericUpDown18.Value * 1000;
                            CaptureBin = (int)numericUpDown27.Value;
                            if (checkBox8.Checked == false)
                                NebCapture();
                            //Thread.Sleep(4000);
                            return;
                        }
                        if (checkBox5.Checked == true)
                        {
                            currentfilter = 5;
                            filteradvance();
                            filteradvance();
                            DarksOn = true;
                            filteradvance();//back to pos 1 
                            Thread.Sleep(1000);
                            editfound = 0;
                            Nebname = "Dark1";
                            subsperfilter = (int)numericUpDown20.Value;
                            CaptureTime = (int)numericUpDown19.Value * 1000;
                            CaptureBin = (int)numericUpDown28.Value;
                            FilterFocusOn = false;
                            //  fileSystemWatcher4.EnableRaisingEvents = true;
                            NebCapture();
                            return;
                        }
                        if (checkBox9.Checked == true)//Dark 2 allows for different time or bin
                        {
                            currentfilter = 6;
                            DarksOn = true;
                            //no advance already at 1
                            Thread.Sleep(1000);
                            editfound = 0;
                            Nebname = "Dark2";
                            subsperfilter = (int)numericUpDown29.Value;
                            CaptureTime = (int)numericUpDown30.Value * 1000;
                            CaptureBin = (int)numericUpDown31.Value;
                            FilterFocusOn = false;
                            //  fileSystemWatcher4.EnableRaisingEvents = true;
                            NebCapture();
                            return;
                        }
                        if ((checkBox13.Checked == true) & ( checkBox18.Checked==false))
                        {
                            currentfilter = 7;
                            filteradvance();
                            filteradvance();
                            if (FlatsOn == false)
                                ToggleFlat();
                            //  filteradvance();//back to pos 1 
                            Thread.Sleep(1000);
                            editfound = 0;
                            Nebname = "Flat1";
                            subsperfilter = (int)numericUpDown32.Value;
                            CaptureTime = (int)numericUpDown34.Value;
                            CaptureBin = (int)numericUpDown36.Value;
                            FilterFocusOn = false;
                            //  fileSystemWatcher4.EnableRaisingEvents = true;
                            if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                            {
                                CalculateFlatExp();
                                return;
                            }

                            else
                                NebCapture();
                            return;
                        }



                    
                }
               
                    if (currentfilter == 4)
                    {
                        if (FilterFocusGroupCurrent < FocusGroup[3])
                        {
                            FilterFocus();
                            subsperfilter = SubsPerFocus[3];
                            FilterFocusGroupCurrent++;
                            Nebname = comboBox5.Text + "_" + FilterFocusGroupCurrent.ToString();
                            NebCapture();
                            return;//may not be needed???
                        }
                        else
                            FilterFocusGroupCurrent = 0;//must be equal to nuber of groups so reset and move to next filter
                        //should continue to cycle here until FFGC + FG[1]
                        if (checkBox18.Checked == true)
                        {
                            CaptureTime3 = FlatExp;
                            CaptureBin = (int)numericUpDown27.Value;
                            DoFlat();
                        }
                        if (checkBox5.Checked == true)
                        {

                            currentfilter = 5;
                            DarksOn = true;
                            filteradvance();//back to pos 1 
                            Thread.Sleep(1000);
                            editfound = 0;
                            Nebname = "Dark1";
                            subsperfilter = (int)numericUpDown20.Value;
                            CaptureTime = (int)numericUpDown19.Value * 1000;
                            CaptureBin = (int)numericUpDown28.Value;
                            FilterFocusOn = false;
                            //  fileSystemWatcher4.EnableRaisingEvents = true;
                            NebCapture();
                            return;

                            /* old one....
                              DarksOn = true;
                            filteradvance();//back to pos 1
                            //    textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                            //    textBox21.Text = Filtertext;
                         //   NebListenStop();
                        //    Thread.Sleep(500);
                        //    SendKeys.SendWait("~");
                       //     SendKeys.Flush();
                            Thread.Sleep(1000);
                            editfound = 0;
                        //    clientSocket = new TcpClient();
                            Nebname = "Dark";
                            CaptureTime = (int)numericUpDown19.Value * 1000;
                            CaptureBin = (int)numericUpDown28.Value;
                            FilterFocusOn = false;
                         //   SetForegroundWindow(Aborthwnd);
                        //    PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                            Thread.Sleep(1000);
                            //  fileSystemWatcher4.EnableRaisingEvents = true;
                            NebCapture();
                            return;
                             * */
                        }

                        if (checkBox9.Checked == true)//Dark 2 allows for different time or bin
                        {
                            currentfilter = 6;
                            filteradvance();//back to pos 1 
                            Thread.Sleep(1000);
                            DarksOn = true;
                            editfound = 0;
                            Nebname = "Dark2";
                            subsperfilter = (int)numericUpDown29.Value;
                            CaptureTime = (int)numericUpDown30.Value * 1000;
                            CaptureBin = (int)numericUpDown31.Value;
                            FilterFocusOn = false;
                            //  fileSystemWatcher4.EnableRaisingEvents = true;
                            NebCapture();
                            return;
                        }
                        if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                        {
                            currentfilter = 7;
                            if (FlatsOn == false)
                                ToggleFlat();
                            filteradvance();//back to pos 1 
                            Thread.Sleep(1000);
                            editfound = 0;
                            Nebname = "Flat1";
                            subsperfilter = (int)numericUpDown32.Value;
                            CaptureTime = (int)numericUpDown34.Value;
                            CaptureBin = (int)numericUpDown36.Value;
                            FilterFocusOn = false;
                            //  fileSystemWatcher4.EnableRaisingEvents = true;

                            if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                            {
                                CalculateFlatExp();
                                return;
                            }

                            else
                                NebCapture();
                            return;
                        }

                    
                }
                if (currentfilter == 5)
                {
                    if (checkBox9.Checked == true)//Dark 2 allows for different time or bin
                    {
                        currentfilter = 6;
                        DarksOn = true;
                        //no advance already at 1
                        Thread.Sleep(1000);
                        editfound = 0;
                        Nebname = "Dark2";
                        subsperfilter = (int)numericUpDown29.Value;
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();
                        return;
                    }
                     if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                    {
                        currentfilter = 7;
                        if (FlatsOn == false)
                            ToggleFlat();
                        //  filteradvance();//back to pos 1 
                        Thread.Sleep(1000);
                        editfound = 0;
                        Nebname = "Flat1";
                        subsperfilter = (int)numericUpDown32.Value;
                        CaptureTime = (int)numericUpDown34.Value;
                        CaptureBin = (int)numericUpDown36.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                        {
                            CalculateFlatExp();
                            return;
                        }

                        else
                            NebCapture();
                        return;
                    }
                   
                
                }
                if (currentfilter == 6)
                {
                    if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                    {
                        currentfilter = 7;
                        if (FlatsOn == false)
                            ToggleFlat();
                        //  filteradvance();//back to pos 1 
                        Thread.Sleep(1000);
                        editfound = 0;
                        Nebname = "Flat1";
                        subsperfilter = (int)numericUpDown32.Value;
                        CaptureTime = (int)numericUpDown34.Value;
                        CaptureBin = (int)numericUpDown36.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                        {
                            CalculateFlatExp();
                            return;
                        }

                        else
                            NebCapture();
                        return;
                    }
                    
                }
               
             //   textBox38.Text = currentfilter.ToString();
         //   }
            /*
            catch(Exception ex)
            {
                Log("FilterSequence Error" + ex.ToString());
                Send("FilterSequence Error" + ex.ToString());
                FileLog("FilterSequence Error" + ex.ToString());

            }
             */
        }



        /*
        //ensure begin at filter 1 then goes to first filter
        private void checkBox22_Click(object sender, EventArgs e)
        {
            if (checkBox22.Checked == true)
            {
                totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value;
               // subsperfilter = totalsubs / filternumber;
               // textBox19.Text = subsperfilter.ToString();
                if (filternumber == 0)
                {
                    DialogResult result2 = MessageBox.Show("Must Select filters prior to enable", "scopefocus - Filter Control", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (result2 == DialogResult.OK)
                    {
                        checkBox22.Checked = false;
                        return;
                    }
                }
                //deterime number of fliters
                DialogResult result1 = MessageBox.Show("Ensure filter wheel at position 1", "scopefocus - Filter Control", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result1 == DialogResult.Cancel)
                {
                    return;
                }
                filterCountCurrent = 0;

                subCountCurrent = 0;


                if (checkBox1.Checked == true)
                {
                    currentfilter = 1;
                    subsperfilter = (int)numericUpDown12.Value;
                }
                else//go to starting filter position
                {
                    if (checkBox2.Checked == true)
                    {
                        filteradvance();
                        currentfilter = 2;
                        subsperfilter = (int)numericUpDown12.Value;

                    }
                    else
                    {
                        if (checkBox3.Checked == true)
                        {
                            filteradvance();
                          //Thread.Sleep(4000);
                            filteradvance();

                            currentfilter = 3;
                            subsperfilter = (int)numericUpDown14.Value;
                        }
                        else
                        {
                            if (checkBox4.Checked == true)
                            {
                                filteradvance();
                                filteradvance();
                                filteradvance();

                                currentfilter = 4;
                                subsperfilter = (int)numericUpDown15.Value;
                            }
                        }
                    }
                }

            //    fileSystemWatcher4.EnableRaisingEvents = true;

            }
        }
        */
        private double subs;
        private bool Capturing = false;
        private void NebCapture()//nebcapturehere (for searching)
        {
            try
            {
               if (IsServer())
                   SendtoSlave((subsperfilter * CaptureTime/1000).ToString());
                if (IsSlave())
                {//set # of subs based on server #subs * capture time.
                    CaptureTime = Convert.ToInt16(textBox37.Text)*1000;
                    subs = Convert.ToDouble(textBox41.Text)*1000 / CaptureTime;
                    subsperfilter = (int)Math.Round(subs, 0);
                    textBox42.Text = subsperfilter.ToString();
                }
             //   if ((SlavePaused == true) & (SlaveFlatOn == false))
                if (SlaveFlatOn == false)
                {
                     SlaveCapture();
                }
                Capturing = true;
                if (IsSlave())
                    Thread.Sleep(4000);//was 6000
                /*  //moved to DoFlats()
                if ((Nebname == "Flat1") || (Nebname == "Flat2"))//********added 5-1 stop phd for flats w/ OAG
                {
                    StopPHD();
                    Log("Guiding Stopped for Flats");
                   
                }
                 */
              //  if (FlatCalcDone == false)
            //    {
                /*
                    if ((IsSlave()) & clientSocket2.Connected == false)
                    {
                        NebListenStart(NebhWnd, SocketPort);
                        Log("Socket 2 connected");
                    }
                    else
                    {
                        if ((!IsSlave()) & (clientSocket.Connected == false))
                        {
                            NebListenStart(NebhWnd, SocketPort);
                            Log("Socket 1 connected");
                        }
                    }
                */
              //  if (NebListenOn == false)
                   NebListenStart(NebhWnd, SocketPort);
                    
             //   }
                    Thread.Sleep(1000);//was 3000 5-29
                    if (FlatsOn == true)
                    {
                        toolStripStatusLabel1.Text = "Saving Flats";
                        this.Refresh();
                    }
                    else
                    {
                        toolStripStatusLabel1.Text = "Capturing";
                      //  this.Refresh();
                    }
                    string prefix = textBox19.Text.ToString();
                 //   NetworkStream serverstream;
                if (IsSlave())
                     serverStream = clientSocket2.GetStream();
                else
                    serverStream = clientSocket.GetStream();
               
                    if (FlatCalcDone == false)
                    {
                        if (NebVNumber == 3)
                            CaptureTime3 = (double)CaptureTime / 1000;//need to allow fractions for flats.
                        //      textBox38.Text = CaptureTime3.ToString();//*******just for debugging                       
                    }
                    if (FlatCalcDone == true)
                        FlatCalcDone = false;
                if (DarksOn == false)
                {
                    
                    byte[] outStream = System.Text.Encoding.ASCII.GetBytes("setname " + prefix + Nebname + "\n" + "setbinning " + CaptureBin + "\n" + "SetShutter 0" + "\n" + "SetDuration " + CaptureTime3 + "\n" + "Capture " + subsperfilter + "\n");
                    try
                    {
                        serverStream.Write(outStream, 0, outStream.Length);
                        Log(prefix + Nebname + "  Captime " +  CaptureTime3.ToString() + "    subs " + subsperfilter.ToString());
                    }
                    catch
                    {
                        MessageBox.Show("Error sending command", "scopefocus");
                        return;

                    }
                }
                if (DarksOn == true)
                {
                    byte[] outStream2 = System.Text.Encoding.ASCII.GetBytes("setname " + prefix + Nebname + "\n" + "setbinning " + CaptureBin + "\n" + "SetShutter 1" + "\n" + "SetDuration " + CaptureTime3 + "\n" + "Capture " + subsperfilter + "\n");
                    try
                    {
                        serverStream.Write(outStream2, 0, outStream2.Length);
                        Log("Darks On  Cap time " + CaptureTime.ToString());
                    }
                    catch 
                    {
                        MessageBox.Show("Error sending command", "scopefocus");
                        return;

                    }
                    // serverStream.Write(outStream2, 0, outStream2.Length);
                    DarksOn = false;
                }
               // if (IsSlave())
                //    SendtoServer("Slave Capturing");//tell server capturing and wait for "sub complete"
               //*****6-6 this doesn't work (2 line below) 
              // if ((IsSlave()) || (IsServer()))
               //    WaitForSequenceDone("Sequence done", NebhWnd);
                
                //don't want this for slave cannot hit abort if using
                 // ***********remd 6-6
                if (FilterFocusOn == false)
                {
                    /*//  *****  this was remd 6-21  seems to work ok w/out.  minimal testing done
                    if (IsSlave())
                    {
                        Log("Check for Server sequence done");
                        while (working == true)
                        {
                            WaitForSequenceDone("Sequence done", NebhWnd);
                        }
                        working = true;
                        Log("Server sequence done");
                    }

                    else
                    {
                     */
                        while (Capturing == true)
                        {


                            int StatusstripHandle = FindWindowEx(NebhWnd, 0, "msctls_statusbar32", null);

                            //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                            IntPtr statusHandle = new IntPtr(StatusstripHandle);
                            StatusHelper sh = new StatusHelper(statusHandle);
                            string[] captions = sh.Captions;

                            if (captions[0] == "Sequence done")
                                Capturing = false;
                        }
                   // }
                }
                

                serverStream.Flush();
                
                //   toolStripStatusLabel1.Text = " ";
                
            }

                
            catch( Exception e)
            {
                Log("NebCapture Error" + e.ToString());
                Send("NebCapture Error" + e.ToString());
                FileLog("NebCapture Error" + e.ToString());

            }
                 
            
        }



        private void Listen0()
        {
            
            if ((IsSlave()) & clientSocket2.Connected == false)
            {
                NebListenStart(NebhWnd, SocketPort);
                Log("Socket 2 connected");
            }
           
                serverStream = clientSocket2.GetStream();
            
            byte[] outStream = System.Text.Encoding.ASCII.GetBytes("Listenport 0" + "\n");
            try
            {
                serverStream.Write(outStream, 0, outStream.Length);
                Log("Listenport 0 sent");
            }
            catch
            {
                MessageBox.Show("Error sending command", "scopefocus");
                return;

            }
            serverStream.Flush();
           // serverStream.Close();
           // clientSocket2.Client.Dispose();
          //  clientSocket2.Client.Close();
           // clientSocket2.Close();
        }
       
//******************************save this**************************

   /*     
  //try send command via clipboard WORKS
  private void ClipNebCapture()
  {
     // string String2 = "/NEB Delay 3000";
   //  Clipboard.SetDataObject(String2, true);
   //   Thread.Sleep(3000);
    //  string String = "/NEB Setduration 1000" + "\n" + "/NEBCapture " + subsperfilter.ToString();
        string String = "/NEB Setduration 1000" + "\n" + "/NEBCapture 5";
      Clipboard.SetDataObject(String, true);
      Thread.Sleep(3000);
     Clipboard.Clear();
    
  }
  */
        //*********************************************************************



        int StatusstripHandle;

        private void WaitForSequenceDone(string Statusstring, int handle)
        {
          
           //Thread.Sleep(500);
           // StatusLabel = Statusstring;
          //  Workerhwnd = handle;

          //  Log("Waiting for Slave");
         //   while (working)
         //   {
           
                 StatusstripHandle = FindWindowEx(handle, 0, "msctls_statusbar32", null);

                //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                IntPtr statusHandle = new IntPtr(StatusstripHandle);
                StatusHelper sh = new StatusHelper(statusHandle);
                captions = sh.Captions;
            //    Log("Caption[0]= " + captions[0].ToString());
                if (captions[0] == Statusstring)
                {
                    working = false;
                }

                this.Refresh();
                /* *****remd 6-20

                          
                             int StatusstripHandle = FindWindowEx(NebhWnd, 0, "msctls_statusbar32", null);

                                   //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                                   IntPtr statusHandle = new IntPtr(StatusstripHandle);
                                   StatusHelper sh = new StatusHelper(statusHandle);
                                   string[] captions = sh.Captions;
                                   while (captions[0] != "Sequence done")
                                   {
                                       Thread.Sleep(100);
                                   }
                            */
          //  }
          //  Log("Slave process complete");
             
        }


        private void ToolStripProgressBar()
        {
            toolStripProgressBar1.Maximum = totalsubs;
            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Increment(1);
            
            toolStripProgressBar1.Value = subCountCurrent;
            toolStripStatusLabel2.Text = "Sub " + subCountCurrent.ToString() + " of " + totalsubs.ToString();
            
        }
        private void EnableAllUpDwn()
        {
                numericUpDown12.Enabled = true;//no more chages allowed
                numericUpDown11.Enabled = true;
                numericUpDown24.Enabled = true;
                numericUpDown13.Enabled = true;//no more chages allowed
                numericUpDown16.Enabled = true;
                numericUpDown25.Enabled = true;
                numericUpDown14.Enabled = true;//no more chages allowed
                numericUpDown17.Enabled = true;
                numericUpDown26.Enabled = true;
                numericUpDown15.Enabled = true;//no more chages allowed
                numericUpDown18.Enabled = true;
                numericUpDown27.Enabled = true;
                numericUpDown20.Enabled = true;//no more chages allowed
                numericUpDown19.Enabled = true;
                numericUpDown28.Enabled = true;
                numericUpDown29.Enabled = true;//no more chages allowed
                numericUpDown30.Enabled = true;
                numericUpDown31.Enabled = true;
                numericUpDown32.Enabled = true;//no more chages allowed
                numericUpDown34.Enabled = true;
                numericUpDown36.Enabled = true;
            }

        private void DisableUpDwn()
        {
            if (currentfilter == 1)
            {
                numericUpDown12.Enabled = false;//no more chages allowed
                numericUpDown11.Enabled = false;
                numericUpDown24.Enabled = false;
            }
            if (currentfilter == 2)
            {
                numericUpDown13.Enabled = false;//no more chages allowed
                numericUpDown16.Enabled = false;
                numericUpDown25.Enabled = false;
            }
            if (currentfilter == 3)
            {

                numericUpDown14.Enabled = false;//no more chages allowed
                numericUpDown17.Enabled = false;
                numericUpDown26.Enabled = false;
            }
            if (currentfilter == 4)
            {
                numericUpDown15.Enabled = false;//no more chages allowed
                numericUpDown18.Enabled = false;
                numericUpDown27.Enabled = false;
            }
            if (currentfilter == 5)
            {
                numericUpDown20.Enabled = false;//no more chages allowed
                numericUpDown19.Enabled = false;
                numericUpDown28.Enabled = false;
            }
            if (currentfilter == 6)
            {
                numericUpDown29.Enabled = false;//no more chages allowed
                numericUpDown30.Enabled = false;
                numericUpDown31.Enabled = false;
            }
            if (currentfilter == 7)
            {
                numericUpDown32.Enabled = false;//no more chages allowed
                numericUpDown34.Enabled = false;
                numericUpDown36.Enabled = false;
            }
        }


        public bool SequenceRunning = false;
        //Go button, goes to starting filter pos, (assumes started at pos 1) and begins script
        private void SequenceGo()
        {
            try
            {
                
                    if ((IsServer()) & (StatusHandleRec == false))
                    {
                        NebSlavehwnd = Convert.ToInt32(textBox41.Text);
                        Log("NebSlave handle " + NebSlavehwnd.ToString());
                        StatusHandleRec = true;
                    }
                /*//*************rem'd 6-9
                //  SetForegroundWindow(NebhWnd);
                if ((IsServer()) & (Slavehwnd != 0))
                {

                    SlaveCapture();

                }
                 */
                watchforOpenPort();
                //   SetForegroundWindow(NebhWnd);
                if (NebCamera == "No camera")
                    NoCameraSelected();
                if (posMin == 0)
                {
                    DialogResult result = MessageBox.Show("Establish focus position first!", "scopefocus", MessageBoxButtons.OK);
                    if (result == DialogResult.OK)
                        return;
                }


                if ((checkBox10.Checked == false) & ((checkBox8.Checked == true) || (checkBox17.Checked == true)) & (button35.Text != "At Target"))
                {
                    MessageBox.Show("Slew to Target Position and retry", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                checkBox6.Checked = false;
                checkBox7.Checked = false;
                if (checkBox1.Checked == false)
                    numericUpDown12.Value = 0;
                if (checkBox2.Checked == false)
                    numericUpDown13.Value = 0;
                if (checkBox3.Checked == false)
                    numericUpDown14.Value = 0;
                if (checkBox4.Checked == false)
                    numericUpDown15.Value = 0;
                if (checkBox5.Checked == false)
                    numericUpDown20.Value = 0;
                if (checkBox9.Checked == false)
                    numericUpDown29.Value = 0;


                if (filtersynced == false)
                {
                    DialogResult result3 = MessageBox.Show("Must snyc filter position first", "scopefocus - Filter Control", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (result3 == DialogResult.OK)
                    {

                        return;
                    }
                }
                if (filtersynced == true)
                {
                    button26.BackColor = System.Drawing.Color.Lime;
                }
                TotalSubs();
                //totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + (int)numericUpDown32.Value;
                ToolStripProgressBar();

                // subsperfilter = totalsubs / filternumber;
                // textBox19.Text = subsperfilter.ToString();
                if (filternumber == 0)
                {
                    DialogResult result2 = MessageBox.Show("Must Select filters prior to enable", "scopefocus - Filter Control", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (result2 == DialogResult.OK)
                    {
                        //checkBox22.Checked = false;
                        return;
                    }
                }


                /*
                    DialogResult result1 = MessageBox.Show("Ensure filter wheel at position 1", "scopefocus - Filter Control", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (result1 == DialogResult.Cancel)
                    {
                        return;
                    }
                 */
                // NebListenStart();
                Thread.Sleep(1000);
                filterCountCurrent = 0;

                subCountCurrent = 0;
                button26.BackColor = System.Drawing.Color.Lime;

                if (checkBox1.Checked == true)
                {
                    currentfilter = 1;
                    textBox21.BackColor = System.Drawing.Color.White;
                    Filtertext = comboBox2.Text;
                    textBox21.Text = Filtertext;
                    if (checkBox17.Checked == false)
                    {
                        subsperfilter = (int)numericUpDown12.Value;
                        Nebname = comboBox2.Text;
                    }
                    else
                    {
                        FilterFocusGroupCurrent = 1;
                        subsperfilter = SubsPerFocus[0];
                        Nebname = comboBox2.Text + "_1";
                    }

                    CaptureTime = (int)numericUpDown11.Value * 1000;
                    CaptureBin = (int)numericUpDown24.Value;

                }
                else//go to starting filter position
                {
                    if (checkBox2.Checked == true)
                    {
                        filteradvance();
                        currentfilter = 2;
                        if (checkBox17.Checked == false)
                            subsperfilter = (int)numericUpDown13.Value;
                        else
                            subsperfilter = SubsPerFocus[1];
                        subsperfilter = (int)numericUpDown13.Value;
                        Nebname = comboBox3.Text.ToString();
                        CaptureTime = (int)numericUpDown16.Value * 1000;
                        CaptureBin = (int)numericUpDown25.Value;

                    }
                    else
                    {
                        if (checkBox3.Checked == true)
                        {
                            filteradvance();
                            //Thread.Sleep(4000);
                            filteradvance();

                            currentfilter = 3;
                            if (checkBox17.Checked == false)
                                subsperfilter = (int)numericUpDown14.Value;
                            else

                                subsperfilter = SubsPerFocus[2];

                            Nebname = comboBox4.Text.ToString();
                            CaptureTime = (int)numericUpDown17.Value * 1000;
                            CaptureBin = (int)numericUpDown26.Value;



                        }
                        else
                        {
                            if (checkBox4.Checked == true)
                            {
                                filteradvance();
                                filteradvance();
                                filteradvance();

                                currentfilter = 4;
                                if (checkBox17.Checked == false)
                                    subsperfilter = (int)numericUpDown15.Value;
                                else
                                    subsperfilter = SubsPerFocus[3];
                                Nebname = comboBox5.Text.ToString();
                                CaptureTime = (int)numericUpDown18.Value * 1000;
                                CaptureBin = (int)numericUpDown27.Value;
                            }
                            else
                            {
                                if (checkBox5.Checked == true)
                                {
                                    DarksOn = true;
                                    // currentfilter = 5;
                                    subsperfilter = (int)numericUpDown20.Value;
                                    Nebname = "Dark";
                                    CaptureTime = (int)numericUpDown19.Value * 1000;
                                    CaptureBin = (int)numericUpDown28.Value;
                                }

                                else
                                {
                                    if (checkBox9.Checked == true)
                                    {
                                        MessageBox.Show("Use Dark1 for single dark frame", "scopefocus");
                                        return;
                                    }

                                    else
                                    {
                                        if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                                        {
                                            FlatsOn = true;
                                            // currentfilter = 5;
                                            subsperfilter = (int)numericUpDown32.Value;
                                            Nebname = "Flat1";
                                            CaptureTime = (int)numericUpDown34.Value;
                                            CaptureBin = (int)numericUpDown36.Value;
                                        }

                                    }
                                }
                            }
                        }
                    }
                    // Thread.Sleep(1000);
                    /*  added the remd section below to neblistenstart


                     if (firstconnect == true)//if already connected once, this closes and re-connects on repeat "Go" push
                    {
                        clientSocket.Close();

                        clientSocket = new TcpClient();
                        clientSocket.Connect("127.0.0.1", 4301);//connects to neb
                    }

                    if (firstconnect == false)// establish first time connection
                    {
                
                        clientSocket.Connect("127.0.0.1", 4301);//connects to neb
                        firstconnect = true;
                    }
             
                   
                    //Thread.Sleep(500);
              
                     */
                }
                // if (checkBox8.Checked == true) //*******rem'd 4_5
                SequenceRunning = true;
                DisableUpDwn();
                fileSystemWatcher4.EnableRaisingEvents = true;
                checkfiltercount();
            }
            catch(Exception ex)
            {
                Log("SequenceGo Error" + ex.ToString());
                Send("SequenceGo Error" + ex.ToString());
                FileLog("SequenceGo Error" + ex.ToString());

            }

            
        }
        private void button26_Click(object sender, EventArgs e)
        {
            if (paused == true)
            {
                button22.PerformClick();
                return;
            }
            SequenceGo();
            
        }
        private void Refocus()
        {
            if (checkBox10.Checked == true)
            {
                NebFineFocus();
            }

        }


        private void numericUpDown12_ValueChanged(object sender, EventArgs e)
        {
            if ((currentfilter == 1) && (subCountCurrent > 0))
                numericUpDown12.Enabled = false;
            if (subCountCurrent == 0)
            {
                if (checkBox6.Checked == true)
                {
                    if (checkBox2.Checked == true)
                        numericUpDown13.Value = numericUpDown12.Value;
                    if (checkBox3.Checked == true)
                        numericUpDown14.Value = numericUpDown12.Value;
                    if (checkBox4.Checked == true)
                        numericUpDown15.Value = numericUpDown12.Value;
                    if (checkBox5.Checked == true)
                        numericUpDown20.Value = numericUpDown12.Value;

                }
            }
            TotalSubs();
               // totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + (int)numericUpDown32.Value;
                textBox23.Text = totalsubs.ToString();
            
           
              
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown11_ValueChanged(object sender, EventArgs e) //exp time filter 1
        {
            if (subCountCurrent == 0)
            {
                if (checkBox7.Checked == true)
                {
                    if (checkBox2.Checked == true)
                        numericUpDown16.Value = numericUpDown11.Value;
                    if (checkBox3.Checked == true)
                        numericUpDown17.Value = numericUpDown11.Value;
                    if (checkBox4.Checked == true)
                        numericUpDown18.Value = numericUpDown11.Value;
                    if (checkBox5.Checked == true)
                        numericUpDown19.Value = numericUpDown11.Value;
                }
            }
            TotalSubs();
          //  totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + (int)numericUpDown32.Value;
            textBox23.Text = totalsubs.ToString();  
            //****5-11 trying to be able to change sub number   
        }

        private void playsound()
        {
            SoundPlayer simpleSound = new SoundPlayer(@"c:\Windows\Media\notify.wav");
            simpleSound.Play();
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            CountFilterTotal();
            if (checkBox5.Checked == false)
            {
                numericUpDown19.Value = 0;
                numericUpDown20.Value = 0;
            }
        }
        //filter pos sync to 1
        private void button28_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (checkBox16.Checked == false)
               
                {
                    toolStripStatusLabel1.Text = "Filter Moving";
                    this.Refresh();
                    string movedone;
                    port.DiscardOutBuffer();
                    port.DiscardInBuffer();
                    Thread.Sleep(20);
                    port.Write("Z");
                    filterMoving = true;
                    Thread.Sleep(50);

                    while (filterMoving == true)
                    {
                        movedone = port.ReadExisting();
                        if (movedone == "D")
                        {
                            filterMoving = false;
                            this.Refresh();
                            toolStripStatusLabel1.Text = "Filter Sync Complete";
                            //   toolStripStatusLabel1.Text = " ";
                            //    toolStripStatusLabel1.Text = " ";

                        }
                    }

                    Thread.Sleep(50);
                    port.DiscardInBuffer();
                }
                Filtertext = comboBox2.Text;
                currentfilter = 1;
                DisplayCurrentFilter();                
                filtersynced = true;
                button28.BackColor = System.Drawing.Color.Lime;

            }

            catch(Exception ex)
            {
                Log("Filter Position Sync Error" + ex.ToString());
            }

        }

        private void numericUpDown16_ValueChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                if (checkBox1.Checked == true)
                    numericUpDown11.Value = numericUpDown16.Value;
                if (checkBox3.Checked == true)
                    numericUpDown17.Value = numericUpDown16.Value;
                if (checkBox4.Checked == true)
                    numericUpDown18.Value = numericUpDown16.Value;
                if (checkBox5.Checked == true)
                    numericUpDown19.Value = numericUpDown16.Value;
            }
            TotalSubs();
            // totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + (int)numericUpDown32.Value;
                
        }

        private void numericUpDown17_ValueChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                if (checkBox2.Checked == true)
                    numericUpDown16.Value = numericUpDown17.Value;
                if (checkBox1.Checked == true)
                    numericUpDown11.Value = numericUpDown17.Value;
                if (checkBox4.Checked == true)
                    numericUpDown18.Value = numericUpDown17.Value;
                if (checkBox5.Checked == true)
                    numericUpDown19.Value = numericUpDown17.Value;
            }
            TotalSubs();
           // totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + (int)numericUpDown32.Value;
                
        }

        private void numericUpDown18_ValueChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                if (checkBox2.Checked == true)
                    numericUpDown16.Value = numericUpDown18.Value;
                if (checkBox3.Checked == true)
                    numericUpDown17.Value = numericUpDown18.Value;
                if (checkBox1.Checked == true)
                    numericUpDown11.Value = numericUpDown18.Value;
                if (checkBox5.Checked == true)
                    numericUpDown19.Value = numericUpDown18.Value;
            }
            TotalSubs();

           // totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + (int)numericUpDown32.Value;
                
        }

        private void numericUpDown19_ValueChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                if (checkBox2.Checked == true)
                    numericUpDown16.Value = numericUpDown19.Value;
                if (checkBox3.Checked == true)
                    numericUpDown17.Value = numericUpDown19.Value;
                if (checkBox4.Checked == true)
                    numericUpDown18.Value = numericUpDown19.Value;
                if (checkBox1.Checked == true)
                    numericUpDown11.Value = numericUpDown19.Value;
            }
            TotalSubs();
        }

        private void numericUpDown13_ValueChanged(object sender, EventArgs e)
        {
            
            if (subCountCurrent == 0)
            {
                if (checkBox6.Checked == true)
                {
                    if (checkBox1.Checked == true)
                        numericUpDown12.Value = numericUpDown13.Value;
                    if (checkBox3.Checked == true)
                        numericUpDown14.Value = numericUpDown13.Value;
                    if (checkBox4.Checked == true)
                        numericUpDown15.Value = numericUpDown13.Value;
                    if (checkBox5.Checked == true)
                        numericUpDown20.Value = numericUpDown13.Value;
                }
            }
            TotalSubs();
        }

        private void numericUpDown14_ValueChanged(object sender, EventArgs e)
        {
            if (subCountCurrent == 0)
            {
                if (checkBox6.Checked == true)
                {
                    if (checkBox2.Checked == true)
                        numericUpDown13.Value = numericUpDown14.Value;
                    if (checkBox1.Checked == true)
                        numericUpDown12.Value = numericUpDown14.Value;
                    if (checkBox4.Checked == true)
                        numericUpDown15.Value = numericUpDown14.Value;
                    if (checkBox5.Checked == true)
                        numericUpDown20.Value = numericUpDown14.Value;

                }
            }
            TotalSubs();
        }

        private void numericUpDown15_ValueChanged(object sender, EventArgs e)
        {
            if (subCountCurrent == 0)
            {
                if (checkBox6.Checked == true)
                {
                    if (checkBox2.Checked == true)
                        numericUpDown13.Value = numericUpDown15.Value;
                    if (checkBox3.Checked == true)
                        numericUpDown14.Value = numericUpDown15.Value;
                    if (checkBox1.Checked == true)
                        numericUpDown12.Value = numericUpDown15.Value;
                    if (checkBox5.Checked == true)
                        numericUpDown20.Value = numericUpDown15.Value;
                }
            }
            TotalSubs(); 
        }

        private void numericUpDown20_ValueChanged(object sender, EventArgs e)
        {
            if (subCountCurrent == 0)
            {
                if (checkBox6.Checked == true)
                {
                    if (checkBox2.Checked == true)
                        numericUpDown13.Value = numericUpDown20.Value;
                    if (checkBox3.Checked == true)
                        numericUpDown14.Value = numericUpDown20.Value;
                    if (checkBox4.Checked == true)
                        numericUpDown15.Value = numericUpDown20.Value;
                    if (checkBox1.Checked == true)
                        numericUpDown12.Value = numericUpDown20.Value;
                }
            }
            TotalSubs();    
        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            this.numericUpDown6.Maximum = travel;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            
           // posMin = Convert.ToInt16(textBox4.Text);//allows direct entry of close focus position
        }
       // string[] metricpath;
        private void fileSystemWatcher5_Changed(object sender, FileSystemEventArgs e)
        {
          //  try
          //  {


                //  string[] metricpath = Directory.GetFiles(path2.ToString(), "*.fit");

                //   button25.PerformClick();
                if (MetricSample == false)
                {
                    vcurve();
                   // MetricCapture();                 
                }
                if (MetricSample == true)
                {
                    metricpath = Directory.GetFiles(path2.ToString(), "metric*.fit");
                    metricHFR = GetMetric(metricpath, roundto);
                    textBox25.Text = metricHFR.ToString();

                    AvgMetricHFR[currentmetricN - 1] = metricHFR;

                    if (currentmetricN == metricN)
                    {
                        MetricSample = false;
                        fileSystemWatcher5.EnableRaisingEvents = false;
                        AvgMetric = (AvgMetricHFR.Sum()) / metricN;
                        textBox25.Text = AvgMetric.ToString();
                       // NetworkStream serverStream = clientSocket.GetStream();
                        /*
                        serverStream = clientSocket.GetStream();
                        byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                        serverStream.Write(outStream, 0, outStream.Length);
                        serverStream.Flush();
                        
                        File.Delete(metricpath[0]);
                        Thread.Sleep(3000);
                        serverStream.Close();
                        Thread.Sleep(3000);
                        SetForegroundWindow(NebhWnd);
                        Thread.Sleep(1000);
                        PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                        Thread.Sleep(1000);
                        NebListenOn = false;
                        // clientSocket.GetStream().Close();//added 5-17-12
                        //  clientSocket.Client.Disconnect(true);//added 5-17-12
                        clientSocket.Close();
                        */
                        if (metricpath != null)
                            File.Delete(metricpath[0]);
                        currentmetricN = 0;
                        return;
                    }

                }
                if (vProgress != vN+1)
                {
                    currentmetricN++;
                    MetricCapture();
                }
            }
        /*
            catch(Exception ex)
            {
                Log("Metric Error Line 5446" + ex.ToString());
                Send("Metric Error Line 5446" + ex.ToString());
                FileLog("Metric Error Line 5446" + ex.ToString());

            }
         */
      //  }
        //metric Go button,  start N captures Fine V
        private void button25_Click(object sender, EventArgs e)
        {
            if (checkBox22.Checked == false)
                checkBox22.Checked = true;
            /*
            if (roughvdone == false)//make sure rough v curve done to establish rough focus point
            {
                DialogResult result2;
                result2 = MessageBox.Show("Must perform a rough V-curve first", "scopefocus",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (result2 == DialogResult.OK)
                {
                    return;
                }
                return;
            }
             */
            button3.PerformClick();
            /*
            if (clientSocket.Connected == false)
            {
                //clientSocket.Connect("127.0.0.1", SocketPort);//connects to neb
                NebListenStart(NebhWnd, SocketPort);
            }
             */
            fileSystemWatcher5.EnableRaisingEvents = true;
            MetricCapture();

        }
        //metriccapturehere
        private void MetricCapture()
        {

            try
            {
                MetricTime = Convert.ToDouble(textBox43.Text);
                if (clientSocket.Connected == false)
                    NebListenStart(NebhWnd, SocketPort);
              

              //  NetworkStream serverStream = clientSocket.GetStream();
                serverStream = clientSocket.GetStream();
                byte[] outStream = System.Text.Encoding.ASCII.GetBytes("SetDuration " + MetricTime + "\n" + "CaptureSingle metric" + "\n");
                serverStream.Write(outStream, 0, outStream.Length);
                serverStream.Flush();
            }
            catch(Exception ex)
            {
                Log("MetricCapture Error line 4239" + ex.ToString());
                Send("MetricCapture Error line 4239" + ex.ToString());
                FileLog("MetricCapture Error line 4239" + ex.ToString());

            }


        }

        private void checkBox22_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox22.Checked == true)
                Filtertext = Filtertext + "Metric";
            if (checkBox22.Checked == false)
                Filtertext = Filtertext.Replace("Metric", "");          
        }

        //sample specified number of metricHFR's---(future)this could be used as a qaulity monitor and prompt 
        //focus based on some parameters
        double MetricTime;
        public void SampleMetric()
        {
            try
            {
                metricN = (int)numericUpDown5.Value;
                AvgMetricHFR = new int[metricN];
                MetricTime = Convert.ToDouble(textBox43.Text);
                currentmetricN = 1;
                MetricSample = true;
                if (clientSocket.Connected == false)
                {
                    NebListenStart(NebhWnd, SocketPort);
                   // clientSocket.Connect("127.0.0.1", SocketPort);//connects to neb
                }
                fileSystemWatcher5.EnableRaisingEvents = true;
               // NetworkStream serverStream = clientSocket.GetStream();
                serverStream = clientSocket.GetStream();
                byte[] outStream = System.Text.Encoding.ASCII.GetBytes("SetDuration " + MetricTime + "\n" + "CaptureSingle metric" + "\n");
                serverStream.Write(outStream, 0, outStream.Length);

                serverStream.Flush();
            }
            catch(Exception ex)
            {
                Log("SampleMetric Error Line 5642" + ex.ToString());
            }

        }

      
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
       
        
        private void button36_Click(object sender, EventArgs e)
        {
            try
            {
                if (port2selected == null)
                {
                    MessageBox.Show("Port not selected", "scopefocus");
                    return;
                }
                if (port2 == null)
                {
                    Port2Open();
                    Thread.Sleep(200);
                }
                if (connect2 == 1)
                {
                    DialogResult result2 = MessageBox.Show("Already Connected", "Arduino scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                if (port2.IsOpen == true)
                {
                    Thread.Sleep(500);
                    CheckNexremoteConnection();
                }
            }
            catch(Exception ex)
            {
                Log("Port 2 Error  Line 4330" + ex.ToString());
            }

            
        }


        private void CheckNexremoteConnection()
        {
            try
            {
                if (port2.IsOpen == false)
                    Port2Open();
                string CommCheck;
                port2.DiscardInBuffer();
                port2.DiscardOutBuffer();
                Thread.Sleep(50);
                port2.Write("Kx");//K is echo command x is an arbitrary char
                Thread.Sleep(100);
                port2.DiscardOutBuffer();

                Thread.Sleep(50);
                CommCheck = port2.ReadExisting();
                Thread.Sleep(50);
                if (CommCheck == "x#") //check for echo of x and stop bit
                {
                    connect2 = 1;
                    button36.Text = "Mount Connected";
                    button36.BackColor = System.Drawing.Color.Lime;
                    port2.DiscardOutBuffer();
                    port2.DiscardInBuffer();
                    Log("Connected to Mount on " + port2selected.ToString());

                }
                else
                    MessageBox.Show("Nexremote Connection Failed", "scopefocus");

                port2.Close();
            }
            catch(Exception ex)
            {
                Log("Check Nexremote Connection Error" + ex.ToString());
            }
       }


        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            port2selected = comboBox6.SelectedItem.ToString();
          
           
        }

       
        public static double ConvertDec(string command, out int deg, out int min, out int sec)
        {
            string HexDecFocus = command.Substring(9, 8);
            string DecDegHex = HexDecFocus.Substring(0, 2);
            string DecMinHex = HexDecFocus.Substring(2, 2);
            string DecSecHex = HexDecFocus.Substring(4, 2);
            int DecDegDec = int.Parse(DecDegHex, System.Globalization.NumberStyles.HexNumber);
            int DecMinDec = int.Parse(DecMinHex, System.Globalization.NumberStyles.HexNumber);
            int DecSecDec = int.Parse(DecSecHex, System.Globalization.NumberStyles.HexNumber);
            double DecTotalSec = (DecDegDec * 256 * 256 + DecMinDec * 256 + DecSecDec)/12/1.078781893;
            deg = ((int)DecTotalSec) / 60 / 60;
            min = ((int)DecTotalSec - (deg*3600))/60;
            sec = ((int)DecTotalSec - (deg * 3600) - (min * 60));
            return 1;


        }

        public static double ConvertRA(string command, out int hr, out int min, out double sec)
        {
            string HexRaFocus = command.Substring(0, 8);//parse RA portion
            string RaDegHex = HexRaFocus.Substring(0, 2);//high byte
            string RaMinHex = HexRaFocus.Substring(2, 2);//med byte
            string RaSecHex = HexRaFocus.Substring(4, 2);//low byte
            int RaDegDec = int.Parse(RaDegHex, System.Globalization.NumberStyles.HexNumber);//convert to decimal
            int RaMinDec = int.Parse(RaMinHex, System.Globalization.NumberStyles.HexNumber);
            int RaSecDec = int.Parse(RaSecHex, System.Globalization.NumberStyles.HexNumber);
            double RaTotalSec = (RaDegDec * 256 * 256 + RaMinDec * 256 + RaSecDec) / 12 / 15 / 1.078781893;
            hr = ((int)RaTotalSec) / 60 / 60;
            min = ((int)RaTotalSec - (hr * 3600)) / 60;
            sec = Math.Round(RaTotalSec - (hr * 3600) - (min * 60), 2);
            return 1;


        }
        public int RAhr;
        public int RAmin;
        public double RAsec;
        public int DecDeg;
        public int DecMin;
        public int DecSec;
        private void GetFocusLocation()
        {
            try
            {
                if (port2 == null)
                {
                    MessageBox.Show("Not Connected to Nexremote", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (port2.IsOpen == false)
                    Port2Open();

                port2.DiscardOutBuffer();
                port2.DiscardInBuffer();
                Thread.Sleep(20);
                port2.Write("e");
                Thread.Sleep(100);
                FocusLoc = port2.ReadExisting();
                ConvertDec(FocusLoc, out DecDeg, out DecMin, out DecSec);
                ConvertRA(FocusLoc, out RAhr, out RAmin, out RAsec);
                Log("Focus Position Set: RA " + RAhr.ToString() + "hr " + RAmin.ToString() + "m " + RAsec.ToString() + "s " + "- Dec " + DecDeg.ToString() + "° " + DecMin.ToString() + "' " + DecSec.ToString() + "\"");
                Thread.Sleep(100);
                port2.DiscardInBuffer();
                Thread.Sleep(10);
                if (FocusLoc.Substring(17, 1) == "#")//check for stop bit
                {
                    FocusLocObtained = true;
                    button32.BackColor = System.Drawing.Color.Lime;
                    button32.Text = "Focus Pos Set";
                }
                string path = textBox11.Text.ToString();
                string fullpath = path + @"\log.txt";
                StreamWriter log;
                string string4 = "Focus Position Set: RA " + RAhr.ToString() + "hr " + RAmin.ToString() + "m " + RAsec.ToString() + "s " + "- Dec " + DecDeg.ToString() + "° " + DecMin.ToString() + "' " + DecSec.ToString() + "\"";
                if (!File.Exists(fullpath))
                {
                    log = new StreamWriter(fullpath);
                }
                else
                {
                    log = File.AppendText(fullpath);
                }
                log.WriteLine(string4);
                log.Close();
                port2.Close();
            }
            catch(Exception ex)
            {
                Log("GetFocus Location Error" + ex.ToString());
            }
        }
        //get focus location
        private void button32_Click(object sender, EventArgs e)
        {
            button33.Text = "At Focus";
            button33.BackColor = System.Drawing.Color.Lime;
            button35.Text = "Goto Target";
            button35.UseVisualStyleBackColor = true;
            GetFocusLocation();
        }

        private void GetTargetLocation()
        {
            try
            {
                if (port2.IsOpen == false)
                    Port2Open();
                if (port2 == null)
                {
                    MessageBox.Show("Not Connected to Nexremote", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                port2.DiscardOutBuffer();
                port2.DiscardInBuffer();
                Thread.Sleep(20);
                port2.Write("e");
                Thread.Sleep(100);
                TargetLoc = port2.ReadExisting();
                //  Thread.Sleep(50);
                port2.DiscardInBuffer();
                //  textBox28.Clear();
                Thread.Sleep(100);
                ConvertDec(TargetLoc, out DecDeg, out DecMin, out DecSec);
                ConvertRA(TargetLoc, out RAhr, out RAmin, out RAsec);
                Log("Target Position Set: RA " + RAhr.ToString() + "hr " + RAmin.ToString() + "m " + RAsec.ToString() + "s " + "- Dec " + DecDeg.ToString() + "° " + DecMin.ToString() + "' " + DecSec.ToString() + "\"");

                //     textBox28.Text = TargetLoc.ToString();
                if (TargetLoc.Substring(17, 1) == "#")//check for stop bit
                {
                    TargetLocObtained = true;
                    button34.BackColor = System.Drawing.Color.Lime;
                    button34.Text = "Target Pos Set";
                }
                string path = textBox11.Text.ToString();
                string fullpath = path + @"\log.txt";
                StreamWriter log;
                string string4 = "Target Position Set: RA " + RAhr.ToString() + "hr " + RAmin.ToString() + "m " + RAsec.ToString() + "s " + "- Dec " + DecDeg.ToString() + "° " + DecMin.ToString() + "' " + DecSec.ToString() + "\"";
                if (!File.Exists(fullpath))
                {
                    log = new StreamWriter(fullpath);
                }
                else
                {
                    log = File.AppendText(fullpath);
                }
                log.WriteLine(string4);
                log.Close();
                port2.Close();
            }
            catch(Exception ex)
            {
                Log("GetTargetLocation Error" + ex.ToString());

            }
        }


 
        private bool TrackingOn = true;
        private int retry = 0;
        private void StopTracking()
        {
            try
            {
                if (TrackingOn == true)
                {
                    if (port2.IsOpen == false)
                        Port2Open();
                    port2.DiscardInBuffer();
                    port2.DiscardOutBuffer();
                    Thread.Sleep(50);
                    byte[] newMsg = { 0x54, 0x00 };
                    port2.Write(newMsg, 0, newMsg.Length);
                    Thread.Sleep(100);
                    port2.DiscardOutBuffer();

                    Thread.Sleep(50);
                    string Handshake = port2.ReadExisting();
                    Thread.Sleep(50);
                    if (Handshake == "#") //check for handshake
                    {
                        Log("Tracking Command Received");
                    }

                    port2.DiscardInBuffer();
                    port2.DiscardOutBuffer();
                    Thread.Sleep(50);
                    port2.Write("t");//check tracking mode
                    Thread.Sleep(100);
                    port2.DiscardOutBuffer();
                    // byte[] confirm = new byte[1];
                    int confirm;
                    confirm = port2.ReadByte();
                    Thread.Sleep(20);
                    if (confirm == 0)
                    {
                        Log("Tracking Off Confirmed");
                        TrackingOn = false;
                    }
                    else
                    {
                        if (retry < 3)
                        {
                            retry++;
                            Thread.Sleep(1000);
                            StopTracking();
                        }
                        else
                        {
                            retry = 0;
                            MessageBox.Show("Unable to stop tracking", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            //*****************eventually add text message here ******************************

                        }
                    }
                }
                else
                    MessageBox.Show("Tracking is Already Off", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                port2.Close();
            }
            catch(Exception ex)
            {
                Log("Stop Tracking Error" + ex.ToString());
                Send("Stop Tracking Error" + ex.ToString());
                FileLog("Stop Tracking Error" + ex.ToString());

            }
        }

        //Get target location 
        private void button34_Click(object sender, EventArgs e)
        {
            button35.Text = "At Target";
            button35.BackColor = System.Drawing.Color.Lime;
            button33.Text = "Goto";
            button33.UseVisualStyleBackColor = true;
            GetTargetLocation(); 
            
        }
        private bool FocusGotoOn = false;
        //goto focus
        private void GotoFocusLocation()
        {
            try
            {
                FocusGotoOn = true;
                if (FocusLocObtained == false)
                {
                    MessageBox.Show("Focus Position Not Set", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (port2.IsOpen == false)
                    Port2Open();
                string FocusCommand = FocusLoc.Substring(0, 17);
                port2.DiscardOutBuffer();
                port2.DiscardInBuffer();
                Thread.Sleep(20);
                port2.Write("r" + FocusCommand);
                MountMoving = true;
                Thread.Sleep(50);
                port2.DiscardOutBuffer();
                port2.DiscardInBuffer();
                Thread.Sleep(20);

              //  toolStripStatusLabel1.Text = "Slewing to Focus";  remd 3_13 reduddant w/ 5022

                CheckForSlewDone();
                /*
                while (MountMoving == true)
                {
                   Thread.Sleep(50); // pause for 1/20 second
                   System.Windows.Forms.Application.DoEvents();
                }
     */
                Thread.Sleep(100);


            }
            catch(Exception ex)
            {
                Log("GotoFocus Location Error" + ex.ToString());
                Send("GotoFocus Location Error" + ex.ToString());
                FileLog("GotoFocus Location Error" + ex.ToString());

            }

        }

        private void CheckForSlewDone()
        {
            try
            {
                if (backgroundWorker1.IsBusy != true)
                {
                    // Start the asynchronous operation.
                    backgroundWorker1.RunWorkerAsync();
                }


                /*
                if (port2.IsOpen == false)
                    port2.Open();
                while (MountMoving == true)
                {
                    port2.DiscardInBuffer();
                    port2.Write("L");
                    Thread.Sleep(20);
                    port2.DiscardOutBuffer();
                
                  //  textBox13.Clear();
                    Thread.Sleep(20);

                    //  
                    GotoDoneCommand = port2.ReadExisting();
                    Thread.Sleep(10);
                  //  port2.DiscardOutBuffer();
                    port2.DiscardInBuffer();
                   // textBox13.Text = GotoDoneCommand.ToString();
                    Thread.Sleep(20);
                    if (GotoDoneCommand == "0#")
                    {
                      //  textBox13.Text = "Goto Done";
                        MountMoving = false;
                        port2.DiscardInBuffer();
                        port2.DiscardOutBuffer();
                       // break;
                        return;
                    }
                }
               */
            }
            catch(Exception ex)
            {
                Log("CheckForSlewDone Error" + ex.ToString());
                Send("CheckForSlewDone Error" + ex.ToString());
                FileLog("CheckForSlewDone Error" + ex.ToString());
                
            }
        }
        //goto focus location button
        private void button33_Click(object sender, EventArgs e)
        {
            try
            {
                GotoFocusLocation();
                while (MountMoving == true)
                {
                    Thread.Sleep(50); // pause for 1/20 second
                    System.Windows.Forms.Application.DoEvents();
                }
                button33.Text = "At Focus";
                button33.BackColor = System.Drawing.Color.Lime;
                button35.Text = "Goto";
                button35.UseVisualStyleBackColor = true;
            }
            catch(Exception ex)
            {
                Log("GotoFocus locaton Button Error" + ex.ToString());
            }
            
        }

        private bool TargetGotoOn = false;
        private void GotoTargetLocation()
        {
            try
            {
                if (TargetLocObtained == false)
                {
                    MessageBox.Show("Target Position Not Set", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                TargetGotoOn = true;
                if (port2.IsOpen == false)
                    Port2Open();
                string TargetCommand = TargetLoc.Substring(0, 17);
                port2.DiscardOutBuffer();
                port2.DiscardInBuffer();
                Thread.Sleep(20);
                port2.Write("r" + TargetCommand);
                MountMoving = true;
                Thread.Sleep(50);
                port2.DiscardOutBuffer();
                port2.DiscardInBuffer();

                Thread.Sleep(20);
                //      toolStripStatusLabel1.Text = "Going to Target";  rem'd 3/9


                CheckForSlewDone();
                /*
                    while (MountMoving == true)
                    {
                        Thread.Sleep(50); // pause for 1/20 second
                        System.Windows.Forms.Application.DoEvents();
                    }
 
             */
            }
            catch(Exception ex)
            {
                Log("GotoTargetLocation Error" + ex.ToString());
                Send("GotoTargetLocation Error" + ex.ToString());
                FileLog("GotoTargetLocation Error" + ex.ToString());

            }
        }
        
        //goto target
        private void button35_Click(object sender, EventArgs e)
        {
            try
            {
                GotoTargetLocation();
                while (MountMoving == true)
                {
                    Thread.Sleep(50); // pause for 1/20 second
                    System.Windows.Forms.Application.DoEvents();
                }
                button35.Text = "At Target";
                button35.BackColor = System.Drawing.Color.Lime;
                button33.Text = "Goto";
                button33.UseVisualStyleBackColor = true;
            }
            catch(Exception ex)
            {
                Log("gotoarget location Button error" + ex.ToString());
            }
        }
        //Abort slew
        private void button37_Click(object sender, EventArgs e)
        {
            try
            {
                if (backgroundWorker1.WorkerSupportsCancellation == true)
                {
                    // Cancel the asynchronous operation.
                    backgroundWorker1.CancelAsync();
                }

                if (port2.IsOpen == false)
                    Port2Open();
                port2.DiscardOutBuffer();
                port2.DiscardInBuffer();
                Thread.Sleep(20);
                port2.Write("M");
                Thread.Sleep(50);
                port2.DiscardOutBuffer();
                port2.DiscardInBuffer();
                FocusGotoOn = false;
                TargetGotoOn = false;
                toolStripStatusLabel1.Text = "Slew Aborted";
                //   port2.Close();
            }
            catch(Exception ex )
            {
                Log("Abort Slew Button Error" + ex.ToString());

            }
        }

//try send focus time via script  *****Close****
        private void SendFocusTime(float dur)
        {
            try
            {
                if (IsSlave())
                    Thread.Sleep(3000);
                
                if (IsSlave())
                {
                    
                    if (clientSocket2.Connected == false)
                    {
                        Log("Socket 2 was not connected");
                        NebListenStart(NebhWnd, SocketPort);                     
                    }
                }
                else
                {
                    
                    if (clientSocket.Connected == false)
                    {
                        Log("Socket 1 was not connected");
                        NebListenStart(NebhWnd, SocketPort);
                    }
                }
                 
               // if (NebListenOn == false)
            //    NebListenStart(NebhWnd, SocketPort); //  this seems to be needed if about is hit too many time on slave
               // Thread.Sleep(1000);remd 6-3
                toolStripStatusLabel1.Text = "SendFocusTime";
                this.Refresh();

              //  int bin = (int)numericUpDown23.Value;
                
                if (IsSlave())
                    serverStream = clientSocket2.GetStream();
                else
                    serverStream = clientSocket.GetStream();

                
                byte[] outStream = System.Text.Encoding.ASCII.GetBytes("SetDuration " + dur + "\n");
               

             //   byte[] outStream = System.Text.Encoding.ASCII.GetBytes("setbinning " + bin + "\n" + "SetDuration " + dur + "\n");****remd 3_13 to try above
                try
                {
                    serverStream.Write(outStream, 0, outStream.Length);
                    Log("FocusTime sent " + dur.ToString());
                }
                catch 
                {
                    MessageBox.Show("Error sending command", "scopefocus");
                    return;

                }
                serverStream.Flush();
                Thread.Sleep(1000);
                //*********************goes farther without this NebListenStop but doesn't hit enter 
                // for client socket closed message boxes
                //*************need to eleminiate mult. socket opwn/close  *******************

                //   NebListenStop();
            }
            catch(Exception ex)
            {
                Log("SendFocutTime Error" + ex.ToString());
                Send("SendFocutTime Error" + ex.ToString());
                FileLog("SendFocutTime Error" + ex.ToString());

            }
           
                }

/*

//works without bin change option.  
        private void SendFocusTime(float dur)
        {
            string send = "";
         //   dur = (double)numericUpDown22.Value;
            
            send = dur.ToString();
            StringBuilder sb = new StringBuilder(send);
            SetForegroundWindow(hwndDuration);
            // sb = dur.ToString();
            SendMessage(hwndDuration, WM_SETTEXT, 0, sb);//set capture time
            
            //set focus Bin  *** for some reason only works once ****
            Thread.Sleep(500);
    *******rem from here down to do without bin change**************works
            SetForegroundWindow(Advancedhwnd);
            PostMessage(Advancedhwnd, BN_CLICKED, 0, 0);
            Thread.Sleep(1000);
            GetHandles();
            Thread.Sleep(500);
            SetForegroundWindow(Advhwnd);
              int x1hwnd = FindWindowEx(Advhwnd, 0, null, "1x1");
              int x4hwnd = FindWindowEx(Advhwnd, 0, null, "4x4");
              int x2hwnd = FindWindowEx(Advhwnd, 0, null, "2x2");
              int x3hwnd = FindWindowEx(Advhwnd, 0, null, "3x3");
           
             if ((int)numericUpDown23.Value == 1)
             {
                 SetForegroundWindow(x1hwnd);
                 PostMessage(x1hwnd, BN_CLICKED, 0, 0);
              }
             if ((int)numericUpDown23.Value == 2)
             {
                 SetForegroundWindow(x2hwnd);
                 PostMessage(x2hwnd, BN_CLICKED, 0, 0);  
             }
             if ((int)numericUpDown23.Value == 3)
             {
                 SetForegroundWindow(x3hwnd);
                 PostMessage(x3hwnd, BN_CLICKED, 0, 0);
             }
             if ((int)numericUpDown23.Value == 4)
             {
                 SetForegroundWindow(x4hwnd);
                 PostMessage(x4hwnd, BN_CLICKED, 0, 0);
              }
             Log("2-" + x2hwnd + "3-" + x3hwnd + "4-" + x4hwnd);
            SetForegroundWindow(Advhwnd);
            SendKeys.SendWait("!d");
            SendKeys.Flush();
            Thread.Sleep(1000);
            
        }
 */
        //filterfocushere
        private void FilterFocus()
        {
            try
            {
                
                 if ((IsServer()) & (checkBox8.Checked == true))
                    SlaveFocus();

                if ((radioButton14.Checked == false) && (radioButton15.Checked == false) && (radioButton16.Checked == false))
                {
                    MessageBox.Show("Must select PHD pause method", "scopefocus");
                    return;
                }
                FilterFocusOn = true;
                fileSystemWatcher4.EnableRaisingEvents = false;//**************added2_29
                /*  NebListenStop();   **************remd 2_29  next 1 rem as well
                  Thread.Sleep(500);
                  SetForegroundWindow(NebhWnd);
                  SendKeys.SendWait("~");
                  SendKeys.Flush();
                  Thread.Sleep(1000);
                 */
                editfound = 0;

                toolStripStatusLabel1.Text = "Filter Focus On";
                //   clientSocket = new TcpClient();
                FocusTime = (float)numericUpDown22.Value;
                //  if (radioButton12.Checked == true)\
                if (NebVNumber == 2)
                    FocusTime = FocusTime * 1000;
                //     textBox22.Text = FocusTime.ToString();
                this.Refresh();
                if (checkBox10.Checked == false) // slew to focus
                {
                    if (TargetLocObtained == false)
                    {
                        MessageBox.Show("Focus Coordinates not set", "scopefocus");
                        return;
                    }
                    if (FocusLocObtained == false)
                    {
                        MessageBox.Show("Target Coordinates not set", "scopefocus");
                        return;
                    }

                }
                /*
                    FocusTime = (float)numericUpDown22.Value;
                    if (radioButton12.Checked == true)
                        FocusTime = FocusTime * 1000;
                 */
                if (checkBox10.Checked == true)//focus star within target frame (no slew needed) 
                {


                    if (checkBox22.Checked == true)//******* try adding 6-29 for metric focus
                        gotoFocus();
                    else
                    {
                        SendFocusTime(FocusTime);
                        NebFineFocus();
                        Thread.Sleep(1000);
                        gotoFocus();
                    }



                }
                else    //need to slew to focus
                {
                    // SlewDelay = (int)numericUpDown32.Value * 1000;
                    //{pause phd, goto focus....(pause) --> Frame/focus --> fine focus(autoclick center) --> goto focus-->
                    // Abort neb--> goto target --> resume PHD} --> capture series
                    StopPHD();
                    Thread.Sleep(500);
                    toolStripStatusLabel1.Text = "Slewing to Focus Location";
                    this.Refresh();
                    //   while (MountMoving == true)
                    GotoFocusLocation();  //  *****rem for debugging******
                    // MessageBox.Show("simulate slew");
                    while (MountMoving == true)
                    {
                        Thread.Sleep(50); // pause for 1/20 second
                        System.Windows.Forms.Application.DoEvents();
                    }

                    //    Thread.Sleep(SlewDelay);
                    // CheckForSlewDone();
                    if (MountMoving == false)
                    {
                        button33.Text = "At Focus";
                        button33.BackColor = System.Drawing.Color.Lime;
                        button35.Text = "Goto Target";
                        button35.UseVisualStyleBackColor = true;
                        Thread.Sleep(1000);
                        SendFocusTime(FocusTime);
                        Thread.Sleep(1000);
                        NebFineFocus();//has 1 sec delay at end
                        Thread.Sleep(1000);
                        //   clientSocket = null;*************remd 2_29??not sure if needed, not tested yet
                        gotoFocus();

                        //    ResumePHD();
                    }




                }
            }
            catch(Exception ex)
            {
                Log("FilterFocus Error" + ex.ToString());
                Send("FilterFocus Error" + ex.ToString());
                FileLog("FilterFocus Error" + ex.ToString());

            }
  
        }
        /*
        private void FocusDoneTargetReturn()
        {
         //   Thread.Sleep(1000);
        //    SetForegroundWindow(Aborthwnd);
       //     PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
            Thread.Sleep(1000);
          //  MessageBox.Show("simulate Slew");
            toolStripStatusLabel1.Text = "Focus Done Slewing to Target";
            this.Refresh();
            GotoTargetLocation();   // *****rem for debugging*****
        //    Thread.Sleep(SlewDelay);
            ResumePHD();
        }

*/

 
            [DllImport("user32.dll")]
            public static extern int PostMessage(int hWnd, uint Msg, int wParam, int lParam);
            [DllImport("User32.dll")]
            public static extern Int32 FindWindow(String lpClassName, String lpWindowName);
            [DllImport("User32.dll")]
            public static extern Int32 SetForegroundWindow(int hWnd);
            [DllImport("User32.dll")]
            public static extern Boolean EnumChildWindows(int hWndParent, Delegate lpEnumFunc, int lParam);
            [DllImport("User32.dll")]
            public static extern Int32 GetWindowText(int hWnd, StringBuilder s, int nMaxCount);
            [DllImport("User32.dll")]
            public static extern Int32 GetClassName(int hWnd, StringBuilder s, int nMaxCount);//added to try to get edit box 
            [DllImport("User32.dll")]
            public static extern Int32 GetWindowTextLength(int hwnd);
            [DllImport("user32.dll", EntryPoint = "GetDesktopWindow")]
            public static extern int GetDesktopWindow();
            [DllImport("user32.dll")]
            public static extern bool SetCursorPos(int X, int Y);
            [DllImport("user32.dll")]
            public static extern int ChildWindowFromPointEx(int hWnd, Point pt, uint uFlags);
           
            public delegate int Callback(int hWnd, int lParam);


            public int hwndDuraton2;
            public int LoadScripthwnd2;
            public int Framehwnd2;
            public int Finehwnd2;
            public int Advhwnd2;
            public int Advancedhwnd2;
            public int Aborthwnd2;
            public int hwndDuration2;
            public int panelhwnd2;
        public int SlaveStatushwnd;
        public int Gotofocushwnd;//slave mode handles(2nd instance of scopefocus)
        public int Capturehwnd;
        public int Flathwnd;
        public int Pausehwnd;
        public int EnumChildGetValue(int hWnd,int lParam)
        {
            
                StringBuilder formDetails = new StringBuilder(256);
                StringBuilder formClass = new StringBuilder(256);
                int txtValue;
                int txtValue2;
                string editText = "";
                string classtext = "";
                txtValue = GetWindowText(hWnd, formDetails, 256);
                editText = formDetails.ToString().Trim();

                txtValue2 = GetClassName(hWnd, formClass, 256);
                classtext = formClass.ToString().Trim();
                if (lParam == 0)
                {
                    if (editText == "panel")//doesn't work w/ msctls_statusbar32  either
                    {
                        panelhwnd = hWnd;

                    }
                    if (classtext == "Edit")
                    {
                        editfound++;
                        if (editfound == 2)
                        {
                            hwndDuration = hWnd;
                        }
                    }

                    if (editText == "Advanced")
                    {

                        Advancedhwnd = hWnd;

                    }
                    //*****************************need to test after writing auto camera find stuff ***********************************         
                    if (editText == NebCamera + " Setup")
                    {
                        if (setupWindowFound == false)//picks first one
                        {
                            Advhwnd = hWnd;
                            //   Log("Adv" + Advhwnd.ToString());
                            setupWindowFound = true;
                        }
                    }

                    if (editText == "Abort")
                    {

                        Aborthwnd = hWnd;
                        //   Log("Abort " + Aborthwnd.ToString());
                    }
                    if (editText == "Frame and Focus")
                    {

                        Framehwnd = hWnd;
                        // Log("Frame " + Framehwnd.ToString());
                    }
                    if (editText == "Fine Focus")
                    {

                        Finehwnd = hWnd;
                        //  Log("Fine " + Finehwnd.ToString());
                    }
                    if (editText == "Load script")//not finding it
                    {
                        LoadScripthwnd = hWnd;
                        //   Log("LoadScript " + LoadScripthwnd.ToString());
                    }
                }
                if (lParam == 2)//added to find handles of second instence of Neb
                {
                    if (editText == "panel")//doesn't work w/ msctls_statusbar32  either
                    {
                        panelhwnd2 = hWnd;

                    }
                    if (classtext == "Edit")
                    {
                        editfound++;
                        if (editfound == 2)
                        {
                            hwndDuration2 = hWnd;
                        }
                    }

                    if (editText == "Advanced")
                    {

                        Advancedhwnd2 = hWnd;

                    }
                    //*****************************need to test after writing auto camera find stuff ***********************************         
                    if (editText == NebCamera + " Setup")
                    {
                        if (setupWindowFound == false)//picks first one
                        {
                            Advhwnd2 = hWnd;
                            //   Log("Adv" + Advhwnd.ToString());
                            setupWindowFound = true;
                        }
                    }

                    if (editText == "Abort")
                    {

                        Aborthwnd2 = hWnd;
                        //   Log("Abort " + Aborthwnd.ToString());
                    }
                    if (editText == "Frame and Focus")
                    {

                        Framehwnd2 = hWnd;
                        // Log("Frame " + Framehwnd.ToString());
                    }
                    if (editText == "Fine Focus")
                    {

                        Finehwnd2 = hWnd;
                        //  Log("Fine " + Finehwnd.ToString());
                    }
                    if (editText == "Load script")//not finding it
                    {
                        LoadScripthwnd2 = hWnd;
                        //   Log("LoadScript " + LoadScripthwnd.ToString());
                    }
                }

            if (checkBox20.Checked == true)
            {
                if (editText == "GotoFocus")
                {
                    Gotofocushwnd = hWnd;
                    Log("Slave Gotofocus handle found  " + Gotofocushwnd.ToString());
                }
                if (editText == "Capture")
                {
                    Capturehwnd = hWnd;
                    Log("Slave Capture Handle Found  " + Capturehwnd.ToString());

                }
                if (editText == "Flat")
                {
                    Flathwnd = hWnd;
                    Log("Slave Flat handle found  " + Flathwnd.ToString());
                }
                if (editText == "Pause")
                {
                    Pausehwnd = hWnd;
                    Log("SlavePause Handle found  " + Pausehwnd.ToString());
                }
                if (editText == "Not Connected")
                {
                    SlaveStatushwnd = hWnd;
                    Log("Status Handle " + SlaveStatushwnd.ToString());
                }

                }
                




                //MessageBox.Show("Contains text of control "+ editText);
                return 1;
           
        }
        
       private const int SW_SHOW = 5;
        private const int SW_SHOWMAXIMIZED = 3;
        private const int SW_RESTORE = 9;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(int hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(int hWnd, int nCmdShow);
       //neblistenstarthere
        private void NebListenStart(int hwnd, int port)//was blank
        {
            try
            {
                if (IsSlave())
                    Thread.Sleep(4000);//*****6-17 was 3000
                ShowWindowAsync(hwnd, SW_SHOW);
                Thread.Sleep(200);
                ShowWindowAsync(hwnd, SW_RESTORE);
                Thread.Sleep(200);
            //    ShowWindow(hwnd, SW_SHOW);
                Thread.Sleep(250);
                SetForegroundWindow(hwnd);//was NebhWnnd;
                Thread.Sleep(500);
                SendKeys.SendWait("^r");
                SendKeys.Flush();
                Thread.Sleep(1000);
                LoadScripthwnd = FindWindow(null, "Load script");
            //    GetHandles();
                if (LoadScripthwnd == 0)
                {
                    ShowWindowAsync(hwnd, SW_SHOW);//was working without this line
                    ShowWindowAsync(hwnd, SW_RESTORE);
                    Thread.Sleep(200);
                    //    ShowWindow(hwnd, SW_SHOW);
                    Thread.Sleep(250);
                    SetForegroundWindow(hwnd);//was NebhWnnd;
                    Thread.Sleep(500);
                    SendKeys.SendWait("^r");
                    SendKeys.Flush();
                    Thread.Sleep(1000);
                    LoadScripthwnd = FindWindow(null, "Load script");
                }
                //***********  added 6-7 so neblisten doesnt start after flat
                /*
                if (LoadScripthwnd == 0)
                {
                    Log("Neb Listen failed");
                    return;
                }
*/
                //******
                SetForegroundWindow(LoadScripthwnd);
                //   int indexhandle = FindWindowByIndex(LoadScripthwnd, 1, "ComboBoxEx32");
                //   Log("Indexhandle " + indexhandle.ToString());
                //   EnumChildGetValue(LoadScripthwnd, 0);
                hwndChild = FindWindowEx(LoadScripthwnd, 0, "ComboBoxEx32", null);
              //  Log("cbex32" + hwndChild.ToString());
                hwndChild1 = FindWindowEx(hwndChild, 0, "ComboBox", null);
             //   Log("cb" + hwndChild1.ToString());
                hwndChild2 = FindWindowEx(hwndChild1, 0, "edit", null);
             //  Log("enterscript name box" + hwndChild2.ToString());

                if (checkBox14.Checked == true)
                {
                    port = 4302;
                    ScriptName = "\\listenPort2.neb";
                }
                else
                {
                    port = 4301;
                    ScriptName = "\\listenPort.neb";
                }
                //need to get combobox then edit.....
                //  string send = "";
                // string send = NebPath + @"\listenPort.neb";
                //   send = "C:\\Program Files (x86)\\Nebulosity3\\listenPort.neb";
             //   string ScriptName = "\\listenPort.neb";
                string send = NebPath + ScriptName;

                StringBuilder sb = new StringBuilder(send);
                //  sb = send;
                SendMessage(hwndChild2, WM_SETTEXT, 0, sb);
                Thread.Sleep(1000);//******5-29 was 250
                SendKeys.SendWait("~");
                SendKeys.Flush();
                Thread.Sleep(1000);
                
                if (LoadScripthwnd != 0)
                {
                    if ((IsServer()) & (SlaveFlatOn == false))
                    {
                        
                        Log("Waiting for Slave connection");
                        toolStripStatusLabel1.Text = "Waiting for slave connection";
                       // this.Refresh();
                       while (working == true)
                        {

                            WaitForSequenceDone("Waiting for connection", NebSlavehwnd);
                          //  Thread.Sleep(50);
                        }
                        
                        
                        working = true;
                        Log("Slave Connected");
                    }
                }

                if (FineFocusAbort == true)
                {
                    while (working == true)
                    {
                        WaitForSequenceDone("Waiting for connection", NebhWnd);
                       // Thread.Sleep(100);
                       // Log("waiting");
                    }
                    working = true;
                    FineFocusAbort = false;
                }
                /*//********remd 6-6
                //***********************add some while Fcn to wait for neb status to say "Waiting for connection"
                bool LoadScript = true;
                if (IsServer())
                {
                    Log("Waiting for Slave load script");
                    while (LoadScript == true)
                    {
                        int StatusstripHandle = FindWindowEx(NebSlavehwnd, 0, "msctls_statusbar32", null);

                        //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                        IntPtr statusHandle = new IntPtr(StatusstripHandle);
                        StatusHelper sh = new StatusHelper(statusHandle);
                        string[] captions = sh.Captions;

                        if (captions[0] == "Waiting for connection")
                        {
                            Log("Slave Load Script Complete");
                            LoadScript = false;
                        }
                        Thread.Sleep(50);
                    }
                }
*/
                if ((clientSocket.Connected == false) & (port == 4301))
                    Connect1(port);

                if ((clientSocket2.Connected == false) & (port == 4302))
                    Connect2(port);
/*
                if (clientSocket.Connected == false)
                {
                    Connect1(port);
                    Log("connecting to " + port.ToString());
                }
*/
                /*
                if (FineFocusAbort == true)
                {
                    Log("waiting for connection");
                    while (working == true)
                    {

                        WaitForSequenceDone("New cnxn", NebhWnd);
                          Thread.Sleep(50);

                    }
                    working = true;
                }
                 */
                NebListenOn = true;  
             }
            catch(Exception ex)
            {
                Log("NebListenStart Error" + ex.ToString());
                Send("NebListenStart Error" + ex.ToString());
                FileLog("NebListenStart Error" + ex.ToString());

            }
                

        }
        private void Connect2(int port)
        {
            try
            {
               // if (clientSocket2.Connected == false)//**************try adding 2_29
               // {
                    clientSocket2 = new TcpClient();
                    LingerOption lingerOption = new LingerOption(true, 0);
                    clientSocket2.LingerState = lingerOption;
                    clientSocket2.Connect("127.0.0.1", port);//*************try adding and red above)
                    Thread.Sleep(1000);
                    Log(port.ToString() + " Conneceted");
              //  }
            }
            catch (Exception e)
            {
                Log("Connect2 Error" + e.ToString());
                Send("Connect2 Error" + e.ToString());
                FileLog("Connect2 Error" + e.ToString());

            }
        }
        
        private void Connect1(int port)
        {
            try
            {
              //  if (clientSocket.Connected == false)//**************try adding 2_29
              //  { 
                    clientSocket = new TcpClient();
                    LingerOption lingerOption = new LingerOption(true, 0);
                    clientSocket.LingerState = lingerOption;
                    clientSocket.Connect("127.0.0.1", port);//*************try adding and red above)
                    Thread.Sleep(1000);
                    Log(port.ToString() + " Conneceted");
             //   }
            }
            catch( Exception e)
            {
                Log("Connect Error Line 6343" + e.ToString());
                Send("Connect Error Line 6343" + e.ToString() );
                FileLog("Connect Error Line 6343" + e.ToString());

            }
        }
        
//Neb Listen Start button...maybe not needed
        /*
        private void button30_Click_1(object sender, EventArgs e)
        {
            NebListenStart();
        }
        */

        //********************added 3_13 try to get window style for fine focus panel  ************
        
        [DllImport("user32")]
        private static extern int IsWindowVisible(int hwnd);
       
        //********************** end addition
        public int FFfailed = 0;
        private void NebFineFocus()//abort - frame - abort - fine - click star at position
        {
            try
           {
                toolStripStatusLabel1.Text = "Neb Fine Focus On";
                this.Refresh();
                SetForegroundWindow(NebhWnd);
                Thread.Sleep(3000);

                PostMessage(Aborthwnd, BN_CLICKED, 0, 0);//*******added 3_13
                Thread.Sleep(3000);

                PostMessage(Framehwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(3000);

                PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(3000);

                PostMessage(Finehwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(5000);


                Point xxx = new Point();
                xxx.X = Convert.ToInt16(textBox29.Text.ToString());//default is center
                xxx.Y = Convert.ToInt16(textBox30.Text.ToString());
                int starpos = ((xxx.Y << 16) | (xxx.X & 0xffff));

                int panelhwnd3 = FindWindowByIndexName(NebhWnd, 3, "panel");//index(2) may change w. neb versions???
                int panelhwnd2 = FindWindowByIndexName(NebhWnd, 2, "panel");
            //******************* added 3_13
                int panel2Vis = IsWindowVisible(panelhwnd2);
                if (panel2Vis == 1)
                {
                    panelhwnd = panelhwnd2;
                  //  Log("panel2vis " + panel2Vis.ToString() + " # " + panelhwnd.ToString());
                }
                else
                {
                    panelhwnd = panelhwnd3;
                }
                int panel3Vis  = IsWindowVisible(panelhwnd3);//WORKS  visable =1
               // Log("panel3vis " + panel3Vis.ToString() + " # " + panelhwnd.ToString());
                Log("pos = " + starpos.ToString() + "  X = " + xxx.X.ToString() + "  Y = " + xxx.Y.ToString());

         

//*****************end addition
                SetForegroundWindow(panelhwnd);
              //  Log("panel2 for fine Focus" + panelhwnd.ToString());//index 3 now for some reason(3/9/12)
                Thread.Sleep(500);
                PostMessage(panelhwnd, WM_LBUTTONDOWN, 0, starpos);//was SendMessage
                PostMessage(panelhwnd, WM_LBUTTONUP, 0, starpos);//was SendMessage
              //  SendMessage(panelhwnd, WM_LBUTTONDOWN, 0, starpos);//was SendMessage
             //   SendMessage(panelhwnd, WM_LBUTTONUP, 0, starpos);//was SendMessage
                Thread.Sleep(1000);
                /*

                //*******************below is untested try coord correction for drift if no star************
                //currently this does nothing 
                string[] filePaths = Directory.GetFiles(path2.ToString(), "*.bmp");
                roundto = (int)numericUpDown1.Value;
                int current = GetFileHFR(filePaths, roundto);
                Log ("test HFR = " + current.ToString());
                
                if (current > enteredMaxHFR)
                {
                    Log("Focus star outside V-curve range");
                    if (FFfailed == 1)
                    {
                        Log("Finefocus failed  - no star selected - Aborted");
                        Send("Finefocus failed  - no star selected - Aborted");
                        FileLog("Finefocus failed  - no star selected - Aborted");
                        standby();

                    }
                    
                }
                 */
            }
                       


            catch(Exception ex)
            {
                Log("NebFineFocus Error" + ex.ToString());
                Send("NebFineFocus Error" + ex.ToString());
                FileLog("NebFineFocus Error" + ex.ToString());

            }
             
        }



  
        private void NebListenStop()
        {
            try
            {
                // try not disconnecting socket 
                /*
               if (firstconnect == true)//if already connected once, this closes and re-connects on repeat "Go" push
               {
                   clientSocket.Close();

                   clientSocket = new TcpClient();
                   clientSocket.Connect("127.0.0.1", 4301);//connects to neb
               }

               if (firstconnect == false)// establish first time connection
               {

                   clientSocket.Connect("127.0.0.1", 4301);//connects to neb
                   firstconnect = true;
               }
                */
              //  NetworkStream serverStream = clientSocket.GetStream();
                serverStream = clientSocket.GetStream();
                byte[] outStream3 = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                serverStream.Write(outStream3, 0, outStream3.Length);
                serverStream.Flush();
                outStream3 = null;
            }
            catch(Exception ex)
            {
                Log("NebListenStop Error" + ex.ToString());
                Send("NebListenStop Error" + ex.ToString());
                FileLog("NebListenStop Error" + ex.ToString());

            }
        }
/*
      //Stop Neb Listen
       private void button38_Click(object sender, EventArgs e)
       {
          NebListenStop();
       }

*/
       



     //  [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    //   static extern int GetWindowText(int hWnd, StringBuilder lpString, int nMaxCount);
       [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
       static extern int GetWindowTextLength(IntPtr hWnd);
       [DllImport("USER32.DLL"),]
       private static extern int FindWindowEx(int parentHandle, int childAfter, string className, string windowTitle);
       [DllImport("USER32.DLL", CharSet = CharSet.Auto)]
       public static extern int SendMessage(int hWnd, int msg, int wparam, StringBuilder text);
       [DllImport("USER32.DLL", CharSet = CharSet.Auto)]
       public static extern int SendMessage(int hWnd, int msg, int wparam, int lparam);
       [DllImport("USER32.DLL"),]
       private static extern int GetWindowRect(int hWnd, out RECT lpRect);
       [StructLayout(LayoutKind.Sequential)]
        public struct RECT
       {
           public int left;
           public int top;
           public int right;
           public int bottom;
       }


//try to get cursor position from neb*****doesn't work****
       private void GetNebCursor()
       {
           try
           {
               string strTemp;
               int res;
               res = SendMessage(132662, SB_GETTEXTLENGTH, 2, null);//this works! gets length of part 2 (curser position)
               int len = (int)res;
               //  len = (len & 0x0000ffff) * 2;//from phd emailer
               Log(len.ToString());
               int parts;
               parts = SendMessage(132662, SB_GETPARTS, 0, null);//gets 4 parts--works
               Log(parts.ToString());
               //  int setparts = SendMessage(459724, SB_SETPARTS, 3, null);
               //  Log(setparts.ToString());
               StringBuilder sb = new StringBuilder(len);//was (len+1) not sure why
               strTemp = sb.ToString();
               int x = strTemp.Length;
               res = SendMessage(132662, SB_GETTEXT, 2, sb);//not getting text from part 2
               Log(sb.ToString());

               //******see phd emailer  nebstatushelper line 196  ******
               //may need to ??? readProcessMemory????
               //also need to find handle, try using: msctls_statusbar32 

               /*


               //GetHandles();
            //   int FFdonehwnd = FindWindowEx(NebhWnd, 1, "", "Frame and focus done");
               SetForegroundWindow(NebhWnd);
               int length       = GetWindowTextLength(NebhWnd);
               StringBuilder sb = new StringBuilder(length + 1);

               GetWindowText(590930, sb, sb.Capacity);
               Log("FFdone " + length.ToString() + "  " + sb.ToString()); 
               MessageBox.Show(sb.ToString());
              */
           }
           catch(Exception ex)
           {
               Log("GetNebCuror Error" + ex.ToString());
               Send("GetNebCuror Error" + ex.ToString());
               FileLog("GetNebCuror Error" + ex.ToString());

           }
       }

       static int FindWindowByIndex(int hWndParent, int index, string name)
       {
           if (index == 0)
               return hWndParent;
           else
           {
               int ct = 0;
               int result = 0;
               do
               {
                   result = FindWindowEx(hWndParent, result, name, null);//name is class name
                   if (result != 0)
                       ++ct;
               }
               while (ct < index && result != 0);
               return result;
           }
       }

       static int FindWindowByIndexName(int hWndParent, int index, string name)//here name is window title
       {
           if (index == 0)
               return hWndParent;
           else
           {
               int ct = 0;
               int result = 0;
               do
               {
                   result = FindWindowEx(hWndParent, result, null, name);
                   if (result != 0)
                       ++ct;
               }
               while (ct < index && result != 0);
               return result;
           }
       }

       private void StopPHD()
       {
           try
           {
               if (radioButton14.Checked == true || radioButton16.Checked == true)
               {
                   SetForegroundWindow(PHDhwnd);
                   int phdchildStop = FindWindowByIndex(PHDhwnd, 5, "button");
                   Log("PHD stop button " + phdchildStop.ToString());
                   SetForegroundWindow(phdchildStop);
                   Thread.Sleep(500);
                   PostMessage(phdchildStop, BN_CLICKED, 0, 0);
               }
               if (radioButton15.Checked == true)
               {
                   PHDSocketPause(true);
                   
               }
           }
           catch(Exception ex)
           {
               Log("StopPHD Error" + ex.ToString());
               Send("StopPHD Error" + ex.ToString());
               FileLog("StopPHD Error" + ex.ToString());
           }

       }


       private void PHDAutoStar()
       {
           try
           {
               SetForegroundWindow(PHDhwnd);
               SendKeys.SendWait("%s");
               SendKeys.Flush();

               int phdcapture = FindWindowByIndex(PHDhwnd, 3, "button");
               Log("PHD capture button " + phdcapture.ToString());
               SetForegroundWindow(phdcapture);

               PostMessage(phdcapture, BN_CLICKED, 0, 0);
               Thread.Sleep(2000);

               int phdguide = FindWindowByIndex(PHDhwnd, 4, "button");
               Log("PHD guide button " + phdguide.ToString());
               SetForegroundWindow(phdguide);
               Thread.Sleep(500);
               PostMessage(phdguide, BN_CLICKED, 0, 0);
           }
           catch(Exception ex)
           {
               Log("PHDAutostar Error" + ex.ToString());
               Send("PHDAutostar Error" + ex.ToString());
               FileLog("PHDAutostar Error" + ex.ToString());
           }
       }

        private void PHDCoordStar()
        {
            try
            {
                //This works w/ auto star select.   below is attempt to select star based on coords..works using screen coords, need
                   // to change to active window coords.

                   SetForegroundWindow(PHDhwnd);
                   int phdcapture = FindWindowByIndex(PHDhwnd, 3, "button");
                   Log("PHD capture button " + phdcapture.ToString());
                   SetForegroundWindow(phdcapture);

                   PostMessage(phdcapture, BN_CLICKED, 0, 0);
                   Thread.Sleep(1000);//time to place cursor
                   Point pt = new Point();
                   GetCursorPos(ref pt);

                   RECT rc;
                   //  Rectangle rect = new Rectangle();


                   //***************WORKS!*************
                   int phdpanel = FindWindowEx(PHDhwnd, 0, null, "panel");
                   Log("Phd panel " + phdpanel.ToString());
                   SetForegroundWindow(phdpanel);
                   GetWindowRect(phdpanel, out rc);//find upper left corner to reference location on screen.  
                   //allows for converion of position to "acitve window"
                   //      textBox35.Text = rc.left.ToString();
                   //      textBox36.Text = rc.top.ToString();
                   Point xxx = new Point();
                   //  xxx.X = Convert.ToInt16(textBox33.Text.ToString()) - Convert.ToInt16(textBox35.Text.ToString());
                   //   xxx.Y = Convert.ToInt16(textBox34.Text.ToString()) - Convert.ToInt16(textBox36.Text.ToString());
                   xxx.X = pt.X - rc.left;//these 2 line replace above 2 
                   xxx.Y = pt.Y - rc.top;

                   textBox31.Text = xxx.X.ToString();
                   textBox32.Text = xxx.Y.ToString();
                   int guidestarpos = ((xxx.Y << 16) | (xxx.X & 0xffff));
                   SetCursorPos(xxx.X, xxx.Y);

                   Thread.Sleep(2000);
                   PostMessage(phdpanel, WM_LBUTTONDOWN, 0, guidestarpos);//was SendMessage
                   PostMessage(phdpanel, WM_LBUTTONUP, 0, guidestarpos);//was SendMessage
                   Thread.Sleep(2000);

                   //Check into... leaving the cursor on the guide star and always using that position 
                   // OR...set it once and make sure the PHD window doesn't move then use those values.  



                   int phdguide = FindWindowByIndex(PHDhwnd, 4, "button");
                   Log("PHD guide button " + phdguide.ToString());
                   SetForegroundWindow(phdguide);
                   Thread.Sleep(500);
                   PostMessage(phdguide, BN_CLICKED, 0, 0);

               }
            catch(Exception ex)
            {
                Log("PHDCoordStar Error" + ex.ToString());
                Send("PHDCoordStar Error" + ex.ToString());
                FileLog("PHDCoordStar Error" + ex.ToString());
            }

        }
       private void ResumePHD()
       {
           try
           {
               if (radioButton14.Checked == true)//auto star= hit stop, slew to focus, slew back then hit guide
               {
                   PHDAutoStar();
               }

               if (radioButton16.Checked == true)//use coordinates to select star , stop, slew, slew back, capture, select star, guide
               {
                   PHDCoordStar();
 //*************************untested***************** 4-9-12
                   Thread.Sleep(4000);//wait then check to make sure ther is no LOW SNR msg (see  below)
                   int StatusstripHandle = FindWindowEx(PHDhwnd, 0, "msctls_statusbar32", null);

                   //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                   IntPtr statusHandle = new IntPtr(StatusstripHandle);
                   StatusHelper sh = new StatusHelper(statusHandle);
                   string[] captions = sh.Captions;
 //******************not sure of exact messages....may just be albe to look for "guiding" OR
                   //may say "guiding" even if no star, thus need to look for LOW SNR"
                      if (captions[0] == "Guiding")//******must rem for sim debugging *****
                   if (captions[1] != null)
                       if (captions[1].Substring(0, 3) == "LOW")
                       {
                           radioButton15.Checked = false;
                           radioButton14.Checked = true;
                           ResumePHD();
                       }
               }
                   

               if (radioButton15.Checked == true)
               {
                   
                   PHDSocketPause(false);

 //*******************untested************** 4-9-12 (see comments above)
                   Thread.Sleep(4000);//wait 4 seconds then make sure there is a guide star
                   //if no guidestar, repeat using autostar select.  
                   //will catch if no star after that
                   int StatusstripHandle = FindWindowEx(PHDhwnd, 0, "msctls_statusbar32", null);

                   //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                   IntPtr statusHandle = new IntPtr(StatusstripHandle);
                   StatusHelper sh = new StatusHelper(statusHandle);
                   
                   string[] captions = sh.Captions;
                    if (captions[0] == "Guiding")//******must rem for sim debugging *****
                   if (captions[1] != null)
                       if (captions[1].Substring(0, 3) == "LOW")
                       {
                           radioButton15.Checked = false;
                           radioButton14.Checked = true;
                           ResumePHD();
                     }
//********also there may be new addition to send autostar by socket command*****************                  
               }

           }
           catch(Exception ex )
           {
               Log("ResumePHD Error" + ex.ToString());
               Send("ResumePHD Error" + ex.ToString());
               FileLog("ResumePHD Error" + ex.ToString());
           }
       }

        public static int PHD_PAUSE = 1;
        public static int PHD_RESUME = 2;

        
//**************this works!*********************
        private void PHDSocketPause(bool pause)
        {
            try
            {
                phdsocket = new TcpClient();
                phdsocket.Connect("127.0.0.1", 4300);

                int pauseCommand = PHD_PAUSE;
                if (!pause)
                    pauseCommand = PHD_RESUME;
                byte[] buf = new byte[1];
                buf[0] = (byte)((char)pauseCommand);
                phdsocket.Client.Send(buf);
                phdsocket.Client.Receive(buf);
            }
//*************may want to add this
            //  - (1.13.5) New server command (id=14) to auto-find star (same as pulling down this from the menu)

            catch(Exception ex)
            {
                Log("PHDSocketPause Error" + ex.ToString());
                Send("PHDSocketPause Error" + ex.ToString());
                FileLog("PHDSocketPause Error" + ex.ToString());
            }
        }

 //this is to show real time cursor position....prob not needed      
       [DllImport("user32.dll")]
       static extern bool GetCursorPos(ref Point lpPoint);

       private void timer1_Tick(object sender, EventArgs e)
       {
           
           Point pt = new Point();
           GetCursorPos(ref pt);
        //   textBox33.Text = pt.X.ToString();
       //    textBox34.Text = pt.Y.ToString();
            
       }
//neb ver 2.5
 
        public static IntPtr SearchForWindow(string wndclass, string title)
    {
        SearchData sd = new SearchData { Wndclass=wndclass, Title=title };
        EnumWindows(new EnumWindowsProc(EnumProc), ref sd);
        return sd.hWnd;
    }

    public static bool EnumProc(IntPtr hWnd, ref SearchData data)
    {
        // Check classname and title 
        // This is different from FindWindow() in that the code below allows partial matches
        StringBuilder sb = new StringBuilder(1024);
        GetClassName(hWnd, sb, sb.Capacity);
        if (sb.ToString().StartsWith(data.Wndclass))
        {
            sb = new StringBuilder(1024);
            GetWindowText(hWnd, sb, sb.Capacity);
            if (sb.ToString().StartsWith(data.Title))
            {
                data.hWnd = hWnd;
                return false;    // Found the wnd, halt enumeration
            }
        }
        return true;
    }

    public class SearchData
    {
        // You can put any vars in here...
        public string Wndclass;
        public string Title = "Nebulosity";
        public IntPtr hWnd;
    }

    private delegate bool EnumWindowsProc(IntPtr hWnd, ref SearchData data);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, ref SearchData data);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);



/* not needed finds automatically 

       private void radioButton12_CheckedChanged(object sender, EventArgs e)
       {
           
           if (radioButton12.Checked == true)
           {
              
               Callback myCallBack = new Callback(EnumChildGetValue);
               NebhWnd = FindWindow(null, "Nebulosity v2.5");
               PHDhwnd = FindWindow(null, "PHD Guiding 1.13.0b  -  www.stark-labs.com (Log active)");
               LoadScripthwnd = FindWindow(null, "Load script");
               Log("PHD " + PHDhwnd.ToString());
               Log("Load script " + LoadScripthwnd.ToString());

               //  SetForegroundWindow(NebhWnd);
               if (NebhWnd == 0)
               {

                   MessageBox.Show("Ensure Nebulosity Ver 2 is running", "scofocus");

               }
               else
               {

                   EnumChildWindows(NebhWnd, myCallBack, 0);
               }
           }
            
       }

        
//Neb ver 3
       private void radioButton13_CheckedChanged(object sender, EventArgs e)
       {
           
         if (radioButton13.Checked == true)
           {
               Callback myCallBack = new Callback(EnumChildGetValue);

               NebhWnd = FindWindow(null, "Nebulosity v3.0-a6");
               PHDhwnd = FindWindow(null, "PHD Guiding 1.13.0b  -  www.stark-labs.com (Log active)");
               LoadScripthwnd = FindWindow(null, "Load script");
               Log("PHD " + PHDhwnd.ToString());
               Log("Load script " + LoadScripthwnd.ToString());

               //  SetForegroundWindow(NebhWnd);
               if ((NebhWnd == 0))
               {

                   MessageBox.Show("Ensure Nebulosity Ver 3 is running ", "scopefocus");

               }
               else
               {

                   EnumChildWindows(NebhWnd, myCallBack, 0);
               }
           }
            
       }
*/
        /*
       private void button29_Click(object sender, EventArgs e)
       {
              NebListenStart();
           
       }

       private void button39_Click(object sender, EventArgs e)
       {
           NebListenStop();
       }
       */    
      //export to excel  
       private void button42_Click(object sender, EventArgs e)
       {
          
       }

        private void radioButton15_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton15.Checked == true)
                MessageBox.Show("Ensure PHD Server is enabled", "scopefocus");
        }

        private void numericUpDown22_ValueChanged(object sender, EventArgs e)
        {
            FocusTime = (float)numericUpDown22.Value;
         
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button31_Click(object sender, EventArgs e)
        {
            if (paused == true)
            {
                button12.PerformClick();
                return;
            }
            SequenceGo();
            button31.BackColor = System.Drawing.Color.Lime;
        }
        
        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            CountFilterTotal();
            if (checkBox9.Checked == false)
            {
                numericUpDown29.Value = 0;
                numericUpDown30.Value = 0;
            }
            if (checkBox5.Checked == false)
            {
                MessageBox.Show("Use Dark 5 for single dark frame", "scopefocus");
                checkBox9.Checked = false;
            }
        }
        

        private void textBox22_TextChanged(object sender, EventArgs e)
        {
         //   camera = textBox22.Text.ToString();
        }




        private static void Excel_FromDataTable(DataTable dt)
        {
              Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                // Create an Excel object and add workbook...
                //orig   ApplicationClass excel = new ApplicationClass();

                Workbook workbook = excel.Application.Workbooks.Add(true); // true for object template???

                // Add column headings...
                int iCol = 0;
                foreach (DataColumn c in dt.Columns)
                {
                    iCol++;
                    excel.Cells[1, iCol] = c.ColumnName;
                }
                // for each row of data...
                int iRow = 0;
                foreach (DataRow r in dt.Rows)
                {
                    iRow++;

                    // add each row's cell data...
                    iCol = 0;
                    foreach (DataColumn c in dt.Columns)
                    {
                        iCol++;
                        excel.Cells[iRow + 1, iCol] = r[c.ColumnName];
                    }
                }

                // Global missing reference for objects not defined
                object missing = System.Reflection.Missing.Value;
            if (ExcelFilename.IndexOf(@".") == 0)//if extension not on filename, add it. 
                workbook.SaveAs(ExcelFilename + ".xls",
                    XlFileFormat.xlExcel7, missing, missing,
                    false, false, XlSaveAsAccessMode.xlNoChange,
                    missing, missing, missing, missing, missing);
            else
                workbook.SaveAs(ExcelFilename,
                    XlFileFormat.xlExcel7, missing, missing,
                    false, false, XlSaveAsAccessMode.xlNoChange,
                    missing, missing, missing, missing, missing);

                // make Excel visible and activate the worksheet...
                excel.Visible = true;
                Worksheet worksheet = (Worksheet)excel.ActiveSheet;
                ((_Worksheet)worksheet).Activate();

                //shutdown excel...
                ((_Application)excel).Quit();
                MessageBox.Show("Data Exported to MyDocuments\\" + ExcelFilename + ".xls", "scopefocus");

                Thread.Sleep(1000);
                foreach (System.Diagnostics.Process process in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    if (process.MainModule.ModuleName.ToUpper().Equals("EXCEL.EXE"))
                    {
                        process.Kill();
                        break;
                    }
                }
            
        }

        
     public void importDataFromExcel()
        {
         
           string sSQLTable = "table1";
         //this works
           string myExcelDataQuery = "SELECT Date, PID, SlopeDWN, SlopeUP, Number, Equip, BestHFR, FocusPos FROM [Sheet1$]";
            try
            {
                if (ImportPath == null)
                {
                    DialogResult result = openFileDialog1.ShowDialog();
                    ImportPath = openFileDialog1.FileName.ToString();
                    textBox34.Text = ImportPath.ToString();
                }
                string sSqlConnectionString = conString;
                string sExcelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ImportPath + @";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""";
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

        private void button29_Click_1(object sender, EventArgs e)
        {
            
        }

               

        private void textBox34_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            ImportPath = openFileDialog1.FileName.ToString();
            textBox34.Text = ImportPath.ToString();
        }

        private void textBox33_Click(object sender, EventArgs e)
        {
            DialogResult result = saveFileDialog1.ShowDialog();
            ExcelFilename = saveFileDialog1.FileName.ToString();
            
            textBox33.Text = ExcelFilename.ToString(); 
        }



        private void FindNebCamera()
        {
            try
            {
                int panel3 = FindWindowByIndexName(NebhWnd, 4, "panel");// try panelhwnd
              //  Log("panel3 found" + panel3.ToString());
                int panel4 = FindWindowByIndexName(panel3, 1, "panel");//panel 4 is small one in panel 3, camera is 1st child in this
             //   Log("found panel4" + panel4);
                int camera = FindWindowByIndex(panel4, 1, null);
            //    Log("Found camera" + camera.ToString());
                StringBuilder sb = new StringBuilder(1024);
                SendMessage(camera, WM_GETTEXT, 1024, sb);
                //  GetWindowText(camera, sb, sb.Capacity);
                Log("Camera " + sb.ToString());
                NebCamera = sb.ToString();
                if (NebCamera == "No camera")
                {
                    NoCameraSelected();
                }
                textBox22.Text = NebCamera.ToString();
            }
            catch(Exception ex)
            {
                Log("FindNebCamera Error" + ex.ToString());
                Send("FindNebCamera Error" + ex.ToString());
                FileLog("FindNebCamera Error" + ex.ToString());

            }
        }
        
        private void NoCameraSelected()
        {
            DialogResult result;
            result = MessageBox.Show("No Camera Selected - Select Camera then push Retry", "scopefocus",
                            MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
            if (result == DialogResult.Retry)
                FindNebCamera();
            if (result == DialogResult.Cancel)
                return;
        }
        public int MaxADU;
        private void FindADU()
        {
           try
           {
                int ADUpanel = FindWindowByIndexName(NebhWnd, 5, "panel");// try panelhwnd
            //    Log("panel5 found" + ADUpanel.ToString());
                int MaxADUbox = FindWindowByIndex(ADUpanel, 2, "edit");//Finds second edit panel under parent panel under main window
            //    Log("found MaxADU box" + MaxADUbox.ToString());
            
               
                StringBuilder sb = new StringBuilder(1024);
                SendMessage(MaxADUbox, WM_GETTEXT, 1024, sb);
                //  GetWindowText(camera, sb, sb.Capacity);

                string ADU = sb.ToString();
                MaxADU = Convert.ToInt32(ADU);
               // Log("MaxADU: " + MaxADU.ToString() + "Exposure Time: " + FlatExp.ToString());
             
        }
        
            catch(Exception ex) 
            {
                Log("FindADU Error" + ex.ToString());
                Send("FindADU Error" + ex.ToString());
                FileLog("FindADU Error" + ex.ToString());

            }

        }
        public double ADUratio;
        public double tol;
        public double FlatExp;
        public int FlatGoal;
        public bool FlatCalcDone = false;
     //   public bool SlaveFlatCalcDone = false;
        private void CalculateFlatExp()//calcflatexphere
        {

            if (IsSlave())
                textBox41.Clear();
            toolStripStatusLabel1.Text = "Calculating Flat Exposure";
            this.Refresh();
            if ((IsSlave()) & (FlatExp == 0))//allows second set of flats to start at calculated time
                FlatExp = Convert.ToDouble(textBox40.Text)/1000;
            if (FlatExp == 0)
            FlatExp = (double)numericUpDown34.Value/1000;
            MaxADU = 0;

            fileSystemWatcher4.EnableRaisingEvents = false;
            /*
            if (NebVNumber == 3)
                CaptureTime3 = (double)CaptureTime / 1000;//need to allow fractions for flats.
           */
           // CaptureTime3 = FlatExp;
           

            tol = 1 + (Convert.ToDouble(textBox39.Text.ToString())/100);
            FlatGoal = Convert.ToInt16(textBox38.Text.ToString()) * 1000;
        //    if (NebListenOn == false)
            NebListenStart(NebhWnd, SocketPort);
            while ((MaxADU > FlatGoal * tol) || (MaxADU < FlatGoal / tol))
            {
                textBox41.Refresh();
              
              //  toolStripStatusLabel1.Text = "Capturing";
             //   this.Refresh();
                string prefix = textBox19.Text.ToString();
              //  NetworkStream serverStream;
                if (IsSlave())
                    serverStream = clientSocket2.GetStream();
                else
                    serverStream = clientSocket.GetStream();
               int subs = 1;
               string name = "FlatCalc";
                byte[] outStream = System.Text.Encoding.ASCII.GetBytes("setname " + prefix + name + "\n" + "setbinning " + CaptureBin + "\n" + "SetShutter 0" + "\n" + "SetDuration " + FlatExp + "\n" + "Capture " + subs + "\n");
                try
                {
                    Capturing = true;
                    serverStream.Write(outStream, 0, outStream.Length);
                    Thread.Sleep(500);//wait to check until after starts capturing 
                    while (Capturing == true)
                    {
                        int StatusstripHandle = FindWindowEx(NebhWnd, 0, "msctls_statusbar32", null);

                        //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                        IntPtr statusHandle = new IntPtr(StatusstripHandle);
                        StatusHelper sh = new StatusHelper(statusHandle);
                        string[] captions = sh.Captions;
                        if (captions[0] == "Sequence done")
                            Capturing = false;
                    }




                }
                   
                catch
                {
                    MessageBox.Show("Error sending command", "scopefocus");
                    return;

                }
               
                serverStream.Flush();
              //  Thread.Sleep(5000);//depends on download time.  may be a way to wait for change of MaxADU
             //   Thread.Sleep((int)FlatExp * 1000);

                
                FindADU();
                Log("MaxADU " + MaxADU.ToString());
                if ((MaxADU < FlatGoal * tol) && (MaxADU > FlatGoal / tol))//this gets out fo the loop before adusting exp time
                {
                    if (IsSlave())
                    {
                        textBox41.Text = "Slave FlatCalc Complete";
                        textBox41.Refresh();
                    }
                    Log("Flat Calc Complete--MaxADU: " + MaxADU.ToString() + "   Exposure Time: " + FlatExp.ToString());
                    
                    CaptureTime3 = FlatExp;
                    FlatCalcDone = true;
                    break;
                }
                ADUratio = (double)FlatGoal / (double)MaxADU;
                FlatExp = FlatExp * ADUratio;
                

            }
            /*
            if ((IsSlave()) & (checkBox21.Checked == true))
            {
                textBox41.Text = "Slave FlatCalc Complete";
                //textBox41.Refresh();
            }
            */
            if (IsServer())
            {

                while (working == true)
                {
     
                    StringBuilder sb = new StringBuilder(1024);
                    SendMessage(SlaveStatushwnd, WM_GETTEXT, 1024, sb);
                    //  GetWindowText(camera, sb, sb.Capacity);
                  //  Log("SlaveStatus= " + sb.ToString());
                    if (sb.ToString()== "Slave FlatCalc Complete")
                        working = false;
                    Thread.Sleep(20);
                }
                working = true; 
            }
                
          //  fileSystemWatcher4.EnableRaisingEvents = true;  //rem'd 5-19


            //******below moved to before break above*************
            /*
            Log("Flat Calc Complete--MaxADU: " + MaxADU.ToString() + "   Exposure Time: " + FlatExp.ToString());
            CaptureTime3 = FlatExp;
            FlatCalcDone = true;
             */
            //***************************************
        //    NebCapture();****rem'd 5-19
         //   NebListenStop();
          //  return;
         //   SlaveFlatCalcDone = false;
            
        }

        private void folderBrowserDialog2_HelpRequest(object sender, EventArgs e)
        {

        }

        private void textBox35_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog2.ShowDialog();
            NebPath = folderBrowserDialog2.SelectedPath.ToString();
            textBox35.Text = NebPath;
        }
        //std dev button
        private void Button2000_Click(object sender, EventArgs e)
        {
            port.DiscardInBuffer();
            port.DiscardOutBuffer();
            fileSystemWatcher2.EnableRaisingEvents = false;
            fileSystemWatcher5.EnableRaisingEvents = false; //added to test metricHFR
            fileSystemWatcher3.EnableRaisingEvents = true;
            fileSystemWatcher1.EnableRaisingEvents = false;
            list = new int[arraysize2];
            abc = new double[arraysize2];

            posMin = 0;
            min = 1;
            sum = 0;
            avg = 0;
            vDone = 0;
            vProgress = 0;
        }

        private void button12_Click_1(object sender, EventArgs e)
        {
            StopPHD();
        }

        private void button19_Click_1(object sender, EventArgs e)
        {
            ResumePHD();
        }

        private void button22_Click_1(object sender, EventArgs e)
        {
            StopTracking();
        }

        private void button30_Click(object sender, EventArgs e)
        {
            ResumeTracking();
        }

        private void ResumeTracking()
        {
            try
            {
                if (TrackingOn == true)
                {
                    MessageBox.Show("Tracking Already On", "scoprefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (port2.IsOpen == false)
                    Port2Open();
                port2.DiscardInBuffer();
                port2.DiscardOutBuffer();
                Thread.Sleep(50);
                byte[] newMsg = { 0x54, 0x02 };
                port2.Write(newMsg, 0, newMsg.Length);
                Thread.Sleep(100);
                port2.DiscardOutBuffer();

                Thread.Sleep(50);
                string Handshake = port2.ReadExisting();
                Thread.Sleep(50);
                if (Handshake == "#") //check for handshake
                {
                    Log("Tracking Command Received");
                }

                port2.DiscardInBuffer();
                port2.DiscardOutBuffer();
                Thread.Sleep(50);
                port2.Write("t");//check tracking mode
                Thread.Sleep(100);
                port2.DiscardOutBuffer();
                // byte[] confirm = new byte[1];
                int confirm;
                confirm = port2.ReadByte();
                Thread.Sleep(20);
                if (confirm == 2)
                {
                    Log("Tracking EQ-North confirmed");
                    TrackingOn = true;
                }
                else
                {
                    if (retry < 3)
                    {
                        retry++;
                        Thread.Sleep(1000);
                        ResumeTracking();
                    }
                    else
                    {
                        retry = 0;
                        MessageBox.Show("Unable to Resume Tracking", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                port2.Close();
            }
            catch(Exception ex) 
            {
                Log("Resume Tracking Error" + ex.ToString());
                Send("Resume Tracking Error" + ex.ToString());
                FileLog("Resume Tracking Error" + ex.ToString());
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                if (port2.IsOpen == false)
                    port2.Open();
                while (MountMoving == true)
                {
                    if (worker.CancellationPending == true)
                    {
                        e.Cancel = true;
                        port2.Close();
                        break;
                        //  return;
                    }
                    port2.DiscardInBuffer();
                    port2.Write("L");
                    Thread.Sleep(20);
                    port2.DiscardOutBuffer();

                    //  textBox13.Clear();
                    Thread.Sleep(20);

                    //  
                    GotoDoneCommand = port2.ReadExisting();
                    //  Thread.Sleep(10);
                    //  port2.DiscardOutBuffer();
                    port2.DiscardInBuffer();
                    // textBox13.Text = GotoDoneCommand.ToString();
                    Thread.Sleep(20);

                    if (GotoDoneCommand == "0#")
                    {
                        //  textBox13.Text = "Goto Done";

                        MountMoving = false;
                        port2.DiscardInBuffer();
                        port2.DiscardOutBuffer();
                        port2.Close();
                        if (TargetGotoOn == true)
                        {
                            //      button35.Text = "At Target Pos";
                            //      button35.BackColor = System.Drawing.Color.Lime;
                            toolStripStatusLabel1.Text = "At Target Position";
                         //   this.Refresh();
                            TargetGotoOn = false;
                        }
                        if (FocusGotoOn == true)
                        {
                            //   button33.Text = "At Focus Pos";
                            //   button33.BackColor = System.Drawing.Color.Lime;
                            toolStripStatusLabel1.Text = "At Focus Position";
                           // this.Refresh();

                            FocusGotoOn = false;
                            // Thread.Sleep(1000);
                            //port2.Close();

                        }
                        break;
                        // return;
                    }

                }
                return;
            }
            catch(Exception ex)
            {
                
                Send("BackgroundWorker1_DoWork Error" + ex.ToString());
                FileLog("BackgroundWorker1_DoWork Error" + ex.ToString());
                Log("BackgroundWorker1_DoWork Error" + ex.ToString());
            }
            }

private string user = WindowsFormsApplication1.Properties.Settings.Default.user;
private string server = WindowsFormsApplication1.Properties.Settings.Default.server;
private string to = WindowsFormsApplication1.Properties.Settings.Default.to;
private string pswd = WindowsFormsApplication1.Properties.Settings.Default.pswd;


public void Send(string msg)
{
    if (checkBox12.Checked == true)
    {
        this.user = WindowsFormsApplication1.Properties.Settings.Default.user;
        MailMessage Mail = new MailMessage();

        Mail.Subject = ("scopefocus - Error");

        Mail.Body = (msg);

        Mail.BodyEncoding = Encoding.GetEncoding("Windows-1254"); // Turkish Character Encoding

        Mail.From = new MailAddress(user);

        Mail.To.Add(new MailAddress(to));

        System.Net.Mail.SmtpClient Smtp = new SmtpClient();

        Smtp.Host = (server); // for example gmail smtp server

        Smtp.EnableSsl = true;
        textBox27.Text = user.ToString();
        // Smtp.Credentials = new System.Net.NetworkCredential("ksipp911@gmail.com", "cloe$1124");
        Smtp.Credentials = new System.Net.NetworkCredential(user, pswd);

        Smtp.Send(Mail);
        FileLog("message sent to " + to.ToString());
        Log("message sent to " + to.ToString());
    }
}

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            server = textBox13.Text.ToString();
        }

        private void textBox27_TextChanged(object sender, EventArgs e)
        {
            user = textBox27.Text.ToString();
        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {
            pswd = textBox28.Text.ToString();

        }

        private void textBox36_TextChanged(object sender, EventArgs e)
        {
            to = textBox36.Text.ToString();
        }

        private void button38_Click_1(object sender, EventArgs e)
        {
            Send("test message");
        }


        public void FileLog(string textlog)
        {
          //  string strLogText = "Std Dev" + "\t  " + abc[vProgress].ToString() + "\t  " + avg.ToString() + "\t" + (vProgress + 1).ToString() + "\t" + stdev.ToString();
            string path = textBox11.Text.ToString();
            string fullpath = path + @"\log.txt";
            StreamWriter log;
            if (!File.Exists(fullpath))
            {
                log = new StreamWriter(fullpath);
            }
            else
            {
                log = File.AppendText(fullpath);
            }              
            
            log.WriteLine(DateTime.Now);
            log.WriteLine("scopefocus - Error");
            log.WriteLine(textlog);
            log.Close();
        }
//*****all of these allow for ONLY numeric entry in these texboxes
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox18_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void LogTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox30_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox31_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox32_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void Bitstuff()
        {
            Bitmap bmp = new Bitmap("c:\\atest3\\Focus_3556_9685_1.76_1.04.bmp");
            int height = bmp.Size.Height;
            int width = bmp.Size.Width;
            //textBox37.Text = height.ToString();


        }
               
        private void button21_Click(object sender, EventArgs e)
        {
            ToggleFlat();
        }
               
        private void ToggleFlat()
        {
                toolStripStatusLabel1.Text = "Flat Panel Moving";
                this.Refresh();
                string movedone;
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
                Thread.Sleep(20);
                port.Write("R");
                filterMoving = true;
                Thread.Sleep(50);

                while (filterMoving == true)
                {


                    movedone = port.ReadExisting();
                    if (movedone == "D")
                    {

                        filterMoving = false;
                        this.Refresh();
                        toolStripStatusLabel1.Text = "Flat Move Complete";
                        this.Refresh();
                        //   toolStripStatusLabel1.Text = " ";
                        //    toolStripStatusLabel1.Text = " ";

                    }
                }
                Thread.Sleep(50);
                port.DiscardInBuffer();
                if (FlatsOn == false)
                {
                    FlatsOn = true;
                    button21.Text = "Flat On";
                    button21.BackColor = System.Drawing.Color.Lime;


                }

                else
                {
                    FlatsOn = false;
                    button21.Text = "Flat Off";
                    button21.BackColor = System.Drawing.Color.Red;
                }   
            
        }

        //Add PHD monitor for lost guidesstar
        //can use as built in soundalert finction AND compensate for meridian flip lost guide star
//added to resumePHd() 4-9-12.  This is just for testing purpose w/ button 39 o/w not used. 

//**********also consider this:
        //  - (1.13.5) New server command (id=14) to auto-find star (same as pulling down this from the menu)

        public double PHDtime = 0.0;
        public double DX = 0.0;
        public double DY = 0.0;
        private void PHDCheckSNR()
        {

            try
            {
                int StatusstripHandle = FindWindowEx(PHDhwnd, 0, "msctls_statusbar32", null);

                //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                IntPtr statusHandle = new IntPtr(StatusstripHandle);
                StatusHelper sh = new StatusHelper(statusHandle);
                string[] captions = sh.Captions;
                //   if (captions[0] == "Guiding")//******must rem for sim debugging *****
                if (captions[1] != null)
                    if (captions[1].Substring(0, 3)== "LOW")
                    {
                        MessageBox.Show("LOW SNR");
                    }
            }
            catch(Exception ex)
            {

                Log("PHD Low SNR error" + ex.ToString());
                Send("PHD Low SNR error" + ex.ToString());
                FileLog("PHD Low SNR error" + ex.ToString());
            }
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            posMin = Convert.ToInt16(textBox4.Text);//allows direct entry of close focus position
        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            CountFilterTotal();
            if (checkBox13.Checked == false)
            {
                numericUpDown32.Value = 0;
                numericUpDown34.Value = 0;
            }
        }

       

        private void comboBox7_Leave(object sender, EventArgs e)
        {
           
        }
        //scope item add
        private void button23_Click_1(object sender, EventArgs e)
        {
            string item = comboBox7.Text.ToString();

            comboBox7.Items.Add(item);
            WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Add(item);
        }
        //scope item remove
        private void button41_Click(object sender, EventArgs e)
        {
            string item = comboBox7.Text.ToString();

            comboBox7.Items.Remove(item);
            WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Remove(item);
        }
        //filter item add
        private void button43_Click(object sender, EventArgs e)
        {
            string item = comboBox8.Text.ToString();

            comboBox8.Items.Add(item);
            WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Add(item);
        }
        //filter combobox item remove
        private void button44_Click(object sender, EventArgs e)
        {
            string item = comboBox8.Text.ToString();

            comboBox8.Items.Remove(item);
            WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Remove(item);
        }
        //pause
        bool paused = false;     
        private void button12_Click(object sender, EventArgs e)
        {
            if (paused == false)
            {
                fileSystemWatcher4.EnableRaisingEvents = false;
                button12.BackColor = System.Drawing.Color.Lime;
                paused = true;
                button12.Text = "Resume";
            }
            else
            {
                fileSystemWatcher4.EnableRaisingEvents = true;
                button12.UseVisualStyleBackColor = true;
                paused = false;
                button12.Text = "Pause";
            }
        }
        //stop
        private void button19_Click_2(object sender, EventArgs e)
        {
             DialogResult result;
            result = MessageBox.Show("Abort the current sequence?", "scopefocus",
                            MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
            if (result == DialogResult.OK)
            {
                fileSystemWatcher4.EnableRaisingEvents = false;
                filterCountCurrent = 0;
                subCountCurrent = 0;
                standby();
            }
            if (result == DialogResult.Cancel)
                return;
        }
        //pause filter tab
        private void button22_Click_2(object sender, EventArgs e)
        {
            if (paused == false)
            {
                fileSystemWatcher4.EnableRaisingEvents = false;
                button22.BackColor = System.Drawing.Color.Lime;
                paused = true;
                button22.Text = "Resume";
            }
            else
            {
                fileSystemWatcher4.EnableRaisingEvents = true;
                button22.UseVisualStyleBackColor = true;
                paused = false;
                button22.Text = "Pause";
            }
        }
       //stop filter tab
        private void button30_Click_1(object sender, EventArgs e)
        {

            DialogResult result;
            result = MessageBox.Show("Abort the current sequence?", "scopefocus",
                            MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
            if (result == DialogResult.OK)
            {
                fileSystemWatcher4.EnableRaisingEvents = false;
                filterCountCurrent = 0;
                subCountCurrent = 0;
                standby();
            }
            if (result == DialogResult.Cancel)
                return;
        }
        //set max travel according to selected scope.....allows for 6 diff scopes
        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            int i = comboBox7.SelectedIndex;
            if (i == 0)
                travel = WindowsFormsApplication1.Properties.Settings.Default.MaxTravel1;
            if (i == 1)
                travel = WindowsFormsApplication1.Properties.Settings.Default.MaxTravel2;
            if (i == 2)
                travel = WindowsFormsApplication1.Properties.Settings.Default.MaxTravel3;
            if (i == 3)
                travel = WindowsFormsApplication1.Properties.Settings.Default.MaxTravel4;
            if (i == 4)
                travel = WindowsFormsApplication1.Properties.Settings.Default.MaxTravel5;
            if (i == 5)
                travel = WindowsFormsApplication1.Properties.Settings.Default.MaxTravel6;
            textBox2.Text = travel.ToString();



        }

        int checkboxs;
        private void TotalSubs()
        {
            checkboxs = 0;
            if (checkBox1.Checked == true)
                checkboxs++;
            if (checkBox2.Checked == true)
                checkboxs++;
            if (checkBox3.Checked == true)
                checkboxs++;
            if (checkBox4.Checked == true)
                checkboxs++;
            if (checkBox18.Checked == true)
            totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + ((int)numericUpDown32.Value * checkboxs);
            else
                totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + (int)numericUpDown32.Value;
            textBox23.Text = totalsubs.ToString();
        }
       
        private void numericUpDown32_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDown32.Value > 0)
                checkBox13.Checked = true;
            TotalSubs();
           // checkBox13.Checked = true;
        }

       
        private void numericUpDown29_ValueChanged(object sender, EventArgs e)
        {
            TotalSubs();
               
        }
        
        private void button39_Click_1(object sender, EventArgs e)
        {
            
          

        }

        private void button40_Click(object sender, EventArgs e)
        {
           
          }
        private double[] FocusGroup = new double[4];
        private int[] SubsPerFocus = new int[4];

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if ((checkBox8.Checked == true) & (checkBox17.Checked == true)) 
                MessageBox.Show("Confirm refocus after filter change AND " + numericUpDown38.Value.ToString() + " subs?"
                    , "scopefocus");
        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {
            if ((checkBox8.Checked == true) & (checkBox17.Checked == true))
                MessageBox.Show("Confirm refocus after filter change AND " + numericUpDown38.Value.ToString() + " subs?"
                    , "scopefocus");

        }

        private void numericUpDown38_Leave(object sender, EventArgs e)
        {
            if (numericUpDown38.Value == 0)
                MessageBox.Show("Value cannot be zero", "scopefoucs");
            if (numericUpDown38.Value < 4)
                MessageBox.Show("Confrim low refocus per sub value of " + numericUpDown38.Value.ToString(), "scopefocus");
            FocusPerSub = (int)numericUpDown38.Value;
            if (checkBox1.Checked == true)
            {
                if (numericUpDown12.Value == 0)
                {
                    MessageBox.Show("Position 1 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                FocusGroup[0] = (double)numericUpDown12.Value / FocusPerSub;
                if (FocusGroup[0] != Math.Round(FocusGroup[0]))
                {
                    MessageBox.Show("Total pos. 1 subs must be multiple of Focus per sub", "scopefocus");
                    numericUpDown38.Value = 1;
                }
                else
                    SubsPerFocus[0] = (int)numericUpDown12.Value / (int)FocusGroup[0];
                if (FocusGroup[0] == 1)//then just use the filter change to refocus
                {
                    SubsPerFocus[0] = 0;
                    FocusGroup[0] = 0;
                }
            }
            if (checkBox2.Checked == true)
            {
                if (numericUpDown13.Value == 0)
                {
                    MessageBox.Show("Position 2 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                    FocusGroup[1] = (double)numericUpDown13.Value / FocusPerSub;
                if (FocusGroup[1] != Math.Round(FocusGroup[1]))
                {
                    MessageBox.Show("Total pos. 2 subs must be multiple of Focus per sub", "scopefocus");
                    numericUpDown38.Value = 1;
                }
                else
                    SubsPerFocus[1] = (int)numericUpDown13.Value / (int)FocusGroup[1];
                if (FocusGroup[1] == 1)//then just use the filter change to refocus
                {
                    SubsPerFocus[1] = 0;
                    FocusGroup[1] = 0;
                }
            }
            if (checkBox3.Checked == true)
            {
                if (numericUpDown14.Value == 0)
                {
                    MessageBox.Show("Position 3 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                FocusGroup[2] = (double)numericUpDown14.Value / FocusPerSub;
                if (FocusGroup[2] != Math.Round(FocusGroup[2]))
                {
                    MessageBox.Show("Total pos.3 subs must be multiple of Focus per sub", "scopefocus");
                    numericUpDown38.Value = 1;
                }
                else
                    SubsPerFocus[2] = (int)numericUpDown14.Value / (int)FocusGroup[2];
                if (FocusGroup[2] == 1)//then just use the filter change to refocus
                {
                    SubsPerFocus[2] = 0;
                    FocusGroup[2] = 0;
                }
            }
            if (checkBox4.Checked == true)
            {
                if (numericUpDown15.Value == 0)
                {
                    MessageBox.Show("Position 4 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                FocusGroup[3] = (double)numericUpDown15.Value / FocusPerSub;
                if (FocusGroup[3] != Math.Round(FocusGroup[3]))
                {
                    MessageBox.Show("Total pos. 4 subs must be multiple of Focus per sub", "scopefocus");
                    numericUpDown38.Value = 1;
                }
                else
                    SubsPerFocus[3] = (int)numericUpDown15.Value / (int)FocusGroup[3];
                if (FocusGroup[3] == 1)//then just use the filter change to refocus
                {
                    SubsPerFocus[3] = 0;
                    FocusGroup[3] = 0;
                }
            }
            FocusPerSubGroupCount = (int)(FocusGroup[0] + FocusGroup[1] + FocusGroup[2] + FocusGroup[3]);      
       
        }
        //  doflathere
        private void DoFlat()
        {
            if ((IsSlave()) & (numericUpDown37.Value ==0))
                return;
            if (IsServer())
            {
                SlaveFlat();
            }
            if (!IsSlave())
            toolStripStatusLabel1.Text = "Saving Flats";
            this.Refresh();
            if (!IsSlave())
            {
            if (checkBox19.Checked == true)
                StopPHD();
            
                if (FlatsOn == false)
                    ToggleFlat();
            }
            
            Thread.Sleep(1000);
            editfound = 0;
            Nebname = Nebname + "_Flat";
            if (IsSlave())
            {
                subsperfilter = (int)numericUpDown37.Value;
                CaptureTime = Convert.ToInt16(textBox40.Text);//starting exposure time
                CaptureBin = (int)numericUpDown35.Value;
            }
            else
            {
                subsperfilter = (int)numericUpDown32.Value;
                CaptureTime = (int)numericUpDown34.Value;//starting exposure time
                if (checkBox18.Checked == false)
                    CaptureBin = (int)numericUpDown36.Value;
            }
            
            FilterFocusOn = false;
              fileSystemWatcher4.EnableRaisingEvents = false;
              if (IsSlave() && checkBox21.Checked == true)
                  CalculateFlatExp();
              if ((IsSlave()) & (checkBox21.Checked == false))
                  textBox41.Text = "Slave FlatCalc Complete";

            if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                CalculateFlatExp();
            
            NebCapture();
            if ((IsServer()) & (SlaveFlatOn == true))
            {
                SlaveFlatOn = false;
                if (IsServer())
                {
                    Log("Waiting for slave flat ");
                    toolStripStatusLabel1.Text = "Waiting for slave flat";
                    this.Refresh();
                    while (working == true)
                    {
                        WaitForSequenceDone("Sequence done", NebSlavehwnd);
                       // Thread.Sleep(50);
                    }
                    working = true;
                    Log("Slave flat done");
                   
                }

            }
            if (!IsSlave())
            ToggleFlat();
            /*
            if ((IsServer()) & (SlaveFlatOn == true))
            {
                SlaveFlatOn = false;
                
            //    SlavePause();
              //  return;
            }
             */
            if (checkBox19.Checked == true)
                ResumePHD();
            if (!IsSlave())
            {
                subCountCurrent = subCountCurrent + (int)numericUpDown32.Value;
                fileSystemWatcher4.EnableRaisingEvents = true;

                checkfiltercount();// added 5-23 this is needed to end sequence if no darks selected
            } 
            }


        private void button29_Click(object sender, EventArgs e)
        {
            if (textBox34.Text == "")
            {
                DialogResult result = openFileDialog1.ShowDialog();
                ImportPath = openFileDialog1.FileName.ToString();
                textBox34.Text = ImportPath.ToString();
            }
                
            importDataFromExcel();
            button17.PerformClick();//update
        }
        //export
        private void button42_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (ExcelFilename == null)
                {
                    DialogResult result = saveFileDialog1.ShowDialog();
                    ExcelFilename = saveFileDialog1.FileName.ToString();
                    textBox33.Text = ExcelFilename.ToString();
                }

                using (SqlCeConnection con = new SqlCeConnection(conString))
                {
                    con.Open();
                    using (SqlCeDataAdapter a = new SqlCeDataAdapter("SELECT * FROM table1", con))
                    {
                        DataTable t = new DataTable();
                        a.Fill(t);
                        // dataGridView1.DataSource = t;
                        a.Update(t);
                        Excel_FromDataTable(t);

                    }
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                Log("SQL Excel Export error 5800" + ex.ToString());
            }
        }
        //if serial connection made after scopefocus start up, click combox to populate list
        private void comboBox1_Click(object sender, EventArgs e)
        {
            string[] portlist = SerialPort.GetPortNames();
            foreach (String s in portlist)
            {
/*
                if (comboBox1.Items.Count == 0)
                    comboBox1.Items.Add(s);

                for (int x = 0; x < comboBox1.Items.Count; x++)
                {
                    if (comboBox1.Items[x] != s)
                        comboBox1.Items.Add(s);
                }
                for (int x = 0; x < comboBox6.Items.Count; x++)
                {
                    if (comboBox6.Items[x] != s)
                        comboBox6.Items.Add(s);
                }
              */ 
            }
        }
        //disable bin selector for flat after every filter
        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox18.Checked == true)
            {
                numericUpDown36.Enabled = false;
                TotalSubs();//recalculate totalsubs
            }
            
               
        }
        //for verion number on about tab
        public string PublishVersion
        {
            get
            {
                if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
                {
                    Version ver = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
                    return string.Format("{0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision);
                }
                else
                    return "Not Published";
            }
        }
        //Dual mode capture button
        int TotalTime;
        
        private void button45_Click(object sender, EventArgs e)
        {
            Log("SlaveCapture Hit");
            /*
            if (checkBox20.Checked == true)
            {
                SetForegroundWindow(Slavehwnd);
                PostMessage(Capturehwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(1000);
            }
             */
            if (checkBox14.Checked == true)
            {
                
              //  NebListenStart(NebhWnd, SocketPort);
              //  Thread.Sleep(4000);
             // textBox41.Refresh();
                SlavePaused = false;

             //   CaptureTime = Convert.ToInt16(textBox37.Text) * 1000;
             //   subsperfilter = ((Convert.ToInt16(textBox41.Text)/CaptureTime));//calculate number of subs based on server total time
               // subsperfilter = (int)numericUpDown23.Value;
                Log("Received total time " + subsperfilter.ToString());
                CaptureBin = (int)numericUpDown33.Value;
                //textBox41.Clear();
                NebCapture();
            }
         //   fileSystemWatcher4.EnableRaisingEvents = true;
        }
  
        //pause (Neb Abort) 
        private void button46_Click(object sender, EventArgs e)
        {
            Log("Slave Pause hit");
            if (checkBox20.Checked == true)
            {
                SlavePause();
            }
            if (checkBox14.Checked == true)
            {
                ShowWindowAsync(NebhWnd, SW_SHOW);
                Thread.Sleep(200);
                SetForegroundWindow(Aborthwnd);
                PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);
                
                PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);
                //********added 6-6
             //   serverStream.Dispose();
             //   serverStream.Close();
              //  clientSocket2.GetStream().Dispose();
              //  clientSocket2.GetStream().Close();


                
                clientSocket2.Close();
                Thread.Sleep(250);
                SendKeys.SendWait("~");
                SendKeys.Flush();
                //*********
                
                NebListenOn = false;
                if (clientSocket2.Connected == false)
                    Log("Slave Disconnected");
                else
                    Log("slave still connected");
               
            }
            
        }

        
        int ServerStatusHandle;
        int NebhWnd2;
        int Slavehwnd;
       int SocketPort2;
        //enable server
        private void checkBox20_CheckedChanged_1(object sender, EventArgs e)
        {
          //  groupBox18.Enabled = false;
            button45.Text = "S_Capture";
            button46.Text = "S_Pause";
            button47.Text = "S_Flat";
            button39.Text = "S_GotoFocus";
            numericUpDown37.Enabled = false;
           // numericUpDown23.Enabled = false;
            textBox37.Enabled = false;
            textBox40.Enabled = false;
            numericUpDown33.Enabled = false;
            numericUpDown35.Enabled = false;

            /*
            if (checkBox14.Checked == true)
            {
                checkBox14.Checked = false;//need to uncheck slave
                //all this needed only if changing FROM slave mode
                MainWindowName = "Nebulosity";//change the mainwindow label in neb to signify sencond copy
                SocketPort = 4301;
                ScriptName = "\\listenPort.neb";
                this.Text = "scopefocus";
                groupBox15.Enabled = true;
                groupBox13.Enabled = true;
                groupBox12.Enabled = true;
                groupBox6.Enabled = true;
                groupBox5.Enabled = true;
                checkBox8.Enabled = true;
                Callback myCallBack = new Callback(EnumChildGetValue);
                FindHandles();
                //   int hWnd;

                if (NebhWnd == 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Nebulosity Not Found - Open or Close & Restart and hit 'Retry' or 'Ignore' to continue",
                       "scopefocus", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);//change so ok moves focus
                    if (result == DialogResult.Ignore)
                        this.Refresh();
                    if (result == DialogResult.Retry)
                    {
                        FindHandles();
                        this.Refresh();
                    }
                    if (result == DialogResult.Abort)
                        System.Environment.Exit(0);

                }
                else
                {

                    EnumChildWindows(NebhWnd, myCallBack, 0);
                }

                FindNebCamera();

                this.Refresh();
            }
             */


          //  SearchData sd = new SearchData { Wndclass = "WindowsForms10", Title = "scopefocus" };
            IntPtr SlavehwndPtr = SearchForWindow("WindowsForms10", "scopefocus - slave mode");
            Log("scopefocus-slave handle found --  " + SlavehwndPtr.ToInt32());
            Slavehwnd = SlavehwndPtr.ToInt32();
            //full class name   WindowsForms10.Window.8.app.0.201d787_r15_ad1
            if (Slavehwnd != 0)
            {
                Callback myCallBack2 = new Callback(EnumChildGetValue);
                EnumChildWindows(Slavehwnd, myCallBack2, 0);//this gets the slave capture, pause, flat and gotofocus buttons
                //this sends the server status box handle to the slave textbox
                string send = textBox41.Handle.ToString();
                StringBuilder sb = new StringBuilder(send);
                SendMessage(SlaveStatushwnd, WM_SETTEXT, 0, sb);
            }
            
            
             
            //make sure it's not the spawned copy  
             
            MainWindowName = "Nebulosity v3.0";//change the mainwindow label in neb to signify sencond copy
            IntPtr hWnd2 = SearchForWindow("wxWindow", MainWindowName);//must open file named test.fit to specifiy second copu
            Log("Ensure not slave handle -- " + hWnd2.ToInt32());
            NebhWnd = hWnd2.ToInt32();
            /*
            Callback myCallBack3 = new Callback(EnumChildGetValue);
            EnumChildWindows(NebhWnd2, myCallBack3, 2);
            Log("Slave Handles-- Abort2: " + Aborthwnd2.ToString() + "Frame" + Framehwnd2.ToString());
            

            SocketPort2 = 4302;
                ScriptName = "\\listenPort2.neb";
               
                Callback myCallBack = new Callback(EnumChildGetValue);
              //  FindHandles();
                //   int hWnd;

                if (NebhWnd2 == 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Nebulosity-slave  Not Found - Open or Close & Restart and hit 'Retry' or 'Ignore' to continue",
                       "scopefocus", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);//change so ok moves focus
                    if (result == DialogResult.Ignore)
                        this.Refresh();
                    if (result == DialogResult.Retry)
                    {
                        IntPtr hWnd3 = SearchForWindow("wxWindow", MainWindowName);//must open file named test.fit to specifiy second copu
                        Log("Neb-slave Handle Found -- " + hWnd3.ToInt32());
                        NebhWnd2 = hWnd2.ToInt32();
                        this.Refresh();
                    }
                    if (result == DialogResult.Abort)
                        System.Environment.Exit(0);

                }
                    
                else
                {

                    EnumChildWindows(NebhWnd, myCallBack, 0);
                }
                     
              //  FindNebCamera();

                this.Refresh();
            */
            
                     
        }
        public int Serverhwnd;
        //enable slave mode
        private void checkBox14_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox14.Checked == true)
            {
                if (textBox37.Text == "")
                {
                    MessageBox.Show("Capture time cannot be zero", "scopefocus");
                    checkBox14.Checked = false;
                    return;
                }
               // MainWindowName = "Nebulosity v3.0 - slave.fit";//change the mainwindow label in neb to signify sencond copy
                MainWindowName = "Nebulosity(spawned) v3.0";//change the mainwindow label in neb to signify sencond copy
                SocketPort = 4302; // was SocketPort2????
                ScriptName = "\\listenPort2.neb";
                this.Text = "scopefocus - slave mode";
                groupBox15.Enabled = false;
                groupBox13.Enabled = false;
                groupBox12.Enabled = false;
                groupBox6.Enabled = false;
                groupBox5.Enabled = false;
                checkBox8.Enabled = false;
                button40.Enabled = false;
                Callback myCallBack3 = new Callback(EnumChildGetValue);
               
                //   int hWnd;
                IntPtr hWnd2 = SearchForWindow("wxWindow", MainWindowName);//must open file named test.fit to specifiy second copu
                Log("Neb(spawned) handle -- " + hWnd2.ToInt32());
                NebhWnd = hWnd2.ToInt32();
               // FindHandles();
                if (NebhWnd == 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Nebulosity slave Not Found - Open or Close & Restart and hit 'Retry' or 'Ignore' to continue",
                       "scopefocus", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);//change so ok moves focus
                    if (result == DialogResult.Ignore)
                        this.Refresh();
                    if (result == DialogResult.Retry)
                    {
                        FindHandles();
                        this.Refresh();
                    }
                    if (result == DialogResult.Abort)
                        System.Environment.Exit(0);

                }
                else
                {

                    EnumChildWindows(NebhWnd, myCallBack3, 0);
                }
                FindNebCamera();
                IntPtr ServerhwndPtr = SearchForWindow("WindowsForms10", "scopefocus - Main");
                Log("scopefocus-server handle found --  " + ServerhwndPtr.ToInt32());
                Serverhwnd = ServerhwndPtr.ToInt32();
              //  this.Refresh();

            }
            else
            {
                SocketPort = 4301;
                ScriptName = "\\listenPort.neb";
            }

        }
        bool StatusHandleRec = false;
        bool StatusHandleSend = false;
        private void textBox41_TextChanged(object sender, EventArgs e)
        {
            if ((checkBox14.Checked == true) && (StatusHandleSend == false))
            {
                ServerStatusHandle = Convert.ToInt32(textBox41.Text);
                Log("Server Status Handle: " + ServerStatusHandle.ToString());
                string send = NebhWnd.ToString();
                StringBuilder sb = new StringBuilder(send);
                //  sb = send;
                SendMessage(ServerStatusHandle, WM_SETTEXT, 0, sb);
                StatusHandleSend = true;
            }
            if ((IsServer()) & (StatusHandleRec == false))
            {
                NebSlavehwnd =  Convert.ToInt32(textBox41.Text);
                StatusHandleRec = true;
            }
            /* //remd***** 6_18
            if (IsServer())
            {
                if ((textBox41.Text == "Sub Complete") || (textBox41.Text == "Focus Done"))
                    timer2.Start();
            }
            if (IsSlave())
            {
                if (textBox41.Text == "Capturing")
                    timer2.Start();
            }
          */

        }
        public bool IsSlave()
        {
            if (checkBox14.Checked)
                return true;
            else
                return false;
        }

        public bool IsServer()
        {
            if (checkBox20.Checked)
                return true;
            else
                return false;
        }
        private bool SlavePaused = false;

        private void SlavePause()
        {
            //if (IsSlave())
           // {
                /*
                //     SlavePaused = true;
                Listen0();
                Thread.Sleep(1000);
                ShowWindowAsync(Slavehwnd, SW_SHOW);
                SetForegroundWindow(Pausehwnd);
                PostMessage(Pausehwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);
                SendKeys.SendWait("^r");
                serverStream.Close();
                SetForegroundWindow(NebhWnd);
                Thread.Sleep(1000);
                PostMessage(Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(1000);
                NebListenOn = false;
                // clientSocket.GetStream().Close();//added 5-17-12
                //  clientSocket.Client.Disconnect(true);//added 5-17-12
                clientSocket.Close();
                 }
                */
                
                SetForegroundWindow(Pausehwnd);
                PostMessage(Pausehwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);
              
                /*
                SetForegroundWindow(Pausehwnd);
                PostMessage(Pausehwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);
                */
            
        }

        private void SendtoSlave(string message)
        {
            StringBuilder sb = new StringBuilder(message);
            ShowWindowAsync(Slavehwnd, SW_SHOW);
            Thread.Sleep(200);
            SetForegroundWindow(Slavehwnd);
            SendMessage(SlaveStatushwnd, WM_SETTEXT, 0, sb);
            Thread.Sleep(500);
        }
        private void SendtoServer(string message)
        {
            StringBuilder sb = new StringBuilder(message);
            ShowWindowAsync(Serverhwnd, SW_SHOW);
            Thread.Sleep(200);
            SetForegroundWindow(Serverhwnd);
            SendMessage(ServerStatusHandle, WM_SETTEXT, 0, sb);
        }
        private void SlaveCapture()
        {
           // SendtoSlave((subsperfilter * CaptureTime3).ToString());
            SlavePaused = false;
            ShowWindowAsync(Slavehwnd, SW_SHOW);
            ShowWindowAsync(Slavehwnd, SW_RESTORE);
            SetForegroundWindow(Slavehwnd);
            PostMessage(Capturehwnd, BN_CLICKED, 0, 0);
            Thread.Sleep(500);
        }

        private void button40_Click_1(object sender, EventArgs e)
        {
           SequenceGo();
        }



        private void button48_Click(object sender, EventArgs e)
        {
            Listen0();
            
            
        }
        /*
        //prob not needed, unless can figure out how to monitor textbox and allow slave sub completion
        private void timer2_Tick(object sender, EventArgs e)//holds text in box for a while
        {

            if (textBox41.Text == "Sub Complete")
            {
                textBox41.Text = "Sub Complete";
                timer2.Stop();
                if (timer2.Enabled == false)
                    textBox41.Clear();

            }
            if (textBox41.Text == "Focus Done")
            {
                textBox41.Text = "Focus Done";
                timer2.Stop();
                if (timer2.Enabled == false)
                    textBox41.Clear();
            }
            if (textBox41.Text == "Capturing")
            {
                textBox41.Text = "Capturing";
                timer2.Stop();
                if (timer2.Enabled == false)
                    textBox41.Clear();
            }
        }
         */
        private void SlaveFlat()
        {
            SlaveFlatOn = true;
            SetForegroundWindow(Slavehwnd);
            PostMessage(Flathwnd, BN_CLICKED, 0, 0);
            Thread.Sleep(250);
        }
        private bool SlaveFlatOn = false;
        private void button47_Click(object sender, EventArgs e)
        {
            Log("SlaveFlat hit");
            if (IsSlave())
            {
                SlavePaused = false;
                SlaveFlatOn = true;
                DoFlat(); 
            }
            if (IsServer())
            {
                SlaveFlatOn = true;
                SetForegroundWindow(Slavehwnd);
                PostMessage(Flathwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);
            }


        }
        string[] captions;
        bool working = true;
        int Workerhwnd;
        string StatusLabel;
      
        private void SlaveFocus()
        {
            FilterFocusOn = true;//??unsure if needed was slaveflat on=true (
            ShowWindow(Gotofocushwnd, SW_SHOW);
            ShowWindow(Gotofocushwnd, SW_RESTORE);
            SetForegroundWindow(Slavehwnd);
            PostMessage(Gotofocushwnd, BN_CLICKED, 0, 0);
            Thread.Sleep(250);
        }
        private void button39_Click(object sender, EventArgs e)
        {
            if (IsSlave())
            {
                Thread.Sleep(3000);
                FilterFocus();
            }

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            path2 = textBox11.Text.ToString();
            Log("Path set to " + path2.ToString());
            fileSystemWatcher1.Path = path2.ToString();
            fileSystemWatcher2.Path = path2.ToString();
            fileSystemWatcher5.Path = path2.ToString();//added to test metricHFR
            fileSystemWatcher3.Path = path2.ToString();
            fileSystemWatcher4.Path = path2.ToString();
        }

       
        private void textBox42_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This value is calcuated from server subs * capture time", "scopefocus");
        }

        private void checkBox22_CheckedChanged_1(object sender, EventArgs e)
        {
            EquipRefresh();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            SampleMetric();
        }


       

        
        

//****************future*************************
// work on "center here" function.  figure out how to translate neb cursor position to mount movement. 
//use conversion factro much like dx/dy.  need to determine movement vectors then send as mount slew





//add above here
    }
   

}