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

// 9-5-13...rem'd a bunch of          //  scope = new ASCOM.DriverAccess.Telescope(devId);   may need to be unrem'd. 

//9-30-13  *****************NEED to FIX max travel.  see line 10858*******************
// 11-6-13 (search for this date) work around for reverse box check with FSQ-85.  need perm fix!
//7-14-14  gotopos, removed all attempts at backlash compensation.  was using a method that is not std for ASCOM.  
//7-116-14  ver12   removed error handeling for > travel or < 0 for rev/fwd buttons, redudnent handled in gotopos


// change log - Ver 19 - published
// ver 91 7-28-14 v curve lost star lines 1815 to 1846  may need to adjust parameter ratio > 1.5 and ratio < .5
//works w/ sim
// extra exposure at end of vcurve, changed vcurve to set vdone = 1 at filesystemwatcher2_ method and set fsw2.enablewatchingevents to false there
//there is stull redundant code in vcurve() that should be removed some time. 

//ver 21 change format of import/export.  click blank textbox brings up browser.  click open/save on browser completes the import/export
//       if fileanme already in textbox, will perform on that filenema without using browser.  can edit filename without browser
//    ***********need to error handle validity of filename**************

//ver 22 fixed problem with blank messaging parameters
//ver 23:  TB1 not displaying corrrect position after GotoFocus.  changed several focuser.Move to gotopos lines 2620,2624,2633,2646,2658,3254,3257 
//ver 24  added prelim Internal filter support.  added checkBox31.  added if (!checkBox31.check) in front of all Filter.filterWheel.position.  
//         ASUMES starting at lowest Filter.filterWheel position.  NebCapture sends SetExtFilter N where N = currentfilter
//         thus far untested w/ flats and darks in sequence.  

//ver 25 corected not saving backlash value, allowed leaving textbox to save as well as hitting enter. 
//  manually changing focuser position does not work for sim...may only work for my focuser.  may need to remove if not ASCOM std.  
//  11-14-14 not committed/published yet.  
//  all findNebCamera() rem'd
//  All NoCameraSelected() rem'd 
//  Globalvariable.NebcCamera rem'd 
//  PHDCheckSNR() - rem'd, not used 
// in vcurve() turned FSW enablerasing events to false while focuser moving then back to true once done.  

//11-22-14 first try at SVN
//1-5-15 v 27?changed to sql server compact 4
// Multiple version trying to resolve SqlServerCe install issues. 
//  ver38 resolved w/ updateed SqlServerCe and ErikEJ.SqlCe40 updates. 
//ver 39 abort gotofocus if data is null, changed position textbox to read only, comment out the setPosition stuff since unlikely implemented in most ascom focusers
//ver 49 added PHD lost star recovery in loststarmonitor().  Line 13517  not released as of 4-15-15
//test subversion from laptop...
//timer2 tick added try/catch 4-17-15
//4-19-15 fixed monir bugs w/ plate solve. 
//****** for plate solve user must create c:\cygwin\home\atro directory AND for center here add tablist.exe to  cygwin\lib\astrometry *****

//8-11-15   ver52 corrected slope line 2207-2214.  dwn was pos and should be negative and vice versa. 

//9-9-15 added check to make sure not gros focus point mis calculation line 964 and 997
// if focus point is beyond sample point on wrong side of slope or 4 times sample point on right side repeat x 3 then abort and send text
//     
//9-9-15 added test for improvement for both single star and metric.  start line 1466 
//   takes a few single star or one more metric and checks for improvement, if no improvement chages slope side and repeats
//     if single star both sides fails...will try full frame metric 

//10-16-15
// fixed auto mouse click not functioing, different "panel" window for Neb 4
// added more delay to send file name to Load Script window
//disabled goto focus quality check for simulaotr

//10-19-15  
// double click camera textbox to rescan (only matters for simulator and disable qaulity checks)

//3-2-16
//added out of bound focus move management got gotofocus and allow 3 retry then abort and send msg
//added while statement to prevent 'in sequence' re-focus from advancing prior to completeing focuser move
//added eneble numericupdown after pushing stop

//3-12-16 added online astrometry.net plate solve
//3-13-16 added simple method to view .fits tables using tablist,  to use add a button and runtablistViewer(filename.xxx) then view output text2.txxt in cygwin/home/astro

// 10-10-16 begin rework of full fram metric.  allow selection of v-curve data for specific traget FOV.  
// do one for eash session/target.  Versus every gotofocus (like SGP).  
//changes dataGridView1 selection ode to "full row select" from "cell select"

// 10-11-16 many changes for new FF metric works w/ CDCSimulator
// 11-8-16  added clipboard use, (works better). removed log timer tick 2 error.  
//   -- add  a bunch of "&& (Filter.DevId3 != null))" for dark and flat filter changes that might be done witout filter wheel
// -- capture status reads from Nebs statuswindow
// -- added globalvariable.Capcurrent and CapTotal 

// 11-9-16 mult fixes for focus and target location use image plate solve

// 11-13-16  added Clipboard confirmation for almost all //NEB commands  Search for !NebCommandConfirm(  if need to correct
// 11-14-16 ver 32  added clear clipboard after all final //NEB commands
 // TODO revisit confrim commands sent with statusbar monitor.  Maybe just check it once after each clipboard.setdata instead of using while loop.....

 // 11-18-16 changed focusgroupset() to only happen at sequencego...not w/ numUD or checkbox change
 // 11-21-16 added resizenebwindow() , fixes status monitor error when small window results in "..." in Nebs status bar.

// TODO: -- pause/resume  line 12820......
//       -- add ability to add comment to v-curve data, maybe in equip cell  


// *******for v-curve simulator.   Use known v-curve data, make sure step size and focus point are set the same/  
//                                   open SimGen, select a focus .bmp file.  start the fine v-curve.  enter the desired hfr (with decimal eg 3.45) hit save.  a new point will be generated
//                                     make sure to unrem lines in fileSystemWatcher2_Changed.  
//                                     make sure to not use the file name originally opened. (change by .01 when get to that one) 
//                                     rem filesystemwatcher2_delete lines 4254 down   // not sure if necessary.  


///  to do:
/// ver21 see above 
/// **** ??  able to abort focus move if needed
/// **** error handeling for plate solve. 
/// 
/// ****fix so when using metric the focus star in target frame (checkBox 10) doesn't matter.  get target position not set error if unchecked
/// ***check MonitorNebStatus()....may have problems is neb is closed before scopefocus
/// 
///  for sequence running add script control of internal ?(or external) Filter.filterWheel e.g. for QSI see pg 95 of manual
///  SetFilterN  or SetExtFilterN (1-indexed)



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
//using System.IO.Ports;
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
using ASCOM.DriverAccess;
using System.Text.RegularExpressions;
using System.Deployment.Application;
using System.Runtime.CompilerServices;
using nom.tam.fits;








namespace Pololu.Usc.ScopeFocus
{

    public partial class MainWindow : Form
    {

        TcpClient clientSocket2 = new TcpClient();//added for neb slave
        TcpClient clientSocket = new TcpClient();//added for neb communication
        TcpClient phdsocket = new TcpClient();//test for phd socket
        NetworkStream serverStream;
        AstrometryNet ast = new AstrometryNet();
        //  AstrometryNet ast = AstrometryNet.OnlyInstance;
        //  AstrometryNet ast;


        //AstrometryNet ast;
        //   SerialPort port;
        //   SerialPort port2;

        Focuser focuser;
      //  FilterWheel filterWheel;
        private ASCOM.DriverAccess.Switch FlatFlap;
        private ASCOM.DriverAccess.Camera cam;


        // 12-2-16 test PHD2 event monitoring
        PHD2comm ph = new PHD2comm();
        


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
            /*  
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
              */
              
            fileSystemWatcher1.EnableRaisingEvents = false;
            fileSystemWatcher2.EnableRaisingEvents = false;
            fileSystemWatcher3.EnableRaisingEvents = false;
            fileSystemWatcher4.EnableRaisingEvents = false;
            fileSystemWatcher5.EnableRaisingEvents = false;//added to test metricHFR
                                                           //       fileSystemWatcher6.EnableRaisingEvents = true;
       //     textBox11.Focus();


            //    string devId2 = ASCOM.DriverAccess.Focuser.Choose("");
        }
        //*******************remove this if the second form FindCenter is not used ***********
        //try adding to make this avail in antoher form;
        public class Variables
        {
            //private static int NebhWnd;

            //public static int NebhWndl
            //{
            //    get { return NebhWnd; }
            //    set { NebhWnd = value; }

            //}

        }

        //  Focuser focuser;
        // ASCOM.DriverAccess.Focuser focuser = new ASCOM.DriverAccess.Focuser("");
        //    Focuser focuser;
        Telescope scope;
        //  ASCOM.DriverAccess.Focuser focuser = new ASCOM.DriverAccess.Focuser("ASCOM.Arduino.Focuser");
        //Focuser focuser;

        //  *** start trying to make Vcurve separate class

        //  private string _combobox2;
        //public string Combobox2
        //{
        //    get { return comboBox2.Text; }
        //    set { comboBox2.Text = value; }
        //}
        //public string Combobox3
        //{
        //    get { return comboBox3.Text; }
        //    set { comboBox3.Text = value; }
        //}
        //public string Combobox4
        //{
        //    get { return comboBox4.Text; }
        //    set { comboBox4.Text = value; }
        //}
        //public string Combobox5
        //{
        //    get { return comboBox5.Text; }
        //    set { comboBox5.Text = value; }
        //}

        private int SocketPort = 4301;
        private string ScriptName = "\\listenPort.neb";
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


        // 8-14-13 start removing these variables
        // string NebVersion;
        //   private static int NebVNumber;
        //   string NebPath;
        //  string ImportPath;

        private static string ExcelFilename
        { get; set; }
        //   public string ImportPath
        //    { get; set;}
        //en
        //   int hWnd;  goes w/ enumerator at botton
        //  bool setupWindowFound = false;
        // int Advancedhwnd;
        //   int NebSlavehwnd;
        //    int Advhwnd;

        private bool FlatsOn = false;

        //End 8-14-13 edits
        private static int posMin;
        //public static int posMin
        //{
        //    get { return PosMin; }
        //    set { PosMin = value; }
        //}

        //8-15-13 edits
        private static int _apexHFR;
        private static int _enteredPID;
        private static double _enteredSlopeDWN;
        private static double _enteredSlopeUP;
        private static string _equip;
        private static int _pID;
        private static string _importPath;
        private static int _intersectPos;

        private static int _hFRarraymin = 999;
        private static int _posminHFR;
        private static double _slopeHFRup = 0;
        private static double _slopeHFRdwn = 0;
        private int rows;
        public static int intersectPos
        {
            get { return _intersectPos; }
            set { _intersectPos = value; }
        }


        public static string ImportPath
        {
            get { return _importPath; }
            set { _importPath = value; }
        }
        public static int PID
        {
            get { return _pID; }
            set { _pID = value; }
        }
        public static int PosminHFR
        {
            get { return _posminHFR; }
            set { _posminHFR = value; }
        }
        public static double SlopeHFRup
        {
            get { return _slopeHFRup; }
            set { _slopeHFRup = value; }
        }
        public static double SlopeHFRdwn
        {
            get { return _slopeHFRdwn; }
            set { _slopeHFRdwn = value; }
        }
        public static int HFRarraymin
        {
            get { return _hFRarraymin; }
            set { _hFRarraymin = value; }
        }
        private static bool flatCalcDone;
        public static bool FlatCalcDone
        {
            get { return flatCalcDone; }
            set { flatCalcDone = value; }
        }
        public static int apexHFR
        {
            get { return _apexHFR; }
            set { _apexHFR = value; }
        }
        public static int EnteredPID
        {
            get { return _enteredPID; }
            set { _enteredPID = value; }
        }
        public static double EnteredSlopeDWN
        {
            get { return _enteredSlopeDWN; }
            set { _enteredSlopeDWN = value; }
        }
        public static double EnteredSlopeUP
        {
            get { return _enteredSlopeUP; }
            set { _enteredSlopeUP = value; }
        }
        public static string Eqiup
        {
            get { return _equip; }
            set { _equip = value; }
        }
        private static bool sequenceRunning;
        public static bool SequenceRunning
        {
            get { return sequenceRunning; }
            set { sequenceRunning = value; }
        }
        private static bool astrometryRunning;
        public static bool AstrometryRunning
        {
            get { return astrometryRunning; }
            set { astrometryRunning = value; }
        }
        private static bool solveRequested = false;

        //  int EnteredPID;
        //   double EnteredSlopeUP;
        //   double EnteredSlopeDWN;

        private static bool FilterFocusOn = false;
        private static float FocusTime;
        //    private static bool startup = true;//used to ensure tab change only changes populates focuspos on startup
        private static int CaptureBin;
        private static bool FocusLocObtained = false;
        private static bool TargetLocObtained = false;

        string NebCamera;
        //  private static bool MountMoving = false;
        //  string GotoDoneCommand;
        //  private static string FocusLoc = "";
        // private static string TargetLoc = "";

        private static string _filtertext; //1-4-17
        public static string Filtertext
        {
            get { return _filtertext; }
            set { _filtertext = value; }
        }
        //    int metricHFR;
        private static int metricN = 0;
        private static int currentmetricN = 0;
        private static int[] AvgMetricHFR = null;
        private static bool MetricSample = false;
        // int AvgMetric = 0;
        //  int testMetricHFR = 0;
        private static bool DarksOn = false;
        //   private static bool filtersynced = false;
        //    private static bool filterMoving = false;
        private static int CaptureTime;
        private static string Nebname;
        private static int filterCountCurrent = 0;
        //  int totalsubs;
        //  int filternumber = 0;
        private static int subsperfilter;
        private static int subCountCurrent = 0; // 1-4-17

       
            protected static int currentfilter = 0;
            public static int CurrentFilter
            {
                get { return currentfilter; }
                set { currentfilter = value; }
            }
      
        //   int filter1used = 0;
        //   int filter2used = 0;
        //   int filter3used = 0;
        //   int filter4used = 0;
        //    int filter5used = 0;
        //   private double BestPos;
        //    int selectedrowcount;
        //   string selectedcell;
        //  bool roughvdone = false;
        private static int arraycountright = 0;
        private static int arraycountleft = 0;
        private static int enteredMaxHFR;



        private static int enteredMinHFR;
        //   private static bool portautoselect = false;
        //  string path2;
        //string portselected;
        // string port2selected;
        //  int fineVrepeatDone;
        //   bool fineVrepeatOn = false;
        //  int fineVrepeat;
        //      string equip;
        //   string equipPrefix;
        //    int rows;
        private static bool _gotoFocusOn = false;
        //int intersectPos;
        //   double slopeHFRdwn = 0;
        //   double slopeHFRup = 0;
        private static double XintUP;
        private static double XintDWN;
        //int PID;
        //      int HFRarraymin = 999;
        //  private static double maxarrayMax = 1;
        //   int apexHFR;
        //int backlashN = 10;
        private static float backlashSum = 0;//*****only used in 2 methods....find way of eliminating this varible!!


        //     int vLostStar = 0;
        //     float backlashAvg = 0;
        int backlashCount;
        bool backlashDetermOn;
        bool backlashOUT;
        int backlashPosIN;
        int backlashPosOUT;
        //   float backlash = 0;
        int backlash;
        double maxMax = 1;
        int posmaxMax;
        //abitrarily set min = 999 to start with
        int total = 0;
        int count;
        int min = 999;
        //   int posMin = 0;
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
       // int ffenable = 0; // 1-4-17 changed to bool, moved to GlobalVariables.   
       // bool ffenable = false;
       
        int repeatTotal = 0;
        int[] list = null;
        double[] listMax = null;
        double[] abc = null;
        int[] peat = null;
        double[] peatMax = null;
        int[] minHFRpos = null;
        double[] maxMaxPos = null;
        //    int posminHFR;
     //   bool tempon = false; // 1-4-17 changed to bool and moved to globalVariables. 
        int tempsum = 0;
        int tempavg = 0;
        int vv = 0;//total aquisitions for a given v-curve, N * repeat
        int travel;//maxstep in driver
                   //   int connect = 0;
                   //  int connect2 = 0;
        int temp = 100;
        int portopen = 0;
        int vcurveenable = 0;

        int closing = 0;
        //  int syncval;
        int templog;
        //  int MoveDelay; //helps ensure no focus movement during capture
        int roundto;
        int vN;
        string conString = WindowsFormsApplication1.Properties.Settings.Default.MyDatabase_2ConnectionString;
        int FocusPerSub;
        int FocusPerSubCurrent = 0;
        int FilterFocusGroupCurrent = 0;  // 11-13-16 changed to =0;

        double CaptureTime3;
        int sigma;
        double Low;
        double High;
        int DwnSz;
        void setarraysize()
        {
            arraysize1 = (int)numericUpDown5.Value;//coarse v N
            arraysize2 = (int)numericUpDown3.Value;//fine v N
        }

        void OnTabPageValidating(object sender, CancelEventArgs e)
        {

        }


        //****** 8-15-13 Go through and redo all checkbox 14 and 20 w. the method below


        //public bool ServerEnabled() // moved to global variables. 
        //{
        //    return checkBox20.Checked;
        //}

       

        //public static bool SlaveModeEnabled
        //{
        //    return checkBox14.Checked;
        //    //  set { checkBox14.Checked = value; } // the set is optional
        //}
        void positionbar()
        {
            progressBar2.Maximum = travel;
            progressBar2.Minimum = 0;
            progressBar2.Increment(10);
            progressBar2.Value = count;
            return;
        }
        #region logging
        //public void Log(Exception e)
        //{
        //    Log(e.Message);
        //}


        //  private string _logFromClass;
        //  public  string LogFromClass
        //  {
        //      get { return _logFromClass; }
        //      set { _logFromClass = value;
        //          OnLogFromClassChanged();
        //          }
        //      }
        //  public event System.EventHandler LogFromClassChanged;
        //  protected virtual void OnLogFromClassChanged()
        //  {
        //      if (LogFromClassChanged != null) LogFromClassChanged(this, EventArgs.Empty);
        //      Log(LogFromClass);
        //  }
        //  private void AppendText(string text)  ///causes cross thread opeation error
        //  {
        //      if (this.InvokeRequired)
        //      {
        //          this.Invoke(new Action<string>(AppendText), new object[] { text });
        //          return;
        //      }
        //      this.LogTextBox.Text += text;
        //  }

        //  public delegate void LogCallback(string text);
        ////  private System.Windows.Forms.TextBox ltb;


        public void Log(string text)
        {
            //if (LogTextBox.InvokeRequired)
            //{
            //    LogCallback method = new LogCallback(Log);
            //    LogTextBox.Invoke(method, new object[] { text });
            //    return;
            //}
            try // added 10-25-16
            {
                
                if (LogTextBox.Text != "")
                    LogTextBox.Text += Environment.NewLine;
                LogTextBox.Text += DateTime.Now.ToLongTimeString() + "  " + text;
                LogTextBox.SelectionStart = LogTextBox.Text.Length;
                LogTextBox.ScrollToCaret();
            }
            catch (Exception e) // added 10-25-16
            {
              //  Log("logging error line 667");
                FileLog2("Logging error " + e.ToString());

            }

        }
        #endregion

        void FillData()
        {
            try
            {
                FileLog2("filldata");
                using (SqlCeConnection con = new SqlCeConnection(conString))
                {
                    con.Open();
                    using (SqlCeDataAdapter a = new SqlCeDataAdapter("SELECT * FROM table1", con))
                    {

                        DataTable t = new DataTable();
                        a.Fill(t);
                        dataGridView1.DataSource = t;
                        a.Update(t);
                        dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Descending); // added 11-3-16 keep most recent on top

                        //    FileLog2(GetCreateFromDataTableSQL("table1", t)); //ytry to create the create string from here 7-25-14

                    }
                    con.Close();
                }
                ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";
                //   dataGridView1.CurrentCell = null;
                //   ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + textBox13.Text.ToString() + "'";

            }
            catch (Exception ex)
            {
                Log("FillData Error" + ex.ToString());
            }
        }





        //analyze SQL data  

        // 10-11-16  try to preserve selected row after refresh
        int selectedID, rowIndex, scrollIndex;
        bool IsSelectedRow;
        //see end of code





        void GetAvg()
        {
            try
            {
                FileLog2("GetAvg");

                // ******* 10-10-16   add code to get numbers from a singles selected row for full fram metric.  
                // if checkbox22.checked == true...  add to each below 
                // if (a row is selected) then skip the averaging....



                if (checkBox22.Checked == true)
                {
                    int selectedrowcount;
                    string selectedcell;
                    selectedrowcount = dataGridView1.SelectedCells.Count;
                    if (selectedrowcount == 0)
                    {
                        dataGridView1.ClearSelection();
                        int nRowIndex = dataGridView1.Rows.Count - 2;
                        if (nRowIndex > 0)
                            dataGridView1.Rows[nRowIndex].Selected = true;
                        else
                        {
                            Log("No Metric V-curve Data");
                            return;
                        }
                        //MessageBox.Show("No Row Selected");
                        //return;
                    }
                    else
                    {
                        selectedcell = dataGridView1.SelectedCells[0].Value.ToString();

                        _enteredPID = Convert.ToInt32(dataGridView1.SelectedCells[2].Value);
                        textBox12.Text = _enteredPID.ToString();
                        _enteredSlopeDWN = Convert.ToDouble(dataGridView1.SelectedCells[3].Value);
                        textBox3.Text = _enteredSlopeDWN.ToString();
                        _enteredSlopeUP = Convert.ToDouble(dataGridView1.SelectedCells[4].Value);
                        textBox10.Text = _enteredSlopeUP.ToString();
                        textBox16.Clear();
                        textBox14.Clear();
                        textBox15.Clear();

                        //   Log("PID=" + EnteredPID.ToString() + "    Dwn=" + _enteredSlopeDWN.ToString() + "    Up=" + _enteredSlopeUP.ToString());
                        //   MessageBox.Show("PID=" + EnteredPID.ToString() + "    Dwn=" + _enteredSlopeDWN.ToString() + "    Up=" + _enteredSlopeUP.ToString());
                    }

                }





                else
                {

                    using (SqlCeConnection con = new SqlCeConnection(conString))
                    {
                        con.Open();





                        using (SqlCeCommand com = new SqlCeCommand("SELECT AVG(PID) FROM table1 WHERE Equip = @equip", con))
                        {
                            com.Parameters.AddWithValue("@equip", _equip);
                            SqlCeDataReader reader = com.ExecuteReader();
                            while (reader.Read())
                            {
                                if (!reader.IsDBNull(0))
                                {
                                    int numb5 = reader.GetInt32(0);
                                    _enteredPID = numb5;
                                    textBox12.Text = _enteredPID.ToString();  // added 1-23-15
                                }
                                else
                                {
                                    textBox12.Clear();
                                    _enteredPID = 0;
                                }
                            }
                            reader.Close();
                        }

                        using (SqlCeCommand com1 = new SqlCeCommand("SELECT AVG(SlopeDWN) FROM table1 WHERE Equip = @equip", con))
                        {
                            com1.Parameters.AddWithValue("@equip", _equip);
                            SqlCeDataReader reader1 = com1.ExecuteReader();
                            while (reader1.Read())
                            {
                                if (!reader1.IsDBNull(0))
                                {
                                    double numb5 = reader1.GetDouble(0);
                                    double numb5Rnd = Math.Round(numb5, 5);
                                    textBox3.Text = numb5Rnd.ToString();
                                    _enteredSlopeDWN = numb5Rnd;
                                }
                                else
                                {
                                    textBox3.Clear();
                                    textBox14.Clear();
                                    _enteredSlopeDWN = 0;
                                }

                            }
                            reader1.Close();
                        }

                        using (SqlCeCommand com2 = new SqlCeCommand("SELECT AVG(SlopeUP) FROM table1 WHERE Equip = @equip", con))
                        {
                            com2.Parameters.AddWithValue("@equip", _equip);
                            SqlCeDataReader reader2 = com2.ExecuteReader();
                            while (reader2.Read())
                            {
                                if (!reader2.IsDBNull(0))
                                {
                                    double numb5 = reader2.GetDouble(0);
                                    double numb5Rnd = Math.Round(numb5, 5);
                                    textBox10.Text = numb5Rnd.ToString();
                                    _enteredSlopeUP = numb5Rnd;
                                }
                                else
                                {
                                    textBox10.Clear();
                                    textBox16.Clear();
                                    _enteredSlopeUP = 0;
                                }
                            }
                            reader2.Close();
                        }
                        using (SqlCeCommand com3 = new SqlCeCommand("SELECT AVG(BestHFR) FROM table1 WHERE Equip = @equip", con))
                        {
                            com3.Parameters.AddWithValue("@equip", _equip);
                            SqlCeDataReader reader3 = com3.ExecuteReader();
                            while (reader3.Read())
                            {
                                if (!reader3.IsDBNull(0))
                                {
                                    int numb6 = reader3.GetInt32(0);
                                    textBox15.Text = numb6.ToString();
                                }
                                else
                                {
                                    textBox15.Clear();

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
            }
            catch (Exception ex)
            {
                Log("GetAvg Error" + ex.ToString());
            }
        }
        private void WriteSQLdata()
        {
            try
            {
                FileLog2("WriteSQLData");
                using (SqlCeConnection con = new SqlCeConnection(conString))
                {

                    con.Open();
                    int num = _hFRarraymin;
                    int num2 = _apexHFR;
                    int num4 = _posminHFR;
                    float up = (float)_slopeHFRup;
                    float down = (float)_slopeHFRdwn;
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
                        com.Parameters.AddWithValue("@PID", _pID);
                        com.Parameters.AddWithValue("@SlopeDWN", down);
                        com.Parameters.AddWithValue("@SlopeUP", up);
                        com.Parameters.AddWithValue("@Number", rows + 1);
                        com.Parameters.AddWithValue("@equip", _equip);
                        com.Parameters.AddWithValue("@BestHFR", _hFRarraymin);
                        com.Parameters.AddWithValue("@FocusPos", _intersectPos);
                        com.ExecuteNonQuery();
                        rows++;
                    }
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                Log("WriteSQLData Error" + ex.ToString());
            }
        }

        DateTime lastRead = DateTime.MinValue;
        private void fileSystemWatcher2_Changed(object sender, FileSystemEventArgs e)
        {
            // simulation by writing fake bmp file results in mult FSW triggers.  try to ignore dupluicates
            // unrem for sim 


            //    delay(1); // add for sim testing only 11-3-16

            //    DateTime lastWriteTime = File.GetLastWriteTime(e.FullPath);
            //     if (lastWriteTime != lastRead)
            //     {
            //       lastRead = lastWriteTime;

            if (vDone != 1)
                vcurve();

            //    }


        }
        bool FineFocusAbort = false;
        // bool NebListenOn = false; // remd all 10-25-16
        //std dev, avg and gotofocus
        // bool focusing = false;
        private double BestPos;
        private int redo = 0;
        private int calcRedo = 0;
        private bool HFRtestON = false;
        private bool focusSampleComplete = false;
        private bool autoMetricVcurve = false;
        private void fileSystemWatcher3_Changed(object sender, FileSystemEventArgs e)
        {
            // int nn;
            FileLog2("fileSystemWatcher3 changed");
            textBox41.Refresh();//dual scope status textbox
            textBox41.Clear();

            if (currentfilter == 1)
                Filtertext = comboBox2.Text;
            if (currentfilter == 2)
                Filtertext = comboBox3.Text;
            if (currentfilter == 3)
                Filtertext = comboBox4.Text;
            if (currentfilter == 4)
                Filtertext = comboBox5.Text;
            //   watchforOpenPort();
         //   if ((portopen == 1 || usingASCOMFocus == true))
         //   {

                // int nn = 10;
                int current;
                int nn;
                roundto = 1; // 10-25-16
                if (checkBox22.Checked == true)
                    nn = (int)numericUpDown21.Value;
                else
                    nn = (int)numericUpDown5.Value; //coarse-v N
                if (HFRtestON == false)
                {
                    if (vProgress < nn)
                    {
                        int[] list = new int[nn];
                        if (checkBox22.Checked == true)//use metric
                        {
                            metricpath = Directory.GetFiles(GlobalVariables.Path2.ToString(), "metric*.fit");
                            current = GetMetric(metricpath, roundto);  //reads the current in-focus status from the HFR in the metric file name
                        }
                        else
                        {
                            fileSystemWatcher3.Filter = "*.bmp";
                            string[] filePaths = Directory.GetFiles(GlobalVariables.Path2.ToString(), "*.bmp");
                            //remd to use path2 settins    string[] filePaths = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.bmp");         
                            current = GetFileHFR(filePaths, roundto);
                            //Need to add check box to select if backup below is wanted
                            //     ****** fix  11-13-13   ******
                            if (current > 600)//was 500, not working w/ nb filters  ******changed to 900 to override was 600*****
                            {
                                //***** 6-28  still seems to goto focus position twice after abort and before first metric.
                                FineFocusAbort = true;
                                Array.Clear(list, 0, arraysize1);
                                Log("Focus star lost, using full frame metric");
                               FileLog2("Focus star lost, using full frame metric");
                                SetForegroundWindow(Handles.NebhWnd);
                                Thread.Sleep(1000);
                                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                                Thread.Sleep(2000);
                                checkBox22.Checked = true;
                                fileSystemWatcher3.EnableRaisingEvents = false;
                                if (!UseClipBoard.Checked)
                                {
                                    if (clientSocket.Client.Connected == true)
                                        clientSocket.Client.Disconnect(true);
                                }
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
                            //  textBox9.Text = stdev.ToString(); // remd 10-25-16
                            //     textBox14.Text = avg.ToString();
                            Log("Goto Focus :\t  N " + (vProgress + 1).ToString() + "\tHFR" + abc[vProgress].ToString() + "\t  Avg " + avg.ToString() + "\t  StdDev " + stdev.ToString());
                            FileLog2("Goto Focus" + "\t  " + abc[vProgress].ToString() + "\t  " + avg.ToString() + "\t" + (vProgress + 1).ToString() + "\t" + stdev.ToString());
                            //string strLogText = "Goto Focus" + "\t  " + abc[vProgress].ToString() + "\t  " + avg.ToString() + "\t" + (vProgress + 1).ToString() + "\t" + stdev.ToString(); //changed 7-25-14
                            //string path = textBox11.Text.ToString();
                            //string fullpath = path + @"\log.txt";
                            //StreamWriter log;
                            //if (!File.Exists(fullpath))
                            //{
                            //    log = new StreamWriter(fullpath);
                            //}
                            //else
                            //{
                            //    log = File.AppendText(fullpath);
                            //}
                            //if (vProgress == 0)
                            //{
                            //    log.WriteLine(DateTime.Now);
                            //    log.WriteLine("Type" + "\tHFR" + "\tAvg" + "\tN" + "\tStdDev");
                            //}
                            //log.WriteLine(strLogText);
                            //log.Close();
                            vProgress++;
                            if ((checkBox22.Checked == true) && (vProgress != nn))
                                MetricCapture();
                            toolStripStatusLabel1.Text = "Finding Focus " + vProgress.ToString();//**************added 2_29 for testing

                        }

                    }
                    if (vProgress == nn) // calculate focus point from HFR
                    {
                        focusSampleComplete = true;
                        //?  may need FSW3 enalbing event = false here for std dev use
                        if (_gotoFocusOn == true)
                        {

                            if (radioButton2.Checked == true)//use upslope
                            {
                                // Data d = new Data();
                                FillData();
                                GetAvg();

                                //  EnteredPID = Convert.ToInt32(textBox12.Text);
                                //    EnteredSlopeUP = Convert.ToDouble(textBox10.Text);
                                //      textBox14.Text = avg.ToString();
                                BestPos = count - (avg / _enteredSlopeUP) + (_enteredPID / 2);

                                //3-2-16
                                Log("Calculated focus point = " + Math.Round(BestPos).ToString());
                                FileLog2("Calculated focus point = " + Math.Round(BestPos).ToString());
                                if ((int)BestPos < 0 || (int)BestPos > travel)
                                {
                                    calcRedo++;
                                    Log("The Calculated Focus Point was out of bounds -- Repeat attempt " + redo.ToString());
                                    FileLog2("The Calculated Focus Point was out of bounds -- Repeat attempt " + redo.ToString());

                                    posMin = count; //reset to prev focus point
                                    BestPos = count;
                                    fileSystemWatcher3.EnableRaisingEvents = false;
                                    if (calcRedo < 4)
                                        gotoFocus();
                                    else
                                    {
                                        Log("The Focus Calcualtion Failed x 3 -- Aborted");
                                        FileLog2("The Focus Calcualtion Failed x 3 -- Aborted");
                                        Send("GotoFocus Calcualtion Error - GotoFocus Aborted");
                                        BestPos = count;
                                        calcRedo = 0;
                                    }

                                    return;
                                }
                                //check to make sure there wasn't gross miscalculation 9-9-15
                                if (!Simulator())
                                {
                                    if (((double)BestPos > (count + (int)numericUpDown39.Value * 2)) || ((double)BestPos < (count - (int)numericUpDown39.Value * 10)))
                                    {
                                        redo++;
                                        Log("The Calculated Focus Point was too far from sample point -- Repeat attempt " + redo.ToString());
                                        FileLog2("The Calculated Focus Point was too far from sample point -- Repeat attempt " + redo.ToString());

                                        posMin = count; //reset to prev focus point
                                        BestPos = count;
                                        fileSystemWatcher3.EnableRaisingEvents = false;
                                        if (calcRedo < 4)
                                            gotoFocus();
                                        else
                                        {
                                            Log("The Focus Calcualtion Failed x 3 -- Aborted");
                                            FileLog2("The Focus Calcualtion Failed x 3 -- Aborted");
                                            Send("GotoFocus Calcualtion Error - GotoFocus Aborted");
                                            BestPos = count;
                                            calcRedo = 0;
                                        }

                                        return;
                                    }



                                    else //if calculation seems ok...move to focus point then check w/ another exposure 
                                    {
                                        if (calcRedo > 0)
                                            calcRedo = 0;
                                        focuser.Move((int)BestPos);
                                        fileSystemWatcher3.EnableRaisingEvents = false;
                                    }
                                }
                                else // added 10-24-16
                                {
                                    focuser.Move((int)BestPos);
                                    fileSystemWatcher3.EnableRaisingEvents = false;
                                }
                                ////9-9-15
                                //// check to make sure HFR improved....it should always improve a little or sample point was too close to PID
                                //string[] filePaths = Directory.GetFiles(GlobalVariables.Path2.ToString(), "*.bmp");
                                //current = GetFileHFR(filePaths, roundto);
                                //Log("Test HFR = " + current.ToString() + " versus average smaple HFR = " + avg.ToString());
                                //FileLog2("Test HFR = " + current.ToString() + " versus average smaple HFR = " + avg.ToString());
                                //if (current > avg)
                                //{
                                //    Log("Focus did not improve, repeating on the other side of V-vurve");
                                //    FileLog2("Focus did not improve, repeating on the other side of V-vurve");
                                //    BestPos = count;
                                //    radioButton2.Checked = false;
                                //    radioButton3.Checked = true;
                                //    //    redo = true;
                                //    fileSystemWatcher3.EnableRaisingEvents = false;
                                //    gotoFocus();
                                //    return;
                                //}


                                //  gotopos(Convert.ToInt32(BestPos));//  4-24-14
                                //Thread.Sleep(1000);
                                //delay(1); 

                                //  fileSystemWatcher3.EnableRaisingEvents = false;

                                HFRtestON = true;

                                _gotoFocusOn = false;
                                Log("Goto Focus Position: " + ((int)BestPos).ToString());
                                FileLog2("Goto Focus Position: " + ((int)BestPos).ToString());
                                textBox4.Text = ((int)BestPos).ToString();

                                toolStripStatusLabel1.Text = "Focus Moving";
                                toolStripStatusLabel1.BackColor = System.Drawing.Color.Red;
                                //3-2-16 
                                while (focuser.IsMoving)
                                {
                                    delay(1); // 10-11-16 moved from space below on all similar x 3
                                    count = focuser.Position;
                                    textBox1.Text = count.ToString();

                                    positionbar();
                                    //  Log(focuser.IsMoving.ToString());
                                }
                                toolStripStatusLabel1.BackColor = System.Drawing.Color.WhiteSmoke;
                                toolStripStatusLabel1.Text = "Ready";


                            }
                            if (radioButton3.Checked == true)//use downslope
                            {
                                //Data d = new Data();
                                FillData();
                                GetAvg();

                                //    EnteredPID = Convert.ToInt32(textBox12.Text);
                                //    EnteredSlopeDWN = Convert.ToDouble(textBox3.Text);
                                // textBox14.Text = avg.ToString();
                                BestPos = count - (avg / _enteredSlopeDWN) - (_enteredPID / 2);
                                Log("Calculated focus point = " + BestPos.ToString());
                               FileLog2("Calculated focus point = " + BestPos.ToString());
                                //3-2-16
                                if ((int)BestPos < 0 || (int)BestPos > travel)
                                {
                                    calcRedo++;
                                    Log("The Calculated Focus Point was out of bounds -- Repeat attempt " + redo.ToString());
                                    FileLog2("The Calculated Focus Point was out of bounds -- Repeat attempt " + redo.ToString());

                                    posMin = count; //reset to prev focus point
                                    BestPos = count;
                                    fileSystemWatcher3.EnableRaisingEvents = false;
                                    if (calcRedo < 4)
                                        gotoFocus();
                                    else
                                    {
                                        Log("The Focus Calcualtion Failed x 3 -- Aborted");
                                        FileLog2("The Focus Calcualtion Failed x 3 -- Aborted");
                                        Send("GotoFocus Calcualtion Error - GotoFocus Aborted");
                                        BestPos = count;
                                        calcRedo = 0;
                                    }

                                    return;
                                }


                                //check calculation 9-9-15
                                if (!Simulator())
                                {
                                    if (((double)BestPos > (count + (int)numericUpDown39.Value * 10)) || ((double)BestPos < (count - (int)numericUpDown39.Value * 2)))
                                    {
                                        calcRedo++;
                                        Log("The Calculated Focus Point was too far from sample point -- Repeat attempt " + redo.ToString());
                                        FileLog2("The Calculated Focus Point was too far from sample point -- Repeat attempt " + redo.ToString());

                                        posMin = count;
                                        BestPos = count;
                                        fileSystemWatcher3.EnableRaisingEvents = false;
                                        if (calcRedo < 4)
                                            gotoFocus();
                                        else
                                        {
                                            Log("The Focus Calcualtion Failed x 3 -- Aborted");
                                            FileLog2("The Focus Calcualtion Failed x 3 -- Aborted");
                                            Send("GotoFocus Calculation Error - GotoFocus Aborted");
                                            BestPos = count;
                                            calcRedo = 0;
                                        }
                                        return;
                                    }


                                    else
                                    {
                                        if (calcRedo > 0)
                                            calcRedo = 0;
                                        fileSystemWatcher3.EnableRaisingEvents = false;
                                        //   focuser.Move((int)BestPos);
                                        gotopos((int)BestPos);  // 10-13-16 try so waits for focuser move to complete...while below doesn't seem to work....
                                    }
                                }
                                else // added 10-24-16
                                {
                                    fileSystemWatcher3.EnableRaisingEvents = false;
                                    //   focuser.Move((int)BestPos);
                                    gotopos((int)BestPos);
                                }
                                //3-2-16 

                                //these were not prev on the downslope one....  10-13-16
                                HFRtestON = true;

                                _gotoFocusOn = false;
                                Log("Goto Focus Position: " + ((int)BestPos).ToString());
                                FileLog2("Goto Focus Position: " + ((int)BestPos).ToString());
                                textBox4.Text = ((int)BestPos).ToString();
                                // end addt

                                toolStripStatusLabel1.Text = "Focus Moving";
                                toolStripStatusLabel1.BackColor = System.Drawing.Color.Red;
                                while (focuser.IsMoving)
                                {
                                    delay(1);
                                    count = focuser.Position;
                                    textBox1.Text = count.ToString();

                                    positionbar();
                                    //  Log(focuser.IsMoving.ToString());
                                }
                                toolStripStatusLabel1.BackColor = System.Drawing.Color.WhiteSmoke;
                                toolStripStatusLabel1.Text = "Ready";

                            }
                            //*****7-25-14 try moving this from below**********


                            //if (checkBox22.Checked == true)
                            //{
                            //    *********moved to test area below********unrem all if doesnt' work 
                            //        fileSystemWatcher3.EnableRaisingEvents = false; //added 7-25-14
                            //    _gotoFocusOn = false;
                            //    //    Log("Goto Focus Position: " + Convert.ToInt32(BestPos).ToString());  // this is redundant
                            //    textBox4.Text = ((int)BestPos).ToString();

                            //    serverStream = clientSocket.GetStream();
                            //    byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                            //    serverStream.Write(outStream, 0, outStream.Length);
                            //    serverStream.Flush();

                            //    Thread.Sleep(3000);
                            //    serverStream.Close();
                            //    SetForegroundWindow(Handles.NebhWnd);
                            //    Thread.Sleep(1000);
                            //    PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                            //    Thread.Sleep(1000);
                            //    NebListenOn = false;
                            //    // clientSocket.GetStream().Close();//added 5-17-12
                            //    //  clientSocket.Client.Disconnect(true);//added 5-17-12
                            //    clientSocket.Close();
                            //    HFRtestON = true; remd 10 - 13 - 16 does nothing
                            //}
                            FileLog2("Goto Focus Position " + ((int)BestPos).ToString() + " Current Filter_"); //7-25-14
                            string strLogText;
                            strLogText = "Goto Focus Position " + ((int)BestPos).ToString() + " Current Filter_";


                            HFRtestON = true; // added 10-13-16

                            //string path = textBox11.Text.ToString();
                            //string fullpath = path + @"\log.txt";
                            //StreamWriter log;
                            //if (!File.Exists(fullpath))
                            //{
                            //    log = new StreamWriter(fullpath);
                            //}
                            //else
                            //{
                            //    log = File.AppendText(fullpath);
                            //}

                            if (Filtertext != null)
                                FileLog2(strLogText + Filtertext.ToString());
                            else
                                FileLog2(strLogText);

                         
                        }
                              


                    }
                    if (FineFocusAbort == true)
                    {
                        //  FineFocusAbort = false;
                        gotoFocus();
                    }
                }
             
         //   } // end if (HFRtestON == false) from way above



            if (!Simulator())
            {
                //check for improvement.... 9-9-15
                if (HFRtestON == true)
                {
                    int nnn;
                    int test;
                    if (checkBox22.Checked == true)
                        nnn = (int)numericUpDown21.Value;
                    else
                        nnn = (int)numericUpDown5.Value;

                    fileSystemWatcher3.EnableRaisingEvents = true;
                    if ((checkBox22.Checked == true) & (vProgress == nnn))
                    {
                        vProgress++;
                        fileSystemWatcher3.Filter = "*.fit";
                        MetricCapture();
                        return;
                        //  Thread.Sleep((int)MetricTime * 1000);
                    }

                    vProgress++;

                    //have it take 4 more exposures after moving then test HFR to make sure not moving
                    //  if (vProgress > nnn)
                    //  {
                    //                      
                    if (((vProgress == nnn + 4) & (checkBox22.Checked == false)) || ((checkBox22.Checked == true) & (vProgress == nnn + 2)))
                    {
                        if (checkBox22.Checked == true)
                        {
                            metricpath = Directory.GetFiles(GlobalVariables.Path2.ToString(), "metric*.fit");
                            test = GetMetric(metricpath, roundto);
                        }
                        else
                        {
                            string[] filePaths2 = Directory.GetFiles(GlobalVariables.Path2.ToString(), "*.bmp");
                            test = GetFileHFR(filePaths2, roundto);
                        }

                        Log("Test HFR = " + test.ToString() + " versus average sample HFR = " + avg.ToString());
                        FileLog2("Test HFR = " + test.ToString() + " versus average sample HFR = " + avg.ToString());
                        if (test > avg)
                        {
                            if ((redo == 1) & (checkBox22.Checked == false))  //if single star fails twice can try metric.    
                            {
                                SetForegroundWindow(Handles.NebhWnd);
                                Thread.Sleep(1000);
                                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                                Thread.Sleep(2000);
                                checkBox22.Checked = true;
                                // 10-14-16 select most recent V-curve
                                Update();
                                GetAvg();
                                dataGridView1.ClearSelection();
                                int nRowIndex = dataGridView1.Rows.Count - 2;
                                if (nRowIndex > 0)
                                    dataGridView1.Rows[nRowIndex].Selected = true;
                                else
                                {
                                    Log("Unable to attempt Full Frame Metric focus due to no V-curve data, attempting auto V-curve");
                                    // BestPos = count;
                                    redo = 0;
                                    //  posMin = Convert.ToInt32(BestPos);
                                    autoMetricVcurve = true;
                                    HFRtestON = false;
                                    fileSystemWatcher3.EnableRaisingEvents = false;
                                    if (!checkBox22.Checked) // added 11-13-16
                                        checkBox22.Checked = true;
                                    button3.PerformClick(); // 11-3-16 was buton 25 "metric V" (**** untested ****)
                                    return;
                                }
                                //
                                fileSystemWatcher3.EnableRaisingEvents = false;
                                if (!UseClipBoard.Checked)
                                {
                                    if (clientSocket.Client.Connected == true)
                                        clientSocket.Client.Disconnect(true);
                                }
                                checkBox22.Checked = true;
                                Log("Repeat Focus did not improve - Attemping full frame metric");
                                FileLog2("Repeat Focus did not improve - Attemped full frame metric");
                                Send("Repeat Focus did not improve - Attemping full frame metric");
                                BestPos = count;
                                redo = 0;
                                posMin = Convert.ToInt32(BestPos);
                                HFRtestON = false;
                                fileSystemWatcher3.EnableRaisingEvents = false;
                                gotoFocus();
                                return;

                            }
                            if (redo == 1)
                            {
                                Log("Repeat Focus did not improve - Aborting");
                                FileLog2("Repeat Focus did not improve - Aborting");
                                Send("Repeat Focus did not improve - Aborting");
                                fileSystemWatcher3.EnableRaisingEvents = false;
                                HFRtestON = false;
                                BestPos = count;
                                redo = 0;
                                return;
                            }
                            Log("Focus did not improve, attempting sample repeat on the other side of V-curve");
                            FileLog2("Focus did not improve, repeat on the other side of V-curve");
                            Send("GotoFocus Improvement Test Failed - trying other side of V-curve");
                            redo = 1;
                            BestPos = count; //reset to prev position
                            // switch slopes
                            if (radioButton3.Checked == true)
                            {
                                radioButton3.Checked = false;
                                radioButton2.Checked = true;
                            }
                            else
                            {
                                radioButton2.Checked = false;
                                radioButton3.Checked = true;
                            }
                            //    redo = true;
                            fileSystemWatcher3.EnableRaisingEvents = false;
                            HFRtestON = false;
                            if (checkBox22.Checked == true)
                            {
                                fileSystemWatcher3.EnableRaisingEvents = false; //added 7-25-14
                                _gotoFocusOn = false;
                                //    Log("Goto Focus Position: " + Convert.ToInt32(BestPos).ToString());  // this is redundant
                                textBox4.Text = ((int)BestPos).ToString();

                                // 11-8-16
                                if (!UseClipBoard.Checked)
                                {

                                    serverStream = clientSocket.GetStream();
                                    byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                                    serverStream.Write(outStream, 0, outStream.Length);
                                    serverStream.Flush();

                                    Thread.Sleep(3000);
                                    serverStream.Close();
                                    SetForegroundWindow(Handles.NebhWnd);
                                    Thread.Sleep(1000);
                                    PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                                    Thread.Sleep(1000);
                                    //   NebListenOn = false;
                                    // clientSocket.GetStream().Close();//added 5-17-12
                                    //  clientSocket.Client.Disconnect(true);//added 5-17-12
                                    clientSocket.Close();
                                }

                                else
                                {
                                    //  delay(1);
                                    Clipboard.Clear();
                                    //  delay(1);
                                    //  SetForegroundWindow(Handles.NebhWnd);
                                    //  delay(1);
                                    ////  Thread.Sleep(1000);
                                    //  PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                                    delay(1);
                                    //Clipboard.SetText("//NEB Listen 0");
                                    //msdelay(750);
                                }
                            }
                            //  posMin = Convert.ToInt32(BestPos);   // no this keeps the previous focus position....which might be worth trying on the other side of 
                            //  ...of the curve but it's not the previou foucs point.  will leave as the old focus point and try on other side of curve.  
                            Log("Using previous focus position - " + posMin.ToString());
                            FileLog2("Using previous focus position - " + posMin.ToString()); //may not be right
                            textBox4.Text = posMin.ToString();  // 10-13-16
                            gotoFocus();
                            return;
                        }
                        HFRtestON = false;

                        Log("Focus Improved at " + Convert.ToInt32(BestPos).ToString());
                        FileLog2("Focus Improved - Final Focus Position = " + Convert.ToInt32(BestPos).ToString());
                        redo = 0;
                        posMin = (int)BestPos;  //   10-13-16  posmin as last good focus position in case next round of focus fails
                        fileSystemWatcher3.EnableRaisingEvents = false; // added 9-9-15 due to testing above. 

                        // try add 10-13-16
                        //FilterFocusOn = false;
                        //if (!IsSlave())
                        //{

                        //fileSystemWatcher4.EnableRaisingEvents = true;//move this to last 3-4 from under abort/sleep
                        //NebCapture();

                        //}
                        // end add




                        if (checkBox22.Checked == true)
                        {
                            fileSystemWatcher3.EnableRaisingEvents = false; //added 7-25-14
                            _gotoFocusOn = false;
                            //    Log("Goto Focus Position: " + Convert.ToInt32(BestPos).ToString());  // this is redundant
                            textBox4.Text = ((int)BestPos).ToString();
                            if (!UseClipBoard.Checked)
                            {
                                serverStream = clientSocket.GetStream();
                                byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                                serverStream.Write(outStream, 0, outStream.Length);
                                serverStream.Flush();

                                Thread.Sleep(3000);
                                serverStream.Close();
                                SetForegroundWindow(Handles.NebhWnd);
                                Thread.Sleep(1000);
                                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                                Thread.Sleep(1000);
                                //  NebListenOn = false;
                                // clientSocket.GetStream().Close();//added 5-17-12
                                //  clientSocket.Client.Disconnect(true);//added 5-17-12
                                clientSocket.Close();
                            }

                            else
                            {
                                // delay(1);
                                Clipboard.Clear();
                                //SetForegroundWindow(Handles.NebhWnd);
                                //delay(1);
                                ////Thread.Sleep(1000);
                                //PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                                //// Thread.Sleep(1000);
                                delay(1);
                                //delay(1);
                                //Clipboard.SetText("//NEB Listen 0");
                                //msdelay(750);
                            }
                        }
                    }


                }
            }
            else
            {
                if (focusSampleComplete)
                {
                    Log("Simulator -- quality check disabled");
                   FileLog2("Simulator -- quality check disabled");

                    // added 10-24-16 


                    // ********  11-8-16 seems to do this even if not done 

                    HFRtestON = false;
                    //    focusSampleComplete = false; // 11-8-16 was true
                    posMin = (int)BestPos;  //   10-13-16  posmin as last good focus position in case next round of focus fails
                    fileSystemWatcher3.EnableRaisingEvents = false; // added 9-9-15 due to testing above. 

                    // try add 10-13-16
                    //FilterFocusOn = false;
                    //if (!IsSlave())
                    //{

                    //fileSystemWatcher4.EnableRaisingEvents = true;//move this to last 3-4 from under abort/sleep
                    //NebCapture();

                    //}
                    // end add

                }


                if (checkBox22.Checked == true)
                {
                    fileSystemWatcher3.EnableRaisingEvents = false; //added 7-25-14
                    _gotoFocusOn = false;
                    //    Log("Goto Focus Position: " + Convert.ToInt32(BestPos).ToString());  // this is redundant
                    textBox4.Text = ((int)BestPos).ToString();

                    //11-8-16
                    if (!UseClipBoard.Checked)
                    {
                        serverStream = clientSocket.GetStream();
                        byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                        serverStream.Write(outStream, 0, outStream.Length);
                        serverStream.Flush();

                        Thread.Sleep(3000);
                        serverStream.Close();
                        SetForegroundWindow(Handles.NebhWnd);
                        Thread.Sleep(1000);
                        PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                        Thread.Sleep(1000);
                        //    NebListenOn = false;
                        // clientSocket.GetStream().Close();//added 5-17-12
                        //  clientSocket.Client.Disconnect(true);//added 5-17-12
                        clientSocket.Close();
                    }
                    else
                    {
                        // delay(1);
                        Clipboard.Clear();
                        // delay(1);
                        // SetForegroundWindow(Handles.NebhWnd);
                        // delay(1);
                        //// Thread.Sleep(1000);
                        // PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                        //// Thread.Sleep(1000);

                        // delay(1);
                        //Clipboard.SetText("//NEB Listen 0");
                        //msdelay(750);
                    }
                }
            }



            // end test for improvement addition
        


            // 10-13-16 this was copied from above 
            if ((FilterFocusOn == true) & (HFRtestON == false) & (focusSampleComplete == true)) // 10-13-16 add focusSampleComplete 
            {

                if (checkBox22.Checked == true)
                {
                    /*//  all below remd 7-1-14 this fixed the problem w/metric in  filterfocus.

                  //   ShowWindow(Handles.NebhWnd, SW_SHOW);
                     //  ShowWindow(NebhWnd, SW_RESTORE);
                     SetForegroundWindow(Handles.NebhWnd);//? may not need 3-4
                   //  Thread.Sleep(1000);
                      delay(1);//may not need 3-4 both (below too) were 1000 6-1
                     SetForegroundWindow(Handles.Aborthwnd);
                     PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                    // Thread.Sleep(1000);
                     delay(1);//was 1000
                   //  NebListenOn = false;
                     */

                    if (metricpath[0] != null)
                        File.Delete(metricpath[0]);

                }


                if (checkBox22.Checked == false) //not using metric   
                {
                    //  HFRtestON = true; // added 10-13-16
                    //  FilterFocusOn = false;move to while belwo
                    ShowWindow(Handles.NebhWnd, SW_SHOW);
                    //  ShowWindow(NebhWnd, SW_RESTORE);
                    SetForegroundWindow(Handles.NebhWnd);//? may not need 3-4
                    Thread.Sleep(500);//may not need 3-4 both (below too) were 1000 6-1
                    SetForegroundWindow(Handles.Aborthwnd);
                    PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                    Thread.Sleep(500);//was 1000
                                      //   NebListenOn = false;
                                      //     if (FineFocusAbort == true)
                                      //       FineFocusAbort = false;

                }

                if (checkBox10.Checked == false)//cb10 is focus in image frame, if not needs to slew back
                {
                    FileLog2("Slew to target");
                    toolStripStatusLabel1.Text = "Slewing to Target";
                    this.Refresh();

                    // toolStripStatusLabel1.Text = "Slewing to Target Location";
                    //   this.Refresh();
                    //    FocusDoneTargetReturn();
                    //   Thread.Sleep(1000);
                    //  MessageBox.Show("simulate Slew");

                    GotoTargetLocation();

                    //while (MountMoving == true)
                    //{
                    //    Thread.Sleep(50); // pause for 1/20 second
                    //    System.Windows.Forms.Application.DoEvents();
                    //}
                    //if (MountMoving == false)
                    //{
                    button35.Text = "At Target";
                    button35.BackColor = System.Drawing.Color.Lime;
                    button33.Text = "Goto";
                    button33.UseVisualStyleBackColor = true;
                    //   }
                    //    Thread.Sleep(SlewDelay);

                    //if (Handles.PHDVNumber == 2)  //resumes in GotoTargetLocation()
                    //    resumePHD2();
                    //else
                    //ResumePHD();
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
                //if (checkBox19.Checked == true)
                //    if (Handles.PHDVNumber == 2)
                //        resumePHD2();
                //    else
                //    ResumePHD();


                if (!IsSlave() && (checkBox22.Checked == false))
                {
                    if (!UseClipBoard.Checked)
                    {
                        clientSocket.GetStream().Close();
                        clientSocket.Close();// rem'd 5-17-12                                    
                    }
                }

                if (IsSlave())
                {
                    // ***** TODO fix for clipboard and slave
                    clientSocket2.GetStream().Close();
                    clientSocket2.Close();
                }


                delay(1);
                //    SlavePaused = true;



                if (IsServer())
                {
                    Log("Waiting for slave focus");
                    toolStripStatusLabel1.Text = "Waiting for slave focus";
                    this.Refresh();
                    while (working == true)
                    {
                        WaitForSequenceDone("Fine focus done", GlobalVariables.NebSlavehwnd);
                        //  Thread.Sleep(50);
                    }
                    working = true;
                    Log("Slave focus done");
                }

                // remd 10-12-16 moved to the very end
                // unremd 10-13-16
                FilterFocusOn = false;
                if (!IsSlave())
                {
                    focusSampleComplete = false; // added 11-12-16
                    fileSystemWatcher4.EnableRaisingEvents = true;//move this to last 3-4 from under abort/sleep
                    NebCapture();
                    FileLog2("FSW4 - enabled, NebCapture started from FSW3");
                   
                    //*************un remd 3_4
                                 /*
                                 if (backgroundWorker2.IsBusy != true)
                                 {
                                     // Start the asynchronous operation.
                                     backgroundWorker2.RunWorkerAsync();
                                 }
 */
                }



            }
            // end copy/paste




        }

        public bool Simulator()
        {

            if (GlobalVariables.Nebcamera.ToString() == "Simulator")
            {
                return true;
            }
            else
                return false;
        }

        void standby()
        {
            fileSystemWatcher2.EnableRaisingEvents = false;
            fileSystemWatcher3.EnableRaisingEvents = false;
            fileSystemWatcher1.EnableRaisingEvents = false;
            fileSystemWatcher4.EnableRaisingEvents = false;//added 3_13
            fileSystemWatcher5.EnableRaisingEvents = false;//added to test  metricHFR
            fileSystemWatcher7.EnableRaisingEvents = false;//added 12-29-16
            if (closing != 1)
            {
                //      port.DiscardOutBuffer();
                //      port.DiscardInBuffer();
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
            return;  // added 11-20-14

        }
        //gotoposhere
        //  bool moveOut;
        void gotopos(Int32 value)
        {

            try
            {
                if (devId2 == null)
                {
                    MessageBox.Show("Not connected to focuser", "scopefocus");
                    return;
                }
                // Log("goto " + value.ToString());
                if (comboBox7.SelectedItem == null)//don't allow movement without equip selection //can cause movement in wrong direction due to reverse not being appropriate
                {

                    MessageBox.Show("Not connected to focuser", "scopefocus");
                    return;
                }
                if (value > travel)
                {
                    DialogResult result1 = MessageBox.Show("Goto Exceeds Max Travel", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (result1 == DialogResult.OK)
                    {
                        numericUpDown6.Value = count;//puts the current position in the goto selector as a safe number
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
                //{
                //   if (numericUpDown6.Value == focuser.Position)//don't move if already there.  *****this inhibits moves from fwd/rev buttons  *******
                //        return;

                //focus moving here for toolstrip rest see 1877
                toolStripStatusLabel1.Text = "Focus Moving";
                toolStripStatusLabel1.BackColor = System.Drawing.Color.Red;
                this.Refresh(); // added 10-24-16

                focuser.Move(value);


                //************added 11-3-13 may not be needed  ************
                //     BUT     *************seems like ascom focusers need some delay while moving  
                //*****  4-22-14 if IsMoving causes problems add a checkbox to allow disable. 
                while (focuser.IsMoving)
                {
                    delay(1);
                    count = focuser.Position;
                    textBox1.Text = count.ToString();

                    positionbar();   //**** remd 11-20-14
                                     //  Log(focuser.IsMoving.ToString());
                }
                // end add
                delay(1);
                count = focuser.Position;
                textBox1.Text = count.ToString();
                //   textBox1.Text = focuser.Position.ToString();
                if (!ContinuousHoldOn)
                    focuser.Halt();
                toolStripStatusLabel1.BackColor = System.Drawing.Color.WhiteSmoke;
                toolStripStatusLabel1.Text = "Ready";
                this.Refresh();
                positionbar();
            }

            catch (Exception e)
            {
                Log("GotoPos Error" + e.ToString());
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
                current2 = Convert.ToInt32(size);
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
            FileLog2("V-Curve");
            int vLostStar = 0;
            float backlashAvg = 0;

            //  Data d = new Data();
            //  try
            //  {



            double maxarrayMax = 1;
            int backlashN = 10;
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
            roundto = 1;  // 10-25-16
            int MoveDelay = (int)numericUpDown9.Value;

            vcurveenable = 1;
            if (GlobalVariables.Tempon)
            {
                fileSystemWatcher1.EnableRaisingEvents = true;
            }
            if ((vProgress == 0) & (GlobalVariables.Tempon))
            {
                int finegoto = posMin - ((((int)numericUpDown3.Value) / 2) * ((int)numericUpDown8.Value));
                //fine v-curve goes to N/2 * step size in from the focus position -- V should be centered
                gotopos(Convert.ToInt32(finegoto));
            }

            if (GlobalVariables.FFenable)
            {
                textBox17.Text = (GlobalVariables.FineRepeatDone + 1).ToString();//exposure repeat
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
            if (GlobalVariables.FFenable)
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
            int testMetricHFR;
            int avg = 0;
            int current = 0;
            int closestHFR = 1;
            long closestMax = 1;
            if (vDone != 1)
            {
                if (checkBox22.Checked == true)
                {
                    //begin test metriHFR
                    metricpath = Directory.GetFiles(GlobalVariables.Path2.ToString(), "metric*.fit");
                    try
                    {
                        testMetricHFR = GetMetric(metricpath, roundto);
                        textBox25.Text = testMetricHFR.ToString();
                        //  Log("MetricHFR" + testMetricHFR.ToString());
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
                //  int[] templist = new int[((vN * repeatTotal) + 1)]; //7-25-14 chesnted to below
                int[] templist = new int[((vN * repeatTotal) + 0)];
                string[] filePaths = Directory.GetFiles(GlobalVariables.Path2.ToString(), "*.bmp");
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
                        listMax[vProgress] = avgMax;  //???? vProgress is 10 bu array is only 0-9?????

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
                // try some lost star handling  7-27-14  ver 19

                if ((GlobalVariables.FFenable) & (!GlobalVariables.Tempon) & (repeatDone == 1) & (backlashDetermOn == false)) //check only for fine V-curve
                {
                    if ((vProgress > 2) && (vProgress < (vN - 2)))//don't use first/last stars
                    {
                        if (vLostStar > 0)//if lost next ones will be high as well
                        {
                            if (list[vProgress] > 600)
                            {
                                vLostStar++;
                                Log("HFR " + list[vProgress].ToString() + " lost star?");
                            }
                        }
                        if (vLostStar == 3)
                        {
                            vLostStar = 0;
                            fileSystemWatcher2.EnableRaisingEvents = false;
                            fileSystemWatcher5.EnableRaisingEvents = false;
                            Send("V-curve aborted, lost star?");
                            MessageBox.Show("Possible Lost Star, V-curve aborted");
                        }

                        if ((((double)list[vProgress] / (double)list[vProgress - 1]) > 1.5) || (((double)list[vProgress] / (double)list[vProgress - 1]) < .5)) //find outlier
                        {
                            vLostStar++;
                            Log("HFR " + list[vProgress].ToString() + " lost star?");

                        }
                    }
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
                    string strLogText3 = "Fine-V: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + count.ToString() + " \t" + avg.ToString() + "\t" + avgMax.ToString();
                    positionbar();

                    if ((!GlobalVariables.Tempon) & (GlobalVariables.FFenable != true) & (repeatDone == 1))
                    {
                        chart1.Series[0].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg)); //chart course v curve
                    }
                    if ((GlobalVariables.FFenable == true) & (!GlobalVariables.Tempon) & (repeatDone == 1) & (backlashDetermOn == false))
                    {
                        if ((avg < enteredMinHFR) || (avg > enteredMaxHFR) & (radioButton1.Checked == true))
                        {
                            chart1.Series[2].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg)); // chart data not used for calcs < min and > max
                        }
                        else
                        {
                            chart1.Series[1].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg));//charte data used for calcs
                        }
                    }

                    //string path = textBox11.Text.ToString();
                    //string fullpath = path + @"\log.txt";

                    //// Create a writer and open the file:
                    //StreamWriter log;

                    //if (!File.Exists(fullpath))
                    //{
                    //    log = new StreamWriter(fullpath);
                    //}
                    //else
                    //{
                    //    log = File.AppendText(fullpath);
                    //}
                    // Write to the file:
                    if (backlashDetermOn == false)
                    {
                        if ((GlobalVariables.FFenable) & (vProgress == 0))
                        {
                            //    log.WriteLine(DateTime.Now);
                            //    log.WriteLine("Fine-V" + "\tTemp" + "\t Pos" + "\tN" + "\tHFR" + "\tminHFR" + "\tposmin");
                            // }
                            FileLog2(DateTime.Now.ToString());
                            //   FileLog2("Fine-V" + "\tTemp" + "\t Pos" + "\tN" + "\tHFR" + "\tminHFR" + "\tposmin");
                        }
                        if ((vProgress == 0) & (!GlobalVariables.Tempon))
                        {
                            //log.WriteLine(DateTime.Now);
                            //log.WriteLine("Type  N:" + "\t  Pos" + "\tHFRAvg" + "\tMaxAvg");
                            FileLog2(DateTime.Now.ToString());
                            FileLog2("Type  N:" + "\t  Pos" + "\tHFRAvg" + "\tMaxAvg");
                        }
                        if ((GlobalVariables.Tempon) & (vProgress == (vN - 1)))
                        {
                            if (templog == 0)
                            {
                                //     FileLog2(DateTime.Now.ToString());
                                FileLog2("Type" + "\tAvgTemp" + "\tposMin");
                                // log.WriteLine(strLogText2);
                                templog = 1;
                                // log.Close();
                            }

                            FileLog2(strLogText2);//was log.WriteLine and next 2 below

                        }
                        if ((!GlobalVariables.Tempon) & (GlobalVariables.FFenable != true))
                        {
                            FileLog2(strLogTextA);
                        }
                        if ((!GlobalVariables.Tempon) & (GlobalVariables.FFenable == true))
                        {
                            FileLog2(strLogText3);
                        }
                    }

                    //    log.Close();
                    repeatDone = 0;
                    //reset repeat
                    repeatProgress = 0;
                    sum = 0;
                    if (vProgress < (vN - 1))
                    {
                        int step = (int)numericUpDown2.Value;
                        if (GlobalVariables.FFenable)
                        {
                            //defines Fine-V step size
                            step = (int)numericUpDown8.Value;
                        }
                        if (backlashDetermOn != true)
                        {
                            vProgress++;
                            vv++;
                            fileSystemWatcher2.EnableRaisingEvents = false; //****  11-20-14 this is added to allow focuser to complete movement before using the next exposure
                            fileSystemWatcher5.EnableRaisingEvents = false;
                            gotopos(Convert.ToInt32(count + step));
                            Thread.Sleep(MoveDelay);//helps prevent focus mvmnt during capture
                            fileSystemWatcher2.EnableRaisingEvents = true;
                            fileSystemWatcher5.EnableRaisingEvents = true;

                        }



                        //  }
                        if (backlashDetermOn == true)
                        {
                            if (backlashOUT == false)
                            {
                                if (vProgress < (vN - 1))
                                {
                                    gotopos(Convert.ToInt32(count - step));
                                }
                                vProgress++;
                                vv++;
                                Thread.Sleep(MoveDelay);
                            }
                            if (backlashOUT == true)
                            {
                                if (vProgress < (vN - 1))
                                {
                                    gotopos(Convert.ToInt32(count + step));
                                }
                                vProgress++;
                                vv++;
                                Thread.Sleep(MoveDelay);
                            }
                        }

                    }
                    else // 7-25-14
                        vDone = 1; // 7-25-14
                }


            }
            //V=1...skips to here.     
            if (vDone == 1)  //**********  7-27-14 seems like this is repeating...
            {
                fileSystemWatcher2.EnableRaisingEvents = false;
                fileSystemWatcher5.EnableRaisingEvents = false; //added to test metritHFR     

                for (int arraycount = 0; arraycount < vProgress; arraycount++)
                {
                    if (list[arraycount] < _hFRarraymin)
                    {
                        _hFRarraymin = list[arraycount];
                        _posminHFR = minHFRpos[arraycount];
                        min = _hFRarraymin;
                        _apexHFR = arraycount;
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
                        if ((posmaxMax == _posminHFR) & (MaxtooHi == false))
                        {
                            posMin = posmaxMax;

                        }

                    }
                }

                //make sure v-curve is symetric
                //                                                                                   
                if (radioButton1.Checked == true)
                {                                //below was vN/2 +4 and -4 changed to +/- vN/4 11-24
                    if (((_apexHFR > (vN / 2 + vN / 4)) || (_apexHFR < (vN / 2 - vN / 4))) & (GlobalVariables.FFenable) & (backlashDetermOn == false))
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
                if ((GlobalVariables.FFenable) & (radioButton1.Checked == true)) // save data on
                {
                    try
                    {

                        //first figure out how many points on each side of v-curve based on entered min/max
                        for (int arraycount = 0; arraycount < _apexHFR; arraycount++)
                        {
                            if ((list[arraycount] < enteredMaxHFR) & (list[arraycount] > enteredMinHFR))
                            {
                                arraycountright++;
                            }
                        }
                        for (int arraycount2 = _apexHFR; arraycount2 < vN; arraycount2++)
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
                        for (int arraycount = 0; arraycount < _apexHFR; arraycount++)
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
                        for (int arraycount2 = _apexHFR; arraycount2 < vN; arraycount2++)
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

                        //  Data d = new Data();
                        //    _slopeHFRdwn = (1 / GetSlope(HFRdwn, HFRposdwn));  // ? 1/ because x and y are reverse here and below.  could maybe just reverse the GetSlope(HFRposdwn, HFRdwn)
                        //    _slopeHFRup = (1 / GetSlope(HFRup, HFRposup));

                        //try this correction, for some reason up slope is neg and dwnslope is pos should be theother way around.  if wrong revert back to above 2 line
                        _slopeHFRdwn = GetSlope(HFRposdwn, HFRdwn);
                        _slopeHFRup = GetSlope(HFRposup, HFRup);


                        //use point 1/2 the way 
                        XintDWN = minHFRpos[_apexHFR - vN / 4] - list[_apexHFR - vN / 4] / _slopeHFRdwn;
                        XintUP = minHFRpos[_apexHFR + vN / 4] - list[_apexHFR + vN / 4] / _slopeHFRup;
                        _pID = (int)XintDWN - (int)XintUP;
                        posMin = _posminHFR;
                        //use point half way up each side to calc intersect position can be used as relative focus position
                        _intersectPos = (int)GetIntersectPos((double)minHFRpos[_apexHFR - vN / 4], (double)minHFRpos[_apexHFR + vN / 4], (double)list[_apexHFR - vN / 4], (double)list[_apexHFR + vN / 4], _slopeHFRdwn, _slopeHFRup);
                        //all vN/2 above were 5

                        textBox4.Text = posMin.ToString();
                        textBox15.Text = _hFRarraymin.ToString();
                        WriteSQLdata();
                        FillData();
                        Log("Slope: N " + (vProgress + 1).ToString() + "\tslopeUP" + _slopeHFRup.ToString() + " \tSlopeDWN" + _slopeHFRdwn.ToString() + "\tIntersect" + _intersectPos.ToString() + "\tPID" + _pID.ToString() + "\t" + Filtertext);
                        textBox18.Enabled = true;
                        textBox20.Enabled = true;

                    }

                    catch
                    {
                        MessageBox.Show("there was an error obtaining v-curve data.  Possibly lost star or too dim, try again", "scopefocus");
                        return;
                    }


                }

                else
                {
                    posMin = _posminHFR;
                }
                textBox4.Text = posMin.ToString();
                textBox1.Text = focuser.Position.ToString();

                //string path = textBox11.Text.ToString();  // 7-25-14
                //string fullpath = path + @"\log.txt";
                //string fullpathData = path + @"\data.txt";
                //StreamWriter log;
                //StreamWriter logData;

                //if (!File.Exists(fullpath))
                //{
                //    log = new StreamWriter(fullpath);
                //    logData = new StreamWriter(fullpathData);
                //}
                //else
                //{
                //    log = File.AppendText(fullpath);
                //    logData = File.AppendText(fullpathData);

                //}
                Log("V-Curve: Results:" + "\tPos" + posMin.ToString() + " \t  minHFR" + min.ToString() + "\tmaxMax" + avgMax.ToString() + "\t" + Filtertext);
                // Log("Posmin: " + posMin.ToString() + "\tminHFR " + min.ToString() + "\tminHFRpos " + _posminHFR.ToString());
                FileLog2("Posmin: " + posMin.ToString() + "\tminHFR " + min.ToString() + "\tminHFRpos " + _posminHFR.ToString() + "\tmaxMAx " + maxarrayMax.ToString() + "\tmaxMaxPos " + posmaxMax.ToString());
                //removed & radiobutton1,checked from below
                if (GlobalVariables.FFenable)
                {
                    for (int xx = 0; xx < vN; xx++)
                    {
                        if (xx == 0)
                        {
                            //log.WriteLine(DateTime.Now);
                            //log.WriteLine("Spredsheet friendly version of data above");
                            //log.WriteLine(Filtertext);
                            FileLog2(DateTime.Now.ToString());
                            FileLog2("Spredsheet friendly version of data above");
                            FileLog2(Filtertext);

                        }
                        // log.WriteLine("HFR " + list[xx].ToString() + "\tPosition " + minHFRpos[xx].ToString()); // 7-25-14
                        FileLog2("HFR " + list[xx].ToString() + "\tPosition " + minHFRpos[xx].ToString());
                    }
                    //   logData.WriteLine(DateTime.Now + "\t" + vN.ToString() + "\t" + _slopeHFRdwn.ToString() + "\t" + _slopeHFRup.ToString() + "\t" + XintDWN.ToString() + "\t" + XintUP.ToString() + "\t" + _pID.ToString() + "\t" + _apexHFR.ToString() + "\t" + Filtertext);
                    FileLog2(DateTime.Now + "\t" + vN.ToString() + "\t" + _slopeHFRdwn.ToString() + "\t" + _slopeHFRup.ToString() + "\t" + XintDWN.ToString() + "\t" + XintUP.ToString() + "\t" + _pID.ToString() + "\t" + min.ToString() + "\t" + Filtertext);
                }
                //  log.Close();
                //   logData.Close();
                if ((vDone == 1) & (GlobalVariables.FFenable))
                {

                    fileSystemWatcher1.EnableRaisingEvents = false;
                    fileSystemWatcher2.EnableRaisingEvents = false;
                    fileSystemWatcher5.EnableRaisingEvents = false; // added to test metric HFR
                                                                    //handle repeated fine v curves
                    if (GlobalVariables.FineVRepeat > 1)
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
                            //11-8-16
                            if (!UseClipBoard.Checked)
                            {

                                //10-11-16  
                                serverStream = clientSocket.GetStream();
                                byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                                serverStream.Write(outStream, 0, outStream.Length);
                                serverStream.Flush();

                                Thread.Sleep(3000);
                                serverStream.Close();
                                SetForegroundWindow(Handles.NebhWnd);
                                Thread.Sleep(1000);
                                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                                Thread.Sleep(1000);
                                //    NebListenOn = false;
                                // clientSocket.GetStream().Close();//added 5-17-12
                                //  clientSocket.Client.Disconnect(true);//added 5-17-12
                                clientSocket.Close();
                                //  Thread.Sleep(500);
                                //   SendKeys.SendWait("~");
                                //   SendKeys.Flush();
                            }
                            else
                            {
                                Clipboard.Clear();
                                // delay(1);
                                // SetForegroundWindow(Handles.NebhWnd);
                                // delay(1);
                                // //Thread.Sleep(1000);
                                // PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                                //// Thread.Sleep(1000);

                                delay(1);
                                //Clipboard.SetText("//NEB Listen 0");
                                //msdelay(750);
                            }
                            if (metricpath != null)
                                File.Delete(metricpath[0]);
                            //  currentmetricN = 0;
                            //return;
                            fileSystemWatcher5.EnableRaisingEvents = false;
                            // end 10-11-16 addt



                            //SetForegroundWindow(Handles.NebhWnd);
                            //Thread.Sleep(1000);
                            //PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                            //Thread.Sleep(1000);
                            //NebListenOn = false;
                            ////  File.Delete(metricpath[0]);
                            ///*
                            //// clientSocket.GetStream().Close();//added 5-17-12
                            ////  clientSocket.Client.Disconnect(true);//added 5-17-12
                            //clientSocket.Close();
                            //*/
                            //if (metricpath != null)
                            //    File.Delete(metricpath[0]);
                            //currentmetricN = 0;
                        }
                        GlobalVariables.FineRepeatDone = 0;
                        GlobalVariables.FineRepeatOn = false;
                        standby();

                    }

                }
                //end coarse v-curve and goes to rough focus point
                if ((vDone == 1) & (!GlobalVariables.Tempon) & (backlashDetermOn == false) & (GlobalVariables.FFenable != true))
                {
                    fileSystemWatcher1.EnableRaisingEvents = false;
                    fileSystemWatcher2.EnableRaisingEvents = false;
                    if (checkBox29.Checked == false)
                    {

                        gotopos(posMin - (int)numericUpDown40.Value * 4);


                        // gotopos(posMin - 3000);//  ***** added 1-4-13 **** take up backlash,  3000 works for geared stepper on fsq85, my not need as much for tsa120
                        Thread.Sleep(1000);
                    }
                    gotopos(Convert.ToInt32(posMin));
                    vDone = 0;//****  added 11-20-14

                    standby();

                }
                //reset for more v -curves for tempcal
                if ((GlobalVariables.Tempon) & (vDone != 1))
                {
                    repeatDone = 0;
                    repeatTotal = 0;
                    fileSystemWatcher1.EnableRaisingEvents = true;
                }
                //tempcal done
                if ((GlobalVariables.Tempon) & (vDone == 1))
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
                    if (checkBox29.Checked == false)
                    {
                        gotopos(Convert.ToInt32(finegoto - 100));//take up backlash                            
                        Thread.Sleep(2000);
                        gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 4));
                        Thread.Sleep(1000);
                        gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 3));
                        Thread.Sleep(1000);
                        gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 2));
                        Thread.Sleep(1000);
                    }
                    gotopos(Convert.ToInt32(finegoto - (int)numericUpDown8.Value));
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt32(posMin));

                    repeatProgress = 0;
                    vcurveenable = 0;
                    Array.Clear(list, 0, arraysize2);
                    Array.Clear(listMax, 0, arraysize2);
                    _hFRarraymin = 999;
                    maxarrayMax = 1;
                    tempcal();

                }
                vcurveenable = 0;//added 9-12
                                 //   return;  //added 11-20-14
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
                    //string path = textBox11.Text.ToString();
                    //string fullpath = path + @"\log.txt";

                    //StreamWriter log;
                    //if (!File.Exists(fullpath))
                    //{
                    //    log = new StreamWriter(fullpath);
                    //}
                    //else
                    //{
                    //    log = File.AppendText(fullpath);
                    //}
                    if (backlashCount == 2)
                    {
                        //log.WriteLine(DateTime.Now);
                        //log.WriteLine("Backlash N" + "\tPosOUT" + "\tPosIN" + "\tBacklash" + "\tAvg");
                        FileLog2(DateTime.Now.ToString());
                        FileLog2("Backlash N" + "\tPosOUT" + "\tPosIN" + "\tBacklash" + "\tAvg");
                    }
                    //  log.WriteLine("         " + backlashCount.ToString() + "\t  " + backlashPosOUT.ToString() + "\t " + backlashPosIN.ToString() + "\t   " + backlash.ToString() + "\t\t" + backlashAvg.ToString());
                    FileLog2("         " + backlashCount.ToString() + "\t  " + backlashPosOUT.ToString() + "\t " + backlashPosIN.ToString() + "\t   " + backlash.ToString() + "\t\t" + backlashAvg.ToString());
                    if (backlashCount == backlashN)
                    {
                        //  log.WriteLine("Avg Backlash: " + backlashAvg.ToString());
                        FileLog2("Avg Backlash: " + backlashAvg.ToString());
                        textBox8.Text = "Avg " + backlashAvg.ToString();

                    }
                    //  log.Close();

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
                _hFRarraymin = 999;
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

        // }

        private void chart1_Click(object sender, EventArgs e)
        {
            chart1.Series[1].Points.AddXY(Convert.ToDouble(min), Convert.ToDouble(count));

        }


        //added to sync for manual changes
        void sync()
        {
            try
            {


                textBox1.Text = focuser.Position.ToString();
                count = focuser.Position;


                /*
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
                 */
            }


            catch (Exception e)
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
            if (GlobalVariables.Tempon)
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
            //   port.DiscardInBuffer();
            GlobalVariables.FFenable = true;
            GlobalVariables.Tempon = false;
            vcurve();
        }

        void tempcal()
        {

            fileSystemWatcher1.EnableRaisingEvents = true;
            //    port.DiscardInBuffer();
            GlobalVariables.FFenable = true;
            GlobalVariables.Tempon = true;
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
        //goes to near focus point uses std dev routine, takes 'Course N' exposures, calculates best focus and goes there
        public void gotoFocus()
        {
            try
            {
                FileLog2("gotoFocus");
                //Data d = new Data();
                GetAvg();
                FillData();
                //toolStripStatusLabel1.Text = "Taking up Backlash"; // remd 10-24-16 too brief until changed by gotopos
                //  this.Refresh(); // remd 10-24-16
                if (radioButton2.Checked == true)//using upslope
                {//this was original 
                    /*
                  //  gotopos(Convert.ToInt32(posMin + 50));
                    focuser.Move(posMin + 50);
                    Thread.Sleep(2000);
                     //  gotopos(Convert.ToInt32(posMin + 20);
                    focuser.Move(posMin + (int)numericUpDown39.Value);
                    count = focuser.Position;
                    Thread.Sleep(3000);
*/

                    //*****  backlash stuff changed  3-4-14 *****
                    if (textBox8.Text == "")
                        gotopos(posMin + ((int)numericUpDown39.Value * 5));
                    //  focuser.Move(posMin + ((int)numericUpDown39.Value * 5));//if no backlash determined move 5* farther than measure position

                    else//if not using backlash compensation....still need to make sure backlash it taken up 
                        gotopos(posMin + Convert.ToInt32(textBox8.Text) + (int)numericUpDown39.Value * 2);
                    //   focuser.Move(posMin + Convert.ToInt32(textBox8.Text) + (int)numericUpDown39.Value * 2);//ensure well beyond backlash + measure position
                    Thread.Sleep(100);
                    count = focuser.Position;
                    textBox1.Text = focuser.Position.ToString();
                    Log("Taking up backlash");
                    FileLog2("Move to " + focuser.Position.ToString());
                    if (!ContinuousHoldOn)
                        focuser.Halt();
                    Thread.Sleep(1000);//was 3000
                    gotopos(posMin + (int)numericUpDown39.Value);
                    //   focuser.Move(posMin + (int)numericUpDown39.Value);//goto measure position
                    Thread.Sleep(100);
                    count = focuser.Position;
                    textBox1.Text = focuser.Position.ToString();
                    if (!ContinuousHoldOn)
                        focuser.Halt();
                    Thread.Sleep(2000);
                    FileLog2("Move to: " + focuser.Position.ToString());
                }
                if (radioButton3.Checked == true)//using downslope added 11-21
                {
                    if (textBox8.Text == "")
                        gotopos(posMin - ((int)numericUpDown39.Value * 5));
                    //  focuser.Move(posMin - ((int)numericUpDown39.Value * 5));
                    else//if not using backlash compensation....still need to make sure backlash it taken up 
                        gotopos(posMin - Convert.ToInt32(textBox8.Text) - (int)numericUpDown39.Value * 2);
                    // focuser.Move(posMin - Convert.ToInt32(textBox8.Text) - (int)numericUpDown39.Value * 2);//ensure well beyond backlash + measure position
                    Thread.Sleep(100);
                    count = focuser.Position;
                    textBox1.Text = focuser.Position.ToString();
                    Log("Taking up backlash");
                    FileLog2("Move to: " + focuser.Position.ToString());
                    if (!ContinuousHoldOn)
                        focuser.Halt();
                    Thread.Sleep(1000);
                    gotopos(posMin - (int)numericUpDown39.Value);
                    // focuser.Move(posMin - (int)numericUpDown39.Value);
                    Thread.Sleep(100);

                    count = focuser.Position;
                    textBox1.Text = focuser.Position.ToString();
                    if (!ContinuousHoldOn)
                        focuser.Halt();
                    Thread.Sleep(2000);
                    FileLog2("Move to: " +focuser.Position.ToString());


                    /*
                                        gotopos(Convert.ToInt32(posMin - 50));
                                        Thread.Sleep(3000);
                                        gotopos(Convert.ToInt32(posMin - 20));//was 30 8-20-13
                                        Thread.Sleep(3000);
                     */
                }
                //    port.DiscardInBuffer();
                //    port.DiscardOutBuffer();
                fileSystemWatcher2.EnableRaisingEvents = false;
                fileSystemWatcher5.EnableRaisingEvents = false; //added to test metricHFR
                fileSystemWatcher3.EnableRaisingEvents = true;
                fileSystemWatcher1.EnableRaisingEvents = false;
                fileSystemWatcher4.EnableRaisingEvents = false;
                focusSampleComplete = false; // 10-13-16
                _gotoFocusOn = true;//~line 440
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
                    NebListenStart(Handles.NebhWnd, SocketPort);
                    //   }


                    MetricCapture();
                }
                else
                    fileSystemWatcher3.Filter = "*.bmp";


            }
            catch (Exception e)
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
        private void UpdateData()
        {
            try
            {
                FileLog2("Update Data");
                if (_equip == null)
                {
                    MessageBox.Show("Must select equipment first", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                // Data d = new Data();
                //   GetAvg();    
                //   FillData();
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
            catch (Exception ex)
            {
                Log("Update Error - Line 1817" + ex.ToString());
            }
        }



        //update button
        private void button13_Click(object sender, EventArgs e)
        {
            //Data d = new Data();
            UpdateData();
            EquipRefresh();
            //try
            //{
            //    if (_equip == null)
            //    {
            //        MessageBox.Show("Must select equipment first", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        return;
            //    }
            //   // Data d = new Data();
            //    GetAvg();
            //    FillData();
            //    EquipRefresh();
            //    //try adding std dev and display in textbox16
            //    // Std Dev UP
            //    if (dataGridView1.RowCount > 2)
            //    {
            //        List<double> StdDevCalcUP = new List<double>();
            //        for (int i = 0; i < (dataGridView1.Rows.Count - 1); i++)
            //        {
            //            StdDevCalcUP.Add(Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value));
            //        }

            //        double standardDeviationUP = CalculateSD(StdDevCalcUP);
            //        double sdUP = Math.Round(standardDeviationUP, 5);
            //        textBox16.Text = sdUP.ToString();
            //    }
            //    /*  make sure there is enough data   *****remd 6-28.  not needed annoying  
            //    if (dataGridView1.RowCount <= 2)
            //    {
            //        DialogResult result;
            //        result = MessageBox.Show("Need More Data for Calculation", "scopefocus",
            //                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        if (result == DialogResult.OK)
            //        {
            //            return;
            //        }
            //    }
            //    */

            //    //Std Dev DWN
            //    if (dataGridView1.RowCount > 2)
            //    {
            //        List<double> StdDevCalcDWN = new List<double>();
            //        for (int i = 0; i < (dataGridView1.Rows.Count - 1); i++)
            //        {
            //            StdDevCalcDWN.Add(Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value));
            //        }

            //        double standardDeviationDWN = CalculateSD(StdDevCalcDWN);
            //        double sdDWN = Math.Round(standardDeviationDWN, 5);
            //        textBox14.Text = sdDWN.ToString();
            //    }
            //    /*    ****remd 6-29 not needed
            //    if (dataGridView1.RowCount <= 2)
            //    {
            //        DialogResult result;
            //        result = MessageBox.Show("Need More Data for Calculation", "scopefocus",
            //                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        if (result == DialogResult.OK)
            //        {
            //            return;
            //        }
            //    }
            //     */
            //  //  EquipRefresh();
            //}
            //catch (Exception ex)
            //{
            //    Log("Update Error - Line 1817" + ex.ToString());
            //}
        }

        //private void GetHandles()
        //{
        //    try
        //    {

        //        Callback myCallBack = new Callback(EnumChildGetValue);

        //        //    NebhWnd = FindWindow(null, "Nebulosity v3.0-a6");
        //        //    LoadScripthwnd = FindWindow(null, "Load script");
        //        //    Log("Load script " + LoadScripthwnd.ToString());
        //        //  SetForegroundWindow(NebhWnd);
        //        //   if (NebhWnd == 0)
        //        //    {
        //        //        MessageBox.Show("Please Start Calling Window Application");
        //        //    }


        //        EnumChildWindows(LoadScripthwnd, myCallBack, 0);
        //    }
        //    catch (Exception e)
        //    {
        //        Log("GetHandles Error" + e.ToString());
        //        Send("GetHandles Error" + e.ToString());
        //        FileLog("GetHandles Error" + e.ToString());

        //    }

        //}

        public string MainWindowName;
        //private bool FindHandlesDone = false;
        //public void FindHandles()
        //{
        //     //******5-30 this might be problem, maybe this should just be done once...not sure if 
        //    //finding the loadscript though. consider doing seperate for the 2 neb windows.  one 
        //    //with (spawned)
        //    try
        //    {

        //        MainWindowName = "Nebulosity";
        //        IntPtr hWnd2 = SearchForWindow("wxWindow", MainWindowName);
        //        Log("Neb Handle Found -- " + hWnd2.ToInt32());
        //        NebhWnd = hWnd2.ToInt32();

        //        if (NebhWnd != 0)
        //        {
        //            string NebVersion;
        //            //finds the Neb version (the number after the "v")
        //            StringBuilder sb = new StringBuilder(1024);
        //            SendMessage(NebhWnd, WM_GETTEXT, 1024, sb);
        //            //  GetWindowText(camera, sb, sb.Capacity);
        //            Log(sb.ToString());
        //            NebVersion = sb.ToString();
        //            int NebVposNumber = NebVersion.IndexOf("v");
        //            string NebVNumberAfter = NebVersion.Substring(NebVposNumber + 1, 1);
        //            NebVNumber = Convert.ToInt32(NebVNumberAfter);
        //        }





        //        if (checkBox14.Checked == false)//don't need for slave mode
        //        {
        //            IntPtr PHDhwnd2 = SearchForWindow("wxWindow", "PHD");
        //            Log("PHD Handle Found --  " + PHDhwnd2.ToInt32());
        //            PHDhwnd = PHDhwnd2.ToInt32();
        //        }
        //        // PHDhwnd = FindWindow(null, "PHD Guiding 1.13.0b  -  www.stark-labs.com (Log active)");
        //        LoadScripthwnd = FindWindow(null, "Load script");
        //        //   Log("PHD " + PHDhwnd.ToString());
        //        //   Log("Load script " + LoadScripthwnd.ToString());



        //    }

        //    catch (Exception e)
        //    {
        //        Log("FindHandles Error" + e.ToString());
        //        Send("FindHandles Error" + e.ToString());
        //        FileLog("FindHandles Error" + e.ToString());

        //    }

        //}
        //  double CalibrationTol;//may not need this variable...in equation uses textbox value and convert to doulg 
        //Loadhere

      



        private void MainWindow_Load_1(object sender, EventArgs e)
        {
            try
            {

                //5 lines below for path2 settings
                GlobalVariables.Path2 = WindowsFormsApplication1.Properties.Settings.Default.path.ToString();
                textBox11.Text = GlobalVariables.Path2.ToString();
                GlobalVariables.NebPath = WindowsFormsApplication1.Properties.Settings.Default.NebPath.ToString();
                textBox35.Text = GlobalVariables.NebPath.ToString();
                // send = NebPath + ScriptName;
                // textBox37.Text = send.ToString();
                toolStripStatusLabel5.Text = "Path: " + GlobalVariables.Path2.ToString();
                //     CalibrationTol = Convert.ToDouble(textBox59.Text);
                textBox24.Text = PublishVersion.ToString();//version number on about tab
                //************ try addition for new equipcomboxes
                if (WindowsFormsApplication1.Properties.Settings.Default.ComboItems7 == null)
                {
                    WindowsFormsApplication1.Properties.Settings.Default.ComboItems7 = new System.Collections.Specialized.StringCollection();
                }

                foreach (string item in WindowsFormsApplication1.Properties.Settings.Default.ComboItems7)
                {
                    comboBox7.Items.Add(item);
                    if (comboBox7.Items.Count == 1)
                        comboBox7.SelectedIndex = 0;
                }
                if (WindowsFormsApplication1.Properties.Settings.Default.ComboItems8 == null)
                {
                    WindowsFormsApplication1.Properties.Settings.Default.ComboItems8 = new System.Collections.Specialized.StringCollection();
                }
                /*
                foreach (string item in WindowsFormsApplication1.Properties.Settings.Default.ComboItems8)
                {
                    comboBox8.Items.Add(item);
                }
                 */
                backlash = WindowsFormsApplication1.Properties.Settings.Default.backlash;
                textBox8.Text = backlash.ToString();
                Handles H = new Handles();
                Callback myCallBack = new Callback(H.EnumChildGetValue);

                H.FindHandles();
                //   int hWnd;


                if (Handles.PHDhwnd == 0)//this can really just serve as reminded that PHD is needed for some functions...not necessary for focusing.  
                {
                    DialogResult result1;
                    result1 = MessageBox.Show("PHD Not Found - Open and 'Retry', 'Ignore' or 'Abort' to Close", "scopefocus",
                         MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);

                    if (result1 == DialogResult.Ignore)
                        this.Refresh();
                    if (result1 == DialogResult.Retry)
                    {
                        H.FindHandles();
                        this.Refresh();
                    }
                    if (result1 == DialogResult.Abort)
                        System.Environment.Exit(0);
                }
                else
                    Log("PHD version " + Handles.PHDVNumber.ToString());


                if (Handles.NebhWnd == 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Nebulosity Not Found - Open or Close & Restart and hit 'Retry' or 'Ignore' to continue",
                       "scopefocus", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);//change so ok moves focus
                    if (result == DialogResult.Ignore)
                        this.Refresh();
                    if (result == DialogResult.Retry)
                    {
                        H.FindHandles();
                        this.Refresh();
                    }
                    if (result == DialogResult.Abort)
                        System.Environment.Exit(0);

                }
                else
                {

                    EnumChildWindows(Handles.NebhWnd, myCallBack, 0);
                    Log("Nebulosity Version " + Handles.NebVNumber.ToString());
                }
              
                ResizeNebWindow();
              
                    


                if (Handles.NebhWnd != 0)
                    FindNebCamera();

                radioButton8.Checked = true;
                textBox27.Text = user.ToString();
                textBox13.Text = server.ToString();
                textBox28.Text = pswd.ToString();
                textBox36.Text = to.ToString();
                fileSystemWatcher1.Path = GlobalVariables.Path2.ToString();
                fileSystemWatcher2.Path = GlobalVariables.Path2.ToString();
                fileSystemWatcher5.Path = GlobalVariables.Path2.ToString();//added to test metricHFR
                fileSystemWatcher3.Path = GlobalVariables.Path2.ToString();
                fileSystemWatcher4.Path = GlobalVariables.Path2.ToString();
                //   travel = WindowsFormsApplication1.Properties.Settings.Default.travel;
                textBox2.Text = travel.ToString();
                textBox60.Text = WindowsFormsApplication1.Properties.Settings.Default.sigma.ToString();
                textBox49.Text = WindowsFormsApplication1.Properties.Settings.Default.DwnSz.ToString();
                textBox61.Text = WindowsFormsApplication1.Properties.Settings.Default.Low.ToString();
                textBox62.Text = WindowsFormsApplication1.Properties.Settings.Default.High.ToString();
                textBox38.Text = WindowsFormsApplication1.Properties.Settings.Default.GoalADU;
                //   camera = WindowsFormsApplication1.Properties.Settings.Default.camera;
                //   textBox22.Text = camera;
                posMin = WindowsFormsApplication1.Properties.Settings.Default.focuspos;
                textBox4.Text = posMin.ToString();

                numericUpDown40.Value = WindowsFormsApplication1.Properties.Settings.Default.stepsize;
                numericUpDown2.Value = WindowsFormsApplication1.Properties.Settings.Default.stepsize * 4;//coarse V
                numericUpDown8.Value = WindowsFormsApplication1.Properties.Settings.Default.stepsize;//fine V
                numericUpDown39.Value = WindowsFormsApplication1.Properties.Settings.Default.stepsize * 3;//gotofocus sample pos
                //first time load both are false which if not set gives tab change error.
                if (WindowsFormsApplication1.Properties.Settings.Default.None == false && WindowsFormsApplication1.Properties.Settings.Default.AO == false)
                    radioButton8.Checked = true;
                else
                    radioButton8.Checked = WindowsFormsApplication1.Properties.Settings.Default.None;

                radioButton9.Checked = WindowsFormsApplication1.Properties.Settings.Default.Extender;
                radioButton10.Checked = WindowsFormsApplication1.Properties.Settings.Default.Reducer;
                radioButton4.Checked = WindowsFormsApplication1.Properties.Settings.Default.AO;

                setarraysize();

                //  Data d = new Data();
                //  EquipRefresh();
                FillData();

                toolStripStatusLabel1.Text = "Startup Complete";
                textBox44.Text = trackBar1.Value.ToString();
                toolStripStatusLabel4.Text = "No Filter";
                toolStripStatusLabel2.Text = "Idle";
                toolStripStatusLabel3.Text = "Equip";
                // if (!Directory.Exists(@"C:\cygwin\home\astro"))
                if (!Directory.Exists((Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\cygwin_ansvr")))
                {
                    radioButton5_astrometry.Checked = true;
                    radioButton_local.Enabled = false;
                }

                GlobalVariables.LocalPlateSolve = radioButton_local.Checked;

                EquipRefresh();

                // auto connect to PHD
                string result2 = ph.Connect();
                if (result2 != "Failed")
                {
                    textBoxPHDstatus.Text = "Connect Ver " + ph.Connect();
                    button36.BackColor = Color.Lime;
                    button36.Text = "DisCnct";
                    Log("Connected to PHD2 version " + result2);
                    FileLog("Connected to PHD2 version " + result2);


                }


            }
            catch (Exception ex)
            {
                Log("MainWindow_Load Error" + ex.ToString());
            }

        }

        private void ButtonDisable_Click_1(object sender, EventArgs e)
        {
            ClosePrep();
        }


        //Abort
        private void button4_Click_2(object sender, EventArgs e)
        {
            GlobalVariables.FineRepeatOn = false;//added 1-7-12
            GlobalVariables.FineRepeatDone = 0;
            textBox20.Enabled = true;
            textBox18.Enabled = true;
            Log("Aborted");
            FileLog2("Aborted");
            standby();
        }
        //reverse
        private void button2_Click_1(object sender, EventArgs e)
        {
            int step = (int)numericUpDown2.Value;
            //    watchforOpenPort();
            //if ((count + step) <= travel)
            //{

            gotopos(Convert.ToInt32(count + step));

            //textBox1.Text = count.ToString();
            //textBox4.Text = posMin.ToString();
            //positionbar();
            //  }
            //if (count + step > travel)
            //    MessageBox.Show("goto exceeds travel");
        }
        //v-curve (coarse)
        private void button6_Click_1(object sender, EventArgs e)
        {
            //if (GlobalVariables.Nebcamera == "No camera")
            //    NoCameraSelected();
            setarraysize();
            //    watchforOpenPort();
            //   if (portopen == 1)
            //  {
            //     port.DiscardOutBuffer();
            //    port.DiscardInBuffer();
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
            GlobalVariables.Tempon = false;
            GlobalVariables.FFenable = false;
            posMin = count;
            vv = 0;
            tempsum = 0;
            maxMax = 0;
            //  roughvdone = true;
           
            if (checkBox22.Checked == true)
                fileSystemWatcher5.EnableRaisingEvents = true;//added to test metricHFR
            else // added 12-29-16 
                fileSystemWatcher2.EnableRaisingEvents = true;
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
            //  }
        }
        //fine V

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {

                if (checkBox22.Checked == true) // added 11-3-16 to remove separate metric v button
                    currentmetricN = 0;
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
                //  button3.PerformClick();
                /*
                if (clientSocket.Connected == false)
                {
                    //clientSocket.Connect("127.0.0.1", SocketPort);//connects to neb
                    NebListenStart(NebhWnd, SocketPort);
                }
                 */



                //if (GlobalVariables.Nebcamera == "No camera")
                //    NoCameraSelected();
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
                //      watchforOpenPort();
                //      if (portopen == 1)
                //      {
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
                    enteredMaxHFR = Convert.ToInt32(textBox20.Text.ToString());
                    enteredMinHFR = Convert.ToInt32(textBox18.Text.ToString());
                    textBox20.Enabled = false; //don't allow changes once done
                    textBox18.Enabled = false;
                }
                backlashDetermOn = false;
                //    port.DiscardInBuffer();
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
                if (checkBox29.Checked == false) // not using backlash compensation
                {
                    Log("Clearing backlash...");//added 1-7-12
                                                //  gotopos(Convert.ToInt32(finegoto - 3000));//take up backlash     
                                                //   Thread.Sleep(1000);

                    if (textBox8.Text == "")
                        gotopos(finegoto - ((int)numericUpDown2.Value * 2)); // 12-29-16 change from * 4 to *2
                    //   focuser.Move(finegoto - ((int)numericUpDown2.Value * 4));//if no backlash determined move 4* course step size to get beyond backalsh
                    else//if not using backlash compensation....still need to make sure backlash it taken up 

                        //remd next 4 on 12-29-16  no need to go this far past.....
                    //    gotopos(finegoto - (Convert.ToInt32(textBox8.Text) + (int)numericUpDown8.Value * 4));
                    ////   focuser.Move(finegoto - (Convert.ToInt32(textBox8.Text) + (int)numericUpDown8.Value * 4));//ensure well beyond backlash + measure position

                    //Thread.Sleep(1000);
                    //gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 3));
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 2));
                    Thread.Sleep(1000);
                }
                gotopos(Convert.ToInt32(finegoto - (int)numericUpDown8.Value));
                Thread.Sleep(1000);
                gotopos(Convert.ToInt32(finegoto));
                Thread.Sleep(1000);
                GlobalVariables.FFenable = true;
                repeatProgress = 0;
                repeatDone = 0;
                vProgress = 0;
                vDone = 0;
                repeatTotal = 0;
                vN = 0;
                sum = 0;
                min = 500;
                _hFRarraymin = 999;
                GlobalVariables.Tempon = false;
                sumMax = 0;
                avgMax = 0;
                maxMax = 0;
                _apexHFR = 0;
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
                if (GlobalVariables.FineRepeatOn == false)//added 1-7-12
                {
                    GlobalVariables.FineRepeatDone = 0;
                }
               

                if (checkBox22.Checked) // added 11-3-16
                {
                    fileSystemWatcher5.EnableRaisingEvents = true;//added to test metric HFR   // moved from above the if 12-29-16
                    MetricCapture();
                }
                else // added 12-29-16
                {
                    fileSystemWatcher2.EnableRaisingEvents = true;

                    fileSystemWatcher3.EnableRaisingEvents = false;
                    fileSystemWatcher1.EnableRaisingEvents = false;
                }
            }

            catch (Exception ex)
            {
                Log("FineV Error - Line 2191" + ex.ToString());
                Send("FineV Error - Line 2191" + ex.ToString());
                FileLog("FineV Error - Line 2191" + ex.ToString());

            }
        }

        void fineVrepeatcounter()
        {
            if ((GlobalVariables.FineRepeatOn == true) & (GlobalVariables.FineVRepeat > 1))
            {
                GlobalVariables.FineRepeatDone++;
                GlobalVariables.FineVRepeat--;
            }
            if (GlobalVariables.FineRepeatOn == false)
            {
                GlobalVariables.FineVRepeat = (int)numericUpDown10.Value;
                if (GlobalVariables.FineVRepeat > 1)
                {
                    GlobalVariables.FineRepeatOn = true;
                }
            }
        }
        //fwd button
        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                //  watchforOpenPort();
                //   if (portopen == 1)
                //   {
                int step = (int)numericUpDown2.Value;
                //if ((count - step) < 0)
                //{
                //    DialogResult result1 = MessageBox.Show("Goto Exceeds Full In", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //    if (result1 == DialogResult.OK)
                //    {
                //        count = 0;
                //        return;
                //    }
                //}
                gotopos(Convert.ToInt32(count - step));
                //textBox1.Text = focuser.Position.ToString();
                //count = focuser.Position;
                //textBox4.Text = posMin.ToString();
                //positionbar();
                //   }
            }
            catch (Exception ex)
            {
                Log("Fwd Button Error" + ex.ToString());
            }
        }

        private void button9_Click_1(object sender, EventArgs e)
        {

            sync();
            // count = syncval;
        }

        private void button7_Click_1(object sender, EventArgs e)
        {

            try
            {


                /*
                if (usingASCOMFocus)
                {
                    int pos = 500;
                 //   MessageBox.Show("test" + focuser.Name);
                    focuser = new ASCOM.DriverAccess.Focuser(devId);
                    focuser.Move(pos);//doesn't work!!!
                    textBox1.Text = focuser.Position.ToString();
                }
                */

                //   if (portopen == 1)
                //   {
                //  fileSystemWatcher2.EnableRaisingEvents = false;
                //    fileSystemWatcher5.EnableRaisingEvents = false; // added to test metricHFR
                //   fileSystemWatcher3.EnableRaisingEvents = false;
                //    int go2 = (int)numericUpDown6.Value;
                /*
                    if (go2 == 1)
                    {
                        MessageBox.Show("The goto position 1 is reservered for filter advance", "scopefocus");
                        return;
                    }
                    if (go2 == 2)
                    {
                        MessageBox.Show("The goto position 2 is reservered for filter sync", "scopefocus");
                        return;
                    }
                    if (go2 == 3)
                    {
                        MessageBox.Show("The goto position 3 is reservered for flat panel toggle", "scopefocus");
                        return;
                    }  
                 */
                //  gotopos(Convert.ToInt32(go2));
                gotopos((int)(numericUpDown6).Value);
                Thread.Sleep(20);
                //  }
            }
            catch (Exception ex)
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
            //   Data d = new Data();
            GetAvg();
            if (checkBox22.Checked == false)
                FillData();
            setarraysize();
            // remd 10-11-16
            //if (_enteredSlopeDWN == 0 || _enteredSlopeUP == 0 || _enteredPID == 0)
            //{
            //    MessageBox.Show("Focus data is blank -- check equipment selelction or save v-curve data then retry.  GotoFocus Aborted", "scopefocus");
            //    return;
            //}
            gotoFocus();
        }


        private void MainWindow_Shown(object sender, EventArgs e)
        {

            //if (portautoselect == true) //not sure if this is needed after adding slave. 
            //{
            //    comboBox1.SelectedIndex = 0;
            //    GlobalVariables.Portselected = comboBox1.SelectedItem.ToString();
            //    button8.PerformClick();
            //    if (port != null)
            //    {
            //        button8.BackColor = System.Drawing.Color.Lime;
            //        this.button8.Text = "Connected";
            //        playsound();
            //    }
            //}
            //else
            //{
            //    comboBox1.Focus();
            //}
        }


        public void button8_Click_2(object sender, EventArgs e)
        {
            try
            {
                if (button8.BackColor != Color.Lime)
                    FocusChooser();
                else
                {
                   
                    focuser.Connected = false;
                    button8.BackColor = Color.WhiteSmoke;
                    focuser.Dispose();
                    Log(devId2 + " disconnected");
                    devId2 = "";
                }
            }
            catch (Exception ex)
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

            //  *** remd 11-3-16   rem for sim

            if (GlobalVariables.FFenable)
            {

                if (vProgress == numericUpDown3.Value)
                {
                    fileSystemWatcher2.EnableRaisingEvents = false;
                    fileSystemWatcher5.EnableRaisingEvents = false;
                }

            }
            else
            {
                if (vProgress == numericUpDown5.Value)

                {

                    fileSystemWatcher2.EnableRaisingEvents = false;
                    fileSystemWatcher5.EnableRaisingEvents = false;

                }

            }
            vcurve();
        }

        private void textBox11_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            GlobalVariables.Path2 = folderBrowserDialog1.SelectedPath.ToString();
            textBox11.Text = GlobalVariables.Path2.ToString();
            toolStripStatusLabel5.Text = GlobalVariables.Path2.ToString();

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
                    WindowsFormsApplication1.Properties.Settings.Default.travel = 380000;
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
            catch (Exception ex)
            {
                Log("Restore defaults - Error" + ex.ToString());
            }
        }

        //private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        //{

        //    GlobalVariables.Portselected = comboBox1.SelectedItem.ToString();
        //}



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
            ClosePrep();
        }
        private void ClosePrep()
        {
            // PHD2comm.PHD2Connected = false;
            if (PHD2comm.PHD2Connected)
                ph.DisConnect();
            if (timer2.Enabled)
                timer2.Enabled = false;
            if (!string.IsNullOrEmpty(devId))
            {
                if (scope.Connected == true)
                {
                    scope.Connected = false;
                    Thread.Sleep(100);
                    scope.Dispose();
                }
            }
            if (!string.IsNullOrEmpty(Filter.DevId3))
            {
                if (Filter.filterWheel.Connected == true)
                {
                    Filter.filterWheel.Connected = false;
                    Thread.Sleep(100);
                    Filter.filterWheel.Dispose();
                }
            }
            if (!string.IsNullOrEmpty(devId2))
            {
                if (focuser.Connected == true)
                {
                    focuser.Connected = false;
                    Thread.Sleep(100);
                    focuser.Dispose();
                }
            }
            if (!string.IsNullOrEmpty(devId4))
            {
                if (FlatFlap.Connected == true)
                {
                    FlatFlap.Connected = false;
                    Thread.Sleep(100);
                    FlatFlap.Dispose();
                }
            }
            //  if (ContinuousHoldOn == true)
            //   button59.PerformClick();
            //*****add for new equip comboboxes
            WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Clear();
            foreach (string Item in comboBox7.Items)
            {
                WindowsFormsApplication1.Properties.Settings.Default.ComboItems7.Add(Item);
            }
            /*
            WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Clear();
            foreach (string Item in comboBox8.Items)
            {
                WindowsFormsApplication1.Properties.Settings.Default.ComboItems8.Add(Item);
            }
            //******end addt.
            */
            //    WindowsFormsApplication1.Properties.Settings.Default.port = portselected.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.backlash = backlash;
            WindowsFormsApplication1.Properties.Settings.Default.path = GlobalVariables.Path2.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.travel = travel;
            WindowsFormsApplication1.Properties.Settings.Default.NebPath = GlobalVariables.NebPath.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.user = this.user;
            WindowsFormsApplication1.Properties.Settings.Default.pswd = this.pswd;
            WindowsFormsApplication1.Properties.Settings.Default.server = this.server;
            WindowsFormsApplication1.Properties.Settings.Default.to = this.to;
            WindowsFormsApplication1.Properties.Settings.Default.sigma = Convert.ToInt32(textBox60.Text);
            WindowsFormsApplication1.Properties.Settings.Default.DwnSz = Convert.ToInt32(textBox49.Text);
            WindowsFormsApplication1.Properties.Settings.Default.Low = Convert.ToDouble(textBox61.Text);
            WindowsFormsApplication1.Properties.Settings.Default.High = Convert.ToDouble(textBox62.Text);
            WindowsFormsApplication1.Properties.Settings.Default.AO = radioButton4.Checked;
            WindowsFormsApplication1.Properties.Settings.Default.Extender = radioButton9.Checked;
            WindowsFormsApplication1.Properties.Settings.Default.None = radioButton8.Checked;
            WindowsFormsApplication1.Properties.Settings.Default.Reducer = radioButton10.Checked;
            WindowsFormsApplication1.Properties.Settings.Default.focuspos = posMin;
            WindowsFormsApplication1.Properties.Settings.Default.GoalADU = textBox38.Text;
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
            WindowsFormsApplication1.Properties.Settings.Default.stepsize = (int)numericUpDown40.Value;
            WindowsFormsApplication1.Properties.Settings.Default.Save();
            closing = 1;
            standby();

          
            //if (port == null)
            //{
            //    System.Environment.Exit(0);
            //}
            //else
            //{


            //   if (focuser.Position != 0) //**************** rem'd goto 0 1-4-13 ************ since arduino stores in EPROM ok to leave out????
            //    gotopos(0);


            //   port.Close();
            //   }
            //if (phdsocket != null)
            //    phdsocket.Close();
            if (clientSocket != null)
            {
                //clientSocket.GetStream().Close();//added 5-17-12
                //  clientSocket.Client.Disconnect(false);//added 5-17-12
                clientSocket.Close();
            }
            //**************added 4-12-13*********

            //  if (port == null)
            //  {
            //   if (focuser.Position == 0)
            System.Environment.Exit(0);
            //   else
            //       return;
            //  }
        }

        private bool checkboxChanged = false;//allows for only 1 time fill of all checkboxes at 
        //startup, once on is changed--tab change no longer fills the boxes
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                EquipRefresh();

            }
            catch (Exception ex)
            {
                Log("TabContorl Changed - Error" + ex.ToString());
            }
        }
        private void EquipRefresh()
        {
            if (checkBox22.Checked == false)
            {
                IsSelectedRow = false;
                dataGridView1.ClearSelection();
            }
            // GlobalVariables.EquipPrefix = comboBox7.Text + "_" + comboBox8.Text;
            GlobalVariables.EquipPrefix = comboBox7.Text;
            /*
            if (comboBox8.Text == "NB")
            {
                if (checkboxChanged == false)
                {
                    StdNB();
                    GlobalVariables.EquipPrefix = comboBox7.Text + "_" + "LRGB";//always start w/ lum for NB
                }
            }
             */
            if (radioButton9.Checked == true)
            {
                _equip = GlobalVariables.EquipPrefix + "_E";
                if (radioButton4.Checked == true)
                    _equip = GlobalVariables.EquipPrefix + "_E_AO";
            }

            else
            {
                if (radioButton10.Checked == true)// 9 and 10 toggle
                {
                    _equip = GlobalVariables.EquipPrefix + " _R";
                    if (radioButton4.Checked == true)
                        _equip = GlobalVariables.EquipPrefix + "_R_AO";
                }
                else
                {
                    if (radioButton4.Checked == true)
                        _equip = GlobalVariables.EquipPrefix + " _AO";
                    if (radioButton8.Checked == true)
                        _equip = GlobalVariables.EquipPrefix;
                }
            }
            /*
            if (radioButton4.Checked == true)
            {
               radioButton10.Checked = false;
               radioButton9.Checked = false;
               radioButton4.Checked = false;
            _equip = GlobalVariables.EquipPrefix;
            }
             */
            /*
            if (comboBox8.Text == "LRGB")
            {
                if (checkboxChanged == false)
                    StdLRGB();
            }
             */
            //*********** 3-3-14 ***********  removed to allow metric to use single star v-curves  ************************
            //  %%%%%%%%% unrem'd 10-10-16 to redo full frame metric, allow for selecting v-curve data   %%%%%%%%%      
            if (checkBox22.Checked == true)
                _equip = _equip + "_Metric";
            //**************************************************************************************************************
            //   textBox13.Text = equip.ToString();
            toolStripStatusLabel3.Text = _equip.ToString();
            //next line filters gridview by equip
            ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";


            //  ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";
            //   Data d = new Data();
            //   ((DataTable)d.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";


            if (_equip == "")
            {
                MessageBox.Show("Must Select Equipment First", "scopefocus");
                tabControl1.SelectedTab = tabPage2;//page 2 is setup (1st tab)
            }
            /*
            if ((tabControl1.SelectedTab == tabPage5) & (GlobalVariables.Portselected2 == null))
            {
                if (comboBox6.Items.Count == 1)
                {
                    comboBox6.SelectedIndex = 0;
                    GlobalVariables.Portselected2 = comboBox6.SelectedItem.ToString();
                    

                }
            }
             */
            // button13.PerformClick();
            UpdateData();  // remd 10-14-16

            //below give a focus position starting point based on v-curve data.

            //************ remd 9-17-14   changed to remember last focuspos setting and load upon opening.  
            //if (startup == true)
            //{
            //    using (SqlCeConnection con = new SqlCeConnection(conString))
            //    {
            //        con.Open();
            //        using (SqlCeCommand com4 = new SqlCeCommand("SELECT AVG(FocusPos) FROM table1 WHERE Equip = @equip", con))
            //        {
            //            com4.Parameters.AddWithValue("@equip", _equip);
            //            SqlCeDataReader reader4 = com4.ExecuteReader();
            //            while (reader4.Read())
            //            {
            //                if (!reader4.IsDBNull(0))
            //                {
            //                    int numb7 = reader4.GetInt32(0);
            //                    posMin = numb7;
            //                    textBox4.Text = numb7.ToString();//sets focus position to avg from data
            //                    numericUpDown6.Value = numb7;//sets goto value to avg focus position
            //                }
            //            }
            //            reader4.Close();
            //        }
            //        con.Close();
            //    }
            //    startup = false;
            //}
        }
        //deletes all SQL data for this Equip
        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                if (_equip == null)
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

                            com.Parameters.Add(new SqlCeParameter("@equip", _equip));
                            com.ExecuteNonQuery();
                        }
                        con.Close();
                    }
                    // Data d = new Data();
                    FillData();
                }
            }
            catch (Exception ex)
            {
                Log("Error Deleting SQL Data" + ex.ToString());
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
            catch (Exception ex)
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
            if (textBox20.Text == "")
                return;
            enteredMaxHFR = Convert.ToInt32(textBox20.Text.ToString());
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
            if (textBox18.Text == "")
                return;
            enteredMinHFR = Convert.ToInt32(textBox18.Text.ToString());
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
        /*
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
         */
        //these pass TB values to data class
        //public string TB3
        //{
        //    get { return textBox3.Text; }
        //    set { textBox3.Text = value; }
        //}
        //public string TB14
        //{
        //    get { return textBox14.Text; }
        //    set { textBox14.Text = value; }
        //}
        //public string TB16
        //{
        //    get { return textBox16.Text; }
        //    set { textBox16.Text = value; }
        //}
        //public string TB34
        //{
        //    get { return textBox34.Text; }
        //    set { textBox34.Text = value; }
        //}    
        //public string TB10
        //{
        //    get { return textBox10.Text; }
        //    set { textBox10.Text = value; }
        //}
        //public string TB15
        //{
        //    get { return textBox15.Text; }
        //    set { textBox15.Text = value; }
        //}

        //will need for Filter class
        //public string TS3
        //{
        //    get { return toolStripStatusLabel3.Text; }
        //}




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
                GlobalVariables.Tempon = false;
                tempsum = 0;
                vv = 0;
                templog = 0;
                backlashDetermOn = true;
                radioButton1.Checked = false;//calculations off
                GlobalVariables.FFenable = true;
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
                if (checkBox29.Checked == false)
                {
                    gotopos(Convert.ToInt32(finegoto - 100));//take up backlash  
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 4));
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 3));
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 2));
                    Thread.Sleep(1000);
                }
                gotopos(Convert.ToInt32(finegoto - (int)numericUpDown8.Value));
                Thread.Sleep(1000);
                gotopos(Convert.ToInt32(finegoto));
                backlashOUT = true;//identifies current v curve direction is going out(reverse)
                fileSystemWatcher2.EnableRaisingEvents = false;
                fileSystemWatcher5.EnableRaisingEvents = false;//added to test metricHFR
                fileSystemWatcher3.EnableRaisingEvents = false;

                fileSystemWatcher1.EnableRaisingEvents = true;
            }
            catch (Exception ex)
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
                GlobalVariables.Tempon = true;
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
                if (checkBox29.Checked == false)
                {
                    gotopos(Convert.ToInt32(finegoto - 100));//take up backlash  
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 4));
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 3));
                    Thread.Sleep(1000);
                    gotopos(Convert.ToInt32(finegoto - ((int)numericUpDown8.Value) * 2));
                    Thread.Sleep(1000);
                }
                gotopos(Convert.ToInt32(finegoto - (int)numericUpDown8.Value));
                Thread.Sleep(1000);
                fileSystemWatcher2.EnableRaisingEvents = false;
                fileSystemWatcher5.EnableRaisingEvents = false; //added to test metricHFR
                fileSystemWatcher3.EnableRaisingEvents = false;
                fileSystemWatcher1.EnableRaisingEvents = true;
                //  tempcal();
            }
            catch (Exception ex)
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
                int selectedrowcount;
                string selectedcell;
                //   selectedrowcount = dataGridView1.SelectedCells.Count;
                selectedrowcount = dataGridView1.SelectedRows.Count;
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
                    // Data d = new Data();
                    dataGridView1.ClearSelection();
                    int nRowIndex = dataGridView1.Rows.Count - 2;
                    if (nRowIndex > 0)
                        dataGridView1.Rows[nRowIndex].Selected = true;
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
                        // Data d = new Data();
                        dataGridView1.ClearSelection();
                        int nRowIndex = dataGridView1.Rows.Count - 2;
                        if (nRowIndex > 0)
                            dataGridView1.Rows[nRowIndex].Selected = true;
                        FillData();
                    }
                }
                IsSelectedRow = false;
            }
            catch (Exception ex)
            {
                Log("Delete Selected SQL data Error" + ex.ToString());
            }

        }

        //filteradvancehere
        int intFilterPos = 1; //Neb indexes 1 more than the ASCOM position 
        private void filteradvance()
        {
            // 11-8-16   TODO for clipboard and external filter
            if (checkBox31.Checked)
            {
                intFilterPos++;
                NebListenStart(Handles.NebhWnd, SocketPort);
                Thread.Sleep(1000);
                try
                {
                    if (!UseClipBoard.Checked)
                        serverStream = clientSocket.GetStream();
                }
                catch (IOException e)
                {
                    Log("Neb socket failure " + e.ToString());

                }
                byte[] outStream = System.Text.Encoding.ASCII.GetBytes("SetExtFilter " + intFilterPos + "\n");//for testing purposes use the Ext filter sim.  
                                                                                                              // byte[] outStream = System.Text.Encoding.ASCII.GetBytes("SetFilter " + intFilterPos + "\n");
                try
                {
                    serverStream.Write(outStream, 0, outStream.Length);
                    Log("Filter Postion set to " + intFilterPos.ToString());
                }
                catch
                {
                    MessageBox.Show("Error sending command", "scopefocus");
                    return;

                }
                //  serverStream = clientSocket.GetStream();
                if (!UseClipBoard.Checked)
                {
                    byte[] outStream2 = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                    serverStream.Write(outStream2, 0, outStream2.Length);
                    serverStream.Flush();
                    //toolStripStatusLabel1.Text = "Sequence Done";
                    //this.Refresh();
                    //if (DarksOn == true)
                    //    DarksOn = false;
                    //currentfilter = 1;
                    //DisplayCurrentFilter();
                    ////      currentfilter = 0;
                    //subCountCurrent = 0;
                    //filterCountCurrent = 0;
                    Thread.Sleep(3000);//prevent overlapping sounds
                    serverStream.Close();
                    SetForegroundWindow(Handles.NebhWnd);
                    Thread.Sleep(1000);
                    PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                    Thread.Sleep(3000);
                    //   NebListenOn = false;
                    // clientSocket.GetStream().Close();//added 5-17-12
                    //  clientSocket.Client.Disconnect(true);//added 5-17-12
                    clientSocket.Close();

                }

                return;
            }

            if (Filter.DevId3 != null)
            {

                string[] names = Filter.filterWheel.Names;
                if (Filter.filterWheel.Position == names.Length - 1)
                    Filter.filterWheel.Position = 0;
                else
                {

                    Filter.filterWheel.Position++;
                }
                while (Filter.filterWheel.Position == -1)
                    toolStripStatusLabel4.Text = "Filter Moving";

                //     DisplayCurrentFilter();//***need to account for 5 position versus prev 4  ****  2-28-14
                // }

                FileLog2("Filter moved to position: " + Filter.filterWheel.Position.ToString());
            }
            //focuser.Move(1);
            //     focuser.Action("FilterAdvance", "");
            //     Thread.Sleep(200);
            //      gotopos(1);//couldn't add commands to focus driver so use an essentially uneused position for this command
            //sync is 2 and flat panal toggle is 3

            //  try
            //  {
            //  focuser.CommandString("Filter", true);
            //   focuser.CommandBlind("Filter", true);
            //    focuser.CommandBool("filter", true);
            //************************************************this prob needs to go somewhere in filtersequence  ?????  **********            
            FlatCalcDone = false;
            if (SequenceRunning == true)
            {
                DisableUpDwn();
                // fileSystemWatcher1.EnableRaisingEvents = true;
            }
            //      toolStripStatusLabel1.Text = "Filter Moving";
            this.Refresh();
            //     string movedone;

            //*************************************************************************************************************             

            //  port.DiscardOutBuffer();
            //  port.DiscardInBuffer();
            //  Thread.Sleep(20);
            //port.Write("A");
            // filterMoving = true;
            // Thread.Sleep(50);

            //while (filterMoving == true)
            //{
            //    movedone = port.ReadExisting();
            //    if (movedone == "D")
            //    {

            //        filterMoving = false;

            //        toolStripStatusLabel1.Text = "Capturing";
            //        this.Refresh();
            //        //   toolStripStatusLabel1.Text = " ";
            //        //    toolStripStatusLabel1.Text = " ";

            //    }
            //}
            //Thread.Sleep(50);
            //port.DiscardInBuffer();

            //if (currentfilter == 0)//allows advance without count to set to starting pos. 
            //{
            //    return;
            //}
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
            //}
            //catch
            //{
            //    Log("Filter Advance Error - Make Sure Arduino Connected");
            //    Send("Filter Advance Error - Make Sure Arduino Connected");
            //    FileLog("Filter Advance Error - Make Sure Arduino Connected");

            //}
        }
        //displaycurrentfilterhere

        public void DisplayCurrentFilter()
        {
            try
            {
                FileLog2("display current filter");
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
                if (Filter.DevId3 != null)
                {
                    if (Filter.filterWheel.Position == -1)
                        toolStripStatusLabel4.Text = "Filter Moving";
                    while (Filter.filterWheel.Position == -1)
                    {
                        Thread.Sleep(100);
                    }
                    //  currentfilter = Convert.ToInt16(Filter.filterWheel.Position);
                    //   Log(Filter.filterWheel.Names[currentfilter].ToString());
                    Filtertext = Filter.filterWheel.Names[Filter.filterWheel.Position];
                    FileLog2("Current Filter: " + Filtertext);
                }
                else
                    Filtertext = comboBox2.Text;
                if (checkBox31.Checked)
                {
                    if (currentfilter == 1)
                        Filtertext = comboBox2.Text;
                    if (currentfilter == 2)
                        Filtertext = comboBox3.Text;
                    if (currentfilter == 3)
                        Filtertext = comboBox4.Text;
                    if (currentfilter == 4)
                        Filtertext = comboBox5.Text;
                    if (currentfilter == 5)
                        Filtertext = comboBox1.Text;
                }


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

                if ((Filtertext == "Luminosity") || (Filtertext == "Clear") || (Filtertext == "No Filter"))
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
                    GlobalVariables.EquipPrefix = comboBox7.Text + "_" + Filtertext;
                if (radioButton9.Checked == true)
                    _equip = GlobalVariables.EquipPrefix + "_E";
                if (radioButton10.Checked == true)
                    _equip = GlobalVariables.EquipPrefix + " _R";
                else
                    _equip = GlobalVariables.EquipPrefix;
                /* //remd 4-2-14
                if ((Filtertext == "Luminosity") ||(Filtertext == "Clear")) //  rem'd 4_19 || (Filtertext == "Red") || (Filtertext == "Green") || (Filtertext == "Blue"))
                //only need this to start at lum w/ NB
                {
                    GlobalVariables.EquipPrefix = comboBox7.Text + "_" + "LRGB";
                    if (radioButton9.Checked == true)
                        _equip = GlobalVariables.EquipPrefix + "_E";
                    if (radioButton10.Checked == true)
                        _equip = GlobalVariables.EquipPrefix + " _R";
                    else
                        _equip = GlobalVariables.EquipPrefix;
                }
                 */
                //***************** 3-3-14 removed to allow metric to use single star v-curves *****************
                // %%%%% unremd 10-10-16 to redo full fam metric and allow selection of v-curve data
                if (checkBox22.Checked == true)
                    _equip = _equip + "_Metric";
                //********************************************************************************************
                if (_equip != null)
                {
                    toolStripStatusLabel3.Text = _equip.ToString();
                    //end filter/equip sync addition
                    toolStripStatusLabel4.Text = Filtertext.ToString();
                    toolStripStatusLabel2.Text = "Total Subs " + subCountCurrent.ToString() + " of " + TotalSubs().ToString();
                    //next line filters gridview by equip
                    //   Data d = new Data();
                    ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";
                    //  ((DataTable)d.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";
                    //FillData();//******These next 2 were rem'd ?????????? 8-20-13
                    // GetAvg();
                    //   textBox27.Text = Filtertext;
                }
                textBox21.Text = Filtertext.ToString();

                //**** try adding from previous FilterAdvnace()***********  3-2-14

                FlatCalcDone = false;
                if (SequenceRunning == true)
                {
                    DisableUpDwn();
                    // fileSystemWatcher1.EnableRaisingEvents = true;
                }
                //      toolStripStatusLabel1.Text = "Filter Moving";
                this.Refresh();
                UpdateData();
                EquipRefresh();
                //*****************  end addt*************************
            }
            catch
            {
                Log("Display Current Filter Error");
            }
        }

        private int FocusPerSubGroupCount;
        // checkfiltercounthere
        void checkfiltercount()//check current sub count and filter position, initiates first capture(subcount = 0)
        {
            //  try
            //   {
            //first 3 line duplicate w/ radio 7 enable, may be able to delete


            //  subsperfilter = totalsubs / filternumber;
            //textBox19.Text = subsperfilter.ToString();
            //   textBox27.Text = Filtertext;
            //    textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;

            textBox23.Text = TotalSubs().ToString();
            if (subCountCurrent == 0)//starts for first one
            {
                NebCapture();
                //  ****** 11-23-13   **** try to get this in background

                /*
                if (backgroundWorker2.IsBusy != true)
                {
                    // Start the asynchronous operation.
                    backgroundWorker2.RunWorkerAsync();
                }
                 */

            }
            if ((subCountCurrent == TotalSubs()) & (subCountCurrent != 0))//signifies end of sequence
            {
                //adioButton7.Checked = false;
                fileSystemWatcher4.EnableRaisingEvents = false;
                //NetworkStream serverStream = clientSocket.GetStream();
                if (!UseClipBoard.Checked)
                {
                    serverStream = clientSocket.GetStream();
                    byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                    serverStream.Write(outStream, 0, outStream.Length);
                    serverStream.Flush();
                    toolStripStatusLabel1.Text = "Sequence Done";
                    toolStripStatusLabel1.BackColor = Color.WhiteSmoke;
                    FileLog2("Sequence Done");
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
                    SetForegroundWindow(Handles.NebhWnd);
                    Thread.Sleep(1000);
                    PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                    Thread.Sleep(3000);
                    //       NebListenOn = false;
                    // clientSocket.GetStream().Close();//added 5-17-12
                    //  clientSocket.Client.Disconnect(true);//added 5-17-12
                    clientSocket.Close();
                }
                else
                {
                    Clipboard.Clear();
                    delay(1);

                    SetForegroundWindow(Handles.NebhWnd);
                    delay(1);
                    PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                    delay(1);
                    SendKeys.SendWait("~"); // this enter key clears frequent failed to read clipboard error
                    SendKeys.Flush();
                    //msdelay(1000);
                    //Clipboard.SetText("//NEB Listen 0");
                    //msdelay(1000);
                    toolStripStatusLabel1.Text = "Sequence Done";
                    toolStripStatusLabel1.BackColor = Color.WhiteSmoke;
                    FileLog2("Sequence Done");
                    this.Refresh();
                    if (DarksOn == true)
                        DarksOn = false;
                    currentfilter = 1;
                    DisplayCurrentFilter();
                    //      currentfilter = 0;
                    subCountCurrent = 0;
                    filterCountCurrent = 0;
                }
                toolStripStatusLabel4.BackColor = System.Drawing.Color.LightGray;
                toolStripStatusLabel4.Text = Filtertext.ToString();
                //  button12.UseVisualStyleBackColor = true;
                //   button31.UseVisualStyleBackColor = true;
                if (checkBox11.Checked == true)
                {
                    StopPHD();
                    //  StopTracking();
                    scope.Tracking = false;
                }
                if (phdsocket.Connected)
                    phdsocket.Close();
                Capturing = false;
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
                        WaitForSequenceDone("Sequence done", GlobalVariables.NebSlavehwnd);
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
                        WaitForSequenceDone("Sequence done", GlobalVariables.NebSlavehwnd);
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


        // for sequences
        private void fileSystemWatcher4_Changed(object sender, FileSystemEventArgs e)
        {
            idleCount = 0;
            //if (FlatDone == false)
            //  {
               FileLog2("SystemFileWatcher 4 changed");
            textBox41.Refresh();
            textBox41.Clear();
            if (!IsSlave())
            {
                subCountCurrent++;
                filterCountCurrent++;//the sub per this filter number
                ToolStripProgressBar();
                checkfiltercount();
                flipCheckDone = false; //reset every sub

                if (checkBox17.Checked == true)
                    FocusPerSubCurrent++;
            }

            //   if (checkBox23.Checked == true & FlipDone == false & okToFlip == true) 
            //      MeridianFlip();

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
        //private void FilterStepFwd()
        //{
        //    try
        //    {
        //        port.DiscardOutBuffer();
        //        port.DiscardInBuffer();
        //        Thread.Sleep(20);
        //        port.Write("F");
        //        Thread.Sleep(200);
        //    }
        //    catch (Exception ex)
        //    {
        //        Log("Filter Fwd Error" + ex.ToString());
        //    }
        //}


        //private void filterStepRev()
        //{
        //    try
        //    {
        //        port.DiscardOutBuffer();
        //        port.DiscardInBuffer();
        //        Thread.Sleep(20);
        //        port.Write("R");
        //        Thread.Sleep(200);
        //    }
        //    catch (Exception ex)
        //    {
        //        Log("Filter rev Error" + ex.ToString());
        //    }
        //}


        private void button19_Click(object sender, EventArgs e)
        {
            //    FilterStepFwd();

        }


        //filter advance on filter tab
        private void button20_Click(object sender, EventArgs e)
        {
            /*
            if (currentfilter == 0)//if unsynced just advance
                filteradvance();
            if ((currentfilter != 0) && (currentfilter != 4))//below allows display of current filter while manually changing
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
             */
            filteradvance();
            Update();  // 7-25-14 this maybe obsolete??????

        }

        //set max travel on setup tab
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
                travel = Convert.ToInt32(textBox2.Text.ToString());
            else
                return;

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


        private int CountFilterTotal()//this may not be needed.  seems to just check for non-zero at sequencego
        {
            int filter1used;
            int filter2used;
            int filter3used;
            int filter4used;
            int filter5used;
            int filternumber;
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

            return filternumber = filter4used + filter3used + filter2used + filter1used + filter5used + filter6used;//calculate total number of filters

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
            FileLog2("filtersequcene called");
            //   System.Object lockThis = new System.Object();
            // try
            // {
            DisplayCurrentFilter(); //try *****  added 5-23  
           

            //   textBox38.Text = currentfilter.ToString();
            if (currentfilter == 1)// was 1 prior to ascom   ***really is current sequence position***
            {

                if (FilterFocusGroupCurrent < FocusGroup[0])
                {
                    subsperfilter = SubsPerFocus[0];
                    FilterFocusGroupCurrent++;
                    Nebname = comboBox2.Text + "_" + FilterFocusGroupCurrent.ToString();
                    FilterFocus();

                    //    NebCapture();  //remd fixes prob 8-26-13

                    return;//may not be needed???
                }
                else
                    FilterFocusGroupCurrent = 0;//must be equal to nuber of groups so reset and move to next filter
                //should continue to cycle here until FFGC + FG[1]
                if (checkBox18.Checked == true)//flat every filter
                {
                    CaptureTime3 = FlatExp;
                    CaptureBin = (int)numericUpDown24.Value;
                    DoFlat();



                }
                if (checkBox2.Checked == true)
                {
                    currentfilter = 2;
                    //   filteradvance();
                    if (!checkBox31.Checked)
                        Filter.filterWheel.Position = (short)comboBox3.SelectedIndex;
                    DisplayCurrentFilter();
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
                        if ((checkBox8.Checked == false) & (checkBox17.Checked == false))//refocus after filter change check box
                            NebCapture();

                    }
                    return;

                }
                if (checkBox3.Checked == true)
                {
                    currentfilter = 3;
                    //   filteradvance();
                    //   filteradvance();
                    if (!checkBox31.Checked)
                        Filter.filterWheel.Position = (short)comboBox4.SelectedIndex;
                    DisplayCurrentFilter();
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
                    if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                        NebCapture();

                    //   Thread.Sleep(4000);
                    return;
                }
                if (checkBox4.Checked == true)
                {
                    currentfilter = 4;
                    //  filteradvance();
                    //   filteradvance();
                    //    filteradvance();
                    if (!checkBox31.Checked)
                        Filter.filterWheel.Position = (short)comboBox5.SelectedIndex;
                    DisplayCurrentFilter();
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
                    if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                        NebCapture();

                    return;
                }

                if (checkBox5.Checked == true)
                {
                    if (comboBox1.SelectedItem.ToString() == "Dark 1")
                    {
                        currentfilter = 5;
                        DarksOn = true;
                        // filteradvance();//back to pos 1 
                        Thread.Sleep(1000);
                        Handles.Editfound = 0;
                        Nebname = "Dark1";
                        subsperfilter = (int)numericUpDown20.Value;
                        CaptureTime = (int)numericUpDown19.Value * 1000;
                        CaptureBin = (int)numericUpDown28.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();

                        return;
                    }

                    else
                    {
                        currentfilter = 5;
                        //  filteradvance();
                        //   filteradvance();
                        //    filteradvance();
                        if (!checkBox31.Checked)
                            Filter.filterWheel.Position = (short)comboBox1.SelectedIndex;
                        DisplayCurrentFilter();
                        //    FlatDone = false;
                        //   textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = currentfilter.ToString();
                        if (checkBox8.Checked == true)
                            FilterFocus();
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown15.Value;
                            Nebname = comboBox1.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[4];
                            Nebname = comboBox1.Text + "_1";
                        }

                        // Thread.Sleep(4000);
                        //  subsperfilter = (int)numericUpDown15.Value;
                        //  Nebname = comboBox5.Text.ToString();
                        CaptureTime = (int)numericUpDown18.Value * 1000;
                        CaptureBin = (int)numericUpDown27.Value;
                        //   if (FilterFocusOn == false)
                        if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                            NebCapture();

                        return;
                    }
                }
                if (checkBox9.Checked == true)//Dark 2 allows for different time or bin
                {
                    if (comboBox8.SelectedItem.ToString() == "Dark 2")
                    {
                        currentfilter = 6;
                        DarksOn = true;
                        //no advance already at 1
                        Thread.Sleep(1000);
                        Handles.Editfound = 0;
                        Nebname = "Dark2";
                        subsperfilter = (int)numericUpDown29.Value;
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();

                        return;
                    }
                    else
                    {
                        currentfilter = 6;
                        //  filteradvance();
                        //   filteradvance();
                        //    filteradvance();
                        if (!checkBox31.Checked)
                            Filter.filterWheel.Position = (short)comboBox8.SelectedIndex;
                        DisplayCurrentFilter();
                        //    FlatDone = false;
                        //   textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = currentfilter.ToString();
                        if (checkBox8.Checked == true)
                            FilterFocus();
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown29.Value;
                            Nebname = comboBox8.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[5];
                            Nebname = comboBox8.Text + "_1";
                        }

                        // Thread.Sleep(4000);
                        //  subsperfilter = (int)numericUpDown15.Value;
                        //  Nebname = comboBox5.Text.ToString();
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        //   if (FilterFocusOn == false)
                        if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                            NebCapture();

                        return;
                    }

                }
                if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                {

                    currentfilter = 7;
                    if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                        Filter.filterWheel.Position = 0;
                    DisplayCurrentFilter();
                    if (FlatsOn == false)
                        ToggleFlat();
                    //filteradvance();//back to pos 1 
                    //   Filter.filterWheel.Position = 0;
                    Thread.Sleep(1000);
                    Handles.Editfound = 0;
                    Nebname = "Flat1";
                    subsperfilter = (int)numericUpDown32.Value;
                    CaptureTime = (int)numericUpDown34.Value;
                    CaptureBin = (int)numericUpDown36.Value;
                    FilterFocusOn = false;
                    //  fileSystemWatcher4.EnableRaisingEvents = true;

                    //11-10-16 changed all similar to this to only doflat();

                    //if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                    //{
                    //    //CalculateFlatExp(); //11-10-16 chenged all to doflat();
                    DoFlat();
                    //            return;
                    //    }

                    //     else
                    //      NebCapture();

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
                    //   NebCapture();
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
                    // filteradvance();
                    if (!checkBox31.Checked)
                        Filter.filterWheel.Position = (short)comboBox4.SelectedIndex;
                    DisplayCurrentFilter();
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
                    if ((checkBox8.Checked == false) & (checkBox17.Checked == false))//wasjust cb 8 == false
                        NebCapture();
                    // Thread.Sleep(4000);
                    return;
                }
                if (checkBox4.Checked == true)
                {

                    currentfilter = 4;
                    //  filteradvance();
                    //  filteradvance();
                    if (!checkBox31.Checked)
                        Filter.filterWheel.Position = (short)comboBox5.SelectedIndex;
                    DisplayCurrentFilter();
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
                    if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                        NebCapture();
                    //   Thread.Sleep(4000);
                    return;
                }
                if (checkBox5.Checked == true)
                {
                    currentfilter = 5;
                    // filteradvance();
                    //  filteradvance();
                    //   filteradvance();
                    if (comboBox6.SelectedItem.ToString() == "Dark 1")
                    {
                        DarksOn = true;
                        // filteradvance();//back to pos 1 
                        Thread.Sleep(1000);
                        Handles.Editfound = 0;
                        Nebname = "Dark1";
                        subsperfilter = (int)numericUpDown20.Value;
                        CaptureTime = (int)numericUpDown19.Value * 1000;
                        CaptureBin = (int)numericUpDown28.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();
                        return;
                    }
                    else
                    {
                        if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                            Filter.filterWheel.Position = (short)comboBox1.SelectedIndex;
                        DisplayCurrentFilter();
                        //    FlatDone = false;
                        // textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = Filtertext;
                        if (checkBox8.Checked == true)
                            FilterFocus();
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown20.Value;
                            Nebname = comboBox1.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[4];
                            Nebname = comboBox1.Text + "_1";
                        }
                        //  subsperfilter = (int)numericUpDown15.Value;
                        //  Nebname = comboBox5.Text.ToString();
                        CaptureTime = (int)numericUpDown19.Value * 1000;
                        CaptureBin = (int)numericUpDown28.Value;
                        //   if (FilterFocusOn == false)
                        if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                            NebCapture();
                        //   Thread.Sleep(4000);
                        return;
                    }

                }
                if (checkBox9.Checked == true)//Dark 2 allows for different time or bin
                {
                    currentfilter = 6;

                    if (comboBox8.SelectedItem.ToString() == "Dark 2")
                    {

                        DarksOn = true;
                        DisplayCurrentFilter();
                        //no advance already at 1
                        Thread.Sleep(1000);
                        Handles.Editfound = 0;
                        Nebname = "Dark2";
                        subsperfilter = (int)numericUpDown29.Value;
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();
                        return;
                    }
                    else
                    {
                        if ((!checkBox31.Checked) && (Filter.DevId3 != null)) // 11-8-16
                            Filter.filterWheel.Position = (short)comboBox8.SelectedIndex;
                        DisplayCurrentFilter();
                        //    FlatDone = false;
                        // textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = Filtertext;
                        if (checkBox8.Checked == true)
                            FilterFocus();
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown29.Value;
                            Nebname = comboBox8.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[5];
                            Nebname = comboBox8.Text + "_1";
                        }
                        //  subsperfilter = (int)numericUpDown15.Value;
                        //  Nebname = comboBox5.Text.ToString();
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        //   if (FilterFocusOn == false)
                        if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                            NebCapture();
                        //   Thread.Sleep(4000);
                        return;
                    }


                }
                if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                {
                    currentfilter = 7;
                    if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                        Filter.filterWheel.Position = 0;//back to pos 0 
                    DisplayCurrentFilter();
                    //    filteradvance();
                    //     filteradvance();
                    //      filteradvance();
                    if (FlatsOn == false)
                        ToggleFlat();
                    Thread.Sleep(1000);
                    Handles.Editfound = 0;
                    Nebname = "Flat1";
                    subsperfilter = (int)numericUpDown32.Value;
                    CaptureTime = (int)numericUpDown34.Value;
                    CaptureBin = (int)numericUpDown36.Value;
                    FilterFocusOn = false;
                    //  fileSystemWatcher4.EnableRaisingEvents = true;
                    //if ((FlatCalcDone == false) & (checkBox15.Checked == true)) //15 is flat autoexp calc
                    //{
                    //    // CalculateFlatExp();
                    DoFlat();
                    //    return;
                    //}

                    //else
                    //    NebCapture();
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
                    //  NebCapture();
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
                    //  filteradvance();
                    if (!checkBox31.Checked)
                        Filter.filterWheel.Position = (short)comboBox5.SelectedIndex;
                    DisplayCurrentFilter();
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
                    if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                        NebCapture();
                    //Thread.Sleep(4000);
                    return;
                }
                if (checkBox5.Checked == true)
                {
                    currentfilter = 5;
                    // filteradvance();
                    //  filteradvance();
                    //   filteradvance();
                    if (comboBox6.SelectedItem.ToString() == "Dark 1")
                    {
                        DarksOn = true;
                        DisplayCurrentFilter();
                        // filteradvance();//back to pos 1 
                        Thread.Sleep(1000);
                        Handles.Editfound = 0;
                        Nebname = "Dark1";
                        subsperfilter = (int)numericUpDown20.Value;
                        CaptureTime = (int)numericUpDown19.Value * 1000;
                        CaptureBin = (int)numericUpDown28.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();
                        return;
                    }
                    else
                    {
                        if (!checkBox31.Checked)
                            Filter.filterWheel.Position = (short)comboBox1.SelectedIndex;
                        DisplayCurrentFilter();
                        //    FlatDone = false;
                        // textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = Filtertext;
                        if (checkBox8.Checked == true)
                            FilterFocus();
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown20.Value;
                            Nebname = comboBox1.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[4];
                            Nebname = comboBox1.Text + "_1";
                        }
                        //  subsperfilter = (int)numericUpDown15.Value;
                        //  Nebname = comboBox5.Text.ToString();
                        CaptureTime = (int)numericUpDown19.Value * 1000;
                        CaptureBin = (int)numericUpDown28.Value;
                        //   if (FilterFocusOn == false)
                        if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                            NebCapture();
                        //   Thread.Sleep(4000);
                        return;
                    }

                }

                if (checkBox9.Checked == true)//Dark 2 allows for different time or bin
                {
                    currentfilter = 6;

                    if (comboBox8.SelectedItem.ToString() == "Dark 2")
                    {

                        DarksOn = true;
                        DisplayCurrentFilter();
                        //no advance already at 1
                        Thread.Sleep(1000);
                        Handles.Editfound = 0;
                        Nebname = "Dark2";
                        subsperfilter = (int)numericUpDown29.Value;
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();
                        return;
                    }
                    else
                    {
                        if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                            Filter.filterWheel.Position = (short)comboBox8.SelectedIndex;
                        DisplayCurrentFilter();
                        //    FlatDone = false;
                        // textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = Filtertext;
                        if (checkBox8.Checked == true)
                            FilterFocus();
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown29.Value;
                            Nebname = comboBox8.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[5];
                            Nebname = comboBox8.Text + "_1";
                        }
                        //  subsperfilter = (int)numericUpDown15.Value;
                        //  Nebname = comboBox5.Text.ToString();
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        //   if (FilterFocusOn == false)
                        if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                            NebCapture();
                        //   Thread.Sleep(4000);
                        return;
                    }


                }
                if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                {
                    currentfilter = 7;
                    if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                        Filter.filterWheel.Position = 0;
                    DisplayCurrentFilter();
                    //  filteradvance();
                    //  filteradvance();
                    if (FlatsOn == false)
                        ToggleFlat();
                    //  filteradvance();//back to pos 1 
                    Thread.Sleep(1000);
                    Handles.Editfound = 0;
                    Nebname = "Flat1";
                    subsperfilter = (int)numericUpDown32.Value;
                    CaptureTime = (int)numericUpDown34.Value;
                    CaptureBin = (int)numericUpDown36.Value;
                    FilterFocusOn = false;
                    //  fileSystemWatcher4.EnableRaisingEvents = true;
                    //if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                    //{
                    DoFlat();
                    //   // CalculateFlatExp();
                    //    return;
                    //}

                    //else
                    //    NebCapture();
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
                    //    NebCapture();
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
                    // filteradvance();
                    //  filteradvance();
                    //   filteradvance();
                    if (comboBox1.SelectedItem.ToString() == "Dark 1")
                    {
                        DarksOn = true;
                        DisplayCurrentFilter();
                        // filteradvance();//back to pos 1 
                        Thread.Sleep(1000);
                        Handles.Editfound = 0;
                        Nebname = "Dark1";
                        subsperfilter = (int)numericUpDown20.Value;
                        CaptureTime = (int)numericUpDown19.Value * 1000;
                        CaptureBin = (int)numericUpDown28.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();
                        return;
                    }
                    else
                    {
                        if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                            Filter.filterWheel.Position = (short)comboBox1.SelectedIndex;
                        DisplayCurrentFilter();
                        //    FlatDone = false;
                        // textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = Filtertext;
                        if (checkBox8.Checked == true)
                            FilterFocus();
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown20.Value;
                            Nebname = comboBox1.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[4];
                            Nebname = comboBox1.Text + "_1";
                        }
                        //  subsperfilter = (int)numericUpDown15.Value;
                        //  Nebname = comboBox5.Text.ToString();
                        CaptureTime = (int)numericUpDown19.Value * 1000;
                        CaptureBin = (int)numericUpDown28.Value;
                        //   if (FilterFocusOn == false)
                        if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                            NebCapture();
                        //   Thread.Sleep(4000);
                        return;
                    }

                }

                /*
                if (checkBox5.Checked == true)
                {

                    currentfilter = 5;
                    DarksOn = true;
                    filteradvance();//back to pos 1 
                    Thread.Sleep(1000);
                    Handles.Editfound = 0;
                    Nebname = "Dark1";
                    subsperfilter = (int)numericUpDown20.Value;
                    CaptureTime = (int)numericUpDown19.Value * 1000;
                    CaptureBin = (int)numericUpDown28.Value;
                    FilterFocusOn = false;
                    //  fileSystemWatcher4.EnableRaisingEvents = true;
                    NebCapture();
                    return;
                */
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
                //   }

                if (checkBox9.Checked == true)//Dark 2 allows for different time or bin
                {
                    currentfilter = 6;

                    if ((comboBox8.SelectedItem.ToString() == "Dark 2") || (comboBox8.SelectedItem.ToString() == "Dark 1"))
                    {

                        DarksOn = true;
                        DisplayCurrentFilter();
                        //no advance already at 1
                        Thread.Sleep(1000);
                        Handles.Editfound = 0;
                        if (comboBox8.SelectedItem.ToString() == "Dark 2")
                            Nebname = "Dark2";
                        if (comboBox8.SelectedItem.ToString() == "Dark 1")
                            Nebname = "Dark1";
                        subsperfilter = (int)numericUpDown29.Value;
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();
                        return;
                    }
                    else
                    {
                        if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                            Filter.filterWheel.Position = (short)comboBox8.SelectedIndex;
                        DisplayCurrentFilter();
                        //    FlatDone = false;
                        // textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = Filtertext;
                        if (checkBox8.Checked == true)
                            FilterFocus();
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown29.Value;
                            Nebname = comboBox8.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[5];
                            Nebname = comboBox8.Text + "_1";
                        }
                        //  subsperfilter = (int)numericUpDown15.Value;
                        //  Nebname = comboBox5.Text.ToString();
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        //   if (FilterFocusOn == false)
                        if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                            NebCapture();
                        //   Thread.Sleep(4000);
                        return;
                    }


                }
                if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                {
                    currentfilter = 7;
                    if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                        Filter.filterWheel.Position = 0;
                    DisplayCurrentFilter();
                    if (FlatsOn == false)
                        ToggleFlat();
                    //   filteradvance();//back to pos 1 
                    Thread.Sleep(1000);
                    Handles.Editfound = 0;
                    Nebname = "Flat1";
                    subsperfilter = (int)numericUpDown32.Value;
                    CaptureTime = (int)numericUpDown34.Value;
                    CaptureBin = (int)numericUpDown36.Value;
                    FilterFocusOn = false;
                    //  fileSystemWatcher4.EnableRaisingEvents = true;

                    //if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                    //{
                    DoFlat();
                    //    //CalculateFlatExp();
                    //    return;
                    //}

                    //else
                    //    NebCapture();
                    return;
                }


            }

            if (currentfilter == 5)
            {
                if (FilterFocusGroupCurrent < FocusGroup[4])
                {
                    FilterFocus();
                    subsperfilter = SubsPerFocus[4];
                    FilterFocusGroupCurrent++;
                    Nebname = comboBox1.Text + "_" + FilterFocusGroupCurrent.ToString();
                    //    NebCapture();
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


                if (checkBox9.Checked == true)//Dark 2 allows for different time or bin
                {
                    currentfilter = 6;

                    if (comboBox8.SelectedItem.ToString() == "Dark 2")
                    {

                        DarksOn = true;
                        DisplayCurrentFilter();
                        //no advance already at 1
                        Thread.Sleep(1000);
                        Handles.Editfound = 0;
                        Nebname = "Dark2";
                        subsperfilter = (int)numericUpDown29.Value;
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        FilterFocusOn = false;
                        //  fileSystemWatcher4.EnableRaisingEvents = true;
                        NebCapture();
                        return;
                    }
                    else
                    {
                        if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                            Filter.filterWheel.Position = (short)comboBox8.SelectedIndex;
                        DisplayCurrentFilter();
                        //    FlatDone = false;
                        // textBox21.Text = "Pos" + currentfilter.ToString() + Filtertext;
                        //   textBox27.Text = Filtertext;
                        if (checkBox8.Checked == true)
                            FilterFocus();
                        if (checkBox17.Checked == false)
                        {
                            subsperfilter = (int)numericUpDown29.Value;
                            Nebname = comboBox8.Text;
                        }
                        else
                        {
                            FilterFocusGroupCurrent = 1;
                            subsperfilter = SubsPerFocus[5];
                            Nebname = comboBox8.Text + "_1";
                        }
                        //  subsperfilter = (int)numericUpDown15.Value;
                        //  Nebname = comboBox5.Text.ToString();
                        CaptureTime = (int)numericUpDown30.Value * 1000;
                        CaptureBin = (int)numericUpDown31.Value;
                        //   if (FilterFocusOn == false)
                        if ((checkBox8.Checked == false) & (checkBox17.Checked == false))
                            NebCapture();
                        //   Thread.Sleep(4000);
                        return;
                    }


                }
                if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                {

                    currentfilter = 7;
                    if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                        Filter.filterWheel.Position = 0;
                    DisplayCurrentFilter();
                    if (FlatsOn == false)
                        ToggleFlat();
                    //  filteradvance();//back to pos 1 
                    Thread.Sleep(1000);
                    Handles.Editfound = 0;
                    Nebname = "Flat1";
                    subsperfilter = (int)numericUpDown32.Value;
                    CaptureTime = (int)numericUpDown34.Value;
                    CaptureBin = (int)numericUpDown36.Value;
                    FilterFocusOn = false;
                    //  fileSystemWatcher4.EnableRaisingEvents = true;
                    //if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                    //{
                    DoFlat();
                    //    //CalculateFlatExp();
                    //    return;
                    //}

                    //else
                    //    NebCapture();
                    return;
                }


            }
            if (currentfilter == 6)
            {
                if (FilterFocusGroupCurrent < FocusGroup[5])
                {
                    FilterFocus();
                    subsperfilter = SubsPerFocus[5];
                    FilterFocusGroupCurrent++;
                    Nebname = comboBox8.Text + "_" + FilterFocusGroupCurrent.ToString();
                    //    NebCapture();
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
                if ((checkBox13.Checked == true) & (checkBox18.Checked == false))
                {
                    currentfilter = 7;
                    if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                        Filter.filterWheel.Position = 0;
                    DisplayCurrentFilter();
                    if (FlatsOn == false)
                        ToggleFlat();
                    //  filteradvance();//back to pos 1 
                    Thread.Sleep(1000);
                    Handles.Editfound = 0;
                    Nebname = "Flat1";
                    subsperfilter = (int)numericUpDown32.Value;
                    CaptureTime = (int)numericUpDown34.Value;
                    CaptureBin = (int)numericUpDown36.Value;
                    FilterFocusOn = false;
                    //  fileSystemWatcher4.EnableRaisingEvents = true;
                    //if ((FlatCalcDone == false) & (checkBox15.Checked == true))
                    //{
                    DoFlat();
                    //   // CalculateFlatExp();
                    //    return;
                    //}

                    //else
                    //    NebCapture();
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

                FileLog2("NebCapture called");
                if (IsServer())
                    SendtoSlave((subsperfilter * CaptureTime / 1000).ToString());
                if (IsSlave())
                {//set # of subs based on server #subs * capture time.
                    CaptureTime = Convert.ToInt32(textBox37.Text) * 1000;
                    subs = Convert.ToDouble(textBox41.Text) * 1000 / CaptureTime;
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
                                       //  if (clientSocket.Connected == false)// added this line 7-1-14
                NebListenStart(Handles.NebhWnd, SocketPort);

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
                    toolStripStatusLabel1.BackColor = Color.Lime; // added 10-24-16
                                                                  //his.Refresh(); 
                }
                string prefix = textBox19.Text.ToString();
                //   NetworkStream serverstream;

                // 11-8-16 clipboard may not work for slave TODO
                if (IsSlave())
                    serverStream = clientSocket2.GetStream();
                else
                {
                    if (!UseClipBoard.Checked) // 11-8-16  just the if
                    {
                        try
                        {
                            serverStream = clientSocket.GetStream();
                        }
                        catch (IOException e)
                        {
                            Log("Neb socket failure " + e.ToString());
                            NebCapture();
                        }
                    }
                }
                if (FlatCalcDone == false)
                {

                    if ((Handles.NebVNumber == 3) || (Handles.NebVNumber == 4))
                        // if (Handles.NebVNumber == 3)
                        CaptureTime3 = (double)CaptureTime / 1000;//need to allow fractions for flats.
                    //      textBox38.Text = CaptureTime3.ToString();//*******just for debugging                       
                }
                if (FlatCalcDone == true)
                    FlatCalcDone = false;
                if (DarksOn == false)
                {
                    if (checkBox31.Checked)
                    {
                        // 11-8-16 TODO fix for clipboard use
                        intFilterPos = currentfilter;
                        byte[] outStream = System.Text.Encoding.ASCII.GetBytes("SetExtFilter " + intFilterPos + "\n" + "setname " + prefix + Nebname + "\n" + "setbinning " + CaptureBin + "\n" + "SetShutter 0" + "\n" + "SetDuration " + CaptureTime3 + "\n" + "Capture " + subsperfilter + "\n");
                        //use the one above for testing w/ external ASCOM Filter.filterWheel
                        //  byte[] outStream = System.Text.Encoding.ASCII.GetBytes("SetFilter " + intFilterPos + "\n" + "setname " + prefix + Nebname + "\n" + "setbinning " + CaptureBin + "\n" + "SetShutter 0" + "\n" + "SetDuration " + CaptureTime3 + "\n" + "Capture " + subsperfilter + "\n");
                        try
                        {
                            serverStream.Write(outStream, 0, outStream.Length);
                            Log(prefix + Nebname + "  Captime " + CaptureTime3.ToString() + "    subs " + subsperfilter.ToString());
                            FileLog2(prefix + Nebname + "  Captime " + CaptureTime3.ToString() + "    subs " + subsperfilter.ToString());
                        }
                        catch
                        {
                            MessageBox.Show("Error sending command", "scopefocus");
                            return;

                        }
                    }
                    else
                    {
                        if (!UseClipBoard.Checked) // 11-8-15 w/ else below
                        {
                            byte[] outStream = System.Text.Encoding.ASCII.GetBytes("setname " + prefix + Nebname + "\n" + "setbinning " + CaptureBin + "\n" + "SetShutter 0" + "\n" + "SetDuration " + CaptureTime3 + "\n" + "Capture " + subsperfilter + "\n");

                            try
                            {
                                serverStream.Write(outStream, 0, outStream.Length);
                                Log(prefix + Nebname + "  Captime " + CaptureTime3.ToString() + "    subs " + subsperfilter.ToString());
                                FileLog2(prefix + Nebname + "  Captime " + CaptureTime3.ToString() + "    subs " + subsperfilter.ToString());
                            }
                            catch
                            {
                                MessageBox.Show("Error sending command", "scopefocus");
                                return;

                            }
                        }
                        else
                        {
                            Log(prefix + Nebname + "  Captime " + CaptureTime3.ToString() + "    subs " + subsperfilter.ToString());
                            FileLog2(prefix + Nebname + "  Captime " + CaptureTime3.ToString() + "    subs " + subsperfilter.ToString());
                            //   int i = 0;
                            //   msdelay(750);
                            Clipboard.Clear();
                            msdelay(500);
                            Clipboard.SetDataObject("//NEB SetName " + prefix + Nebname, false, 3, 500);
                            // Clipboard.SetText("//NEB SetName " + prefix + Nebname);
                            msdelay(750);
                            // 11-12 16 added nebcommandconfirms below 


                            //if (!NebCommandConfirm("SetName " + prefix + Nebname, 0))
                            //    if (!clipboardRetry)
                            //    {
                            //        clipboardRetry = true;
                            //        NebCapture();
                            //    }
                            //    else
                            //        ClipboardFailure("set name");
      
                            //{
                            //    i++;
                            //    msdelay(100);
                            //    if (i == 30)
                            //    {
                            //        Log("NebCapture Failed");
                            //        Send("NebCapture Failed");
                            //        return;
                            //    }
                            //}
                            //i = 0;


                            //  Clipboard.SetText("//NEB SetBinning " + CaptureBin);
                            Clipboard.SetDataObject("//NEB SetBinning " + CaptureBin, false, 3, 500);
                            msdelay(750);
                            //if (!NebCommandConfirm("SetBinning " + CaptureBin, 0))
                            //    if (!clipboardRetry)
                            //    {
                            //        clipboardRetry = true;
                            //        NebCapture();
                            //    }
                            //    else
                            //        ClipboardFailure("SetBinning");


                            //while (!NebCommandConfirm("SetBinning "+ CaptureBin, 0))
                            //{
                            //    i++;
                            //    msdelay(100);
                            //    if (i == 30)
                            //    {
                            //        Log("NebCapture Failed");
                            //        Send("NebCapture Failed");
                            //        return;
                            //    }
                            //}
                            //i = 0;
                            ////    Clipboard.SetText("//NEB SetShutter 0");
                            Clipboard.SetDataObject("//NEB SetShutter 0", false, 3, 500);
                            msdelay(750);
                            //if (!NebCommandConfirm("SetShutter 0", 0))
                            //    if (!clipboardRetry)
                            //    {
                            //        clipboardRetry = true;
                            //        NebCapture();
                            //    }
                            //    else
                            //        ClipboardFailure("setShutter");
                            //while (!NebCommandConfirm("SetShutter 0", 0))
                            //{
                            //    i++;
                            //    msdelay(100);
                            //    if (i == 30)
                            //    {
                            //        Log("NebCapture Failed");
                            //        Send("NebCapture Failed");
                            //        return;
                            //    }
                            //}
                            //i = 0;
                            //  Clipboard.SetText("//NEB SetDuration " + CaptureTime3);
                            Clipboard.SetDataObject("//NEB SetDuration " + CaptureTime3, false, 3, 500);
                            msdelay(750);
                            //if (!NebCommandConfirm("SetDuration " + CaptureTime3, 0))
                            //    if (!clipboardRetry)
                            //    {
                            //        clipboardRetry = true;
                            //        NebCapture();
                            //    }
                            //    else
                                  //  ClipboardFailure("SetDuration");
                            //while (!NebCommandConfirm("SetDuration " + CaptureTime3, 0))
                            //{
                            //    i++;
                            //    msdelay(100);
                            //    if (i == 30)
                            //    {
                            //        Log("NebCapture Failed");
                            //        Send("NebCapture Failed");
                            //        return;
                            //    }
                            //}

                            //  Clipboard.SetText("//NEB Capture " + subsperfilter);
                            Clipboard.SetDataObject("//NEB Capture " + subsperfilter, false, 3, 500);
                            msdelay(750);
                            // todo confirm

                            //if (!NebCommandConfirm("Exposing", 1))
                            //    if (!clipboardRetry)
                            //    {
                            //        clipboardRetry = true;
                            //        NebCapture();
                            //    }
                            //    else
                            //        ClipboardFailure("Exposing");

                            //Clipboard.Clear();
                            Capturing = true; // added 11-9-16 
                        }
                    }
                }
                //tryadd setforeground ***** 3-8-13  ****
                SetForegroundWindow(Handle.ToInt32());

                if (DarksOn == true)
                {

                    if (!UseClipBoard.Checked)
                    {
                        byte[] outStream2 = System.Text.Encoding.ASCII.GetBytes("setname " + prefix + Nebname + "\n" + "setbinning " + CaptureBin + "\n" + "SetShutter 1" + "\n" + "SetDuration " + CaptureTime3 + "\n" + "Capture " + subsperfilter + "\n");
                        try
                        {
                            serverStream.Write(outStream2, 0, outStream2.Length);
                            Log("Darks On  Cap time " + CaptureTime3.ToString());
                            FileLog2("Darks On  Cap time " + CaptureTime3.ToString());
                        }
                        catch
                        {
                            MessageBox.Show("Error sending command", "scopefocus");
                            return;

                        }
                    }
                    else
                    {
                        //  int i= 0;
                        //  msdelay(750);
                        FileLog2("nebCapture Darks on = true");
                        Clipboard.Clear();
                        msdelay(500);
                        //  Clipboard.SetText("//NEB setname " + prefix + Nebname);
                        Clipboard.SetDataObject("//NEB setname " + prefix + Nebname, false, 3, 500);
                        msdelay(750);
                        //if (!NebCommandConfirm("setname " + prefix + Nebname, 0))
                        //    if (!clipboardRetry)
                        //    {
                        //        clipboardRetry = true;
                        //        NebCapture();
                        //    }
                        //    else
                        //        ClipboardFailure("setName");
                        //while (!NebCommandConfirm("SetDuration " + CaptureTime3, 0))
                        //{
                        //    i++;
                        //    msdelay(100);
                        //    if (i == 30)
                        //    {
                        //        Log("NebCapture Failed");
                        //        Send("NebCapture Failed");
                        //        return;
                        //    }
                        //}
                        //i = 0;
                        //    Clipboard.SetText("//NEB setbinning " + CaptureBin);
                        Clipboard.SetDataObject("//NEB setbinning " + CaptureBin, false, 3, 500);
                        msdelay(750);
                        //if (!NebCommandConfirm("setbinning " + CaptureBin, 0))
                        //    if (!clipboardRetry)
                        //    {
                        //        clipboardRetry = true;
                        //        NebCapture();
                        //    }
                        //    else
                        //        ClipboardFailure("SetBinning");
                        //while (!NebCommandConfirm("SetDuration " + CaptureTime3, 0))
                        //{
                        //    i++;
                        //    msdelay(100);
                        //    if (i == 30)
                        //    {
                        //        Log("NebCapture Failed");
                        //        Send("NebCapture Failed");
                        //        return;
                        //    }
                        //}
                        //i = 0;
                        //  Clipboard.SetText("//NEB SetShutter 1");
                        Clipboard.SetDataObject("//NEB SetShutter 1", false, 3, 500);
                        msdelay(750);
                        //if (!NebCommandConfirm("SetShutter 1", 0))
                        //    if (!clipboardRetry)
                        //    {
                        //        clipboardRetry = true;
                        //        NebCapture();
                        //    }
                        //    else
                        //        ClipboardFailure("setShutter");
                        //while (!NebCommandConfirm("SetDuration " + CaptureTime3, 0))
                        //{
                        //    i++;
                        //    msdelay(100);
                        //    if (i == 30)
                        //    {
                        //        Log("NebCapture Failed");
                        //        Send("NebCapture Failed");
                        //        return;
                        //    }
                        //}
                        //i = 0;
                        //  Clipboard.SetText("//NEB SetDuration " + CaptureTime3);
                        Clipboard.SetDataObject("//NEB SetDuration " + CaptureTime3, false, 3, 500);
                        msdelay(750);
                        //if (!NebCommandConfirm("SetDuration " + CaptureTime3, 0))
                        //    if (!clipboardRetry)
                        //    {
                        //        clipboardRetry = true;
                        //        NebCapture();
                        //    }
                        //    else
                        //        ClipboardFailure("setDuration");
                        //while (!NebCommandConfirm("SetDuration " + CaptureTime3, 0))
                        //{
                        //    i++;
                        //    msdelay(100);
                        //    if (i == 30)
                        //    {
                        //        Log("NebCapture Failed");
                        //        Send("NebCapture Failed");
                        //        return;
                        //    }
                        //}
                        //i = 0;
                        //     Clipboard.SetText("//NEB Capture " + subsperfilter);
                        Clipboard.SetDataObject("//NEB Capture " + subsperfilter, false, 3, 500);
                        msdelay(750);
                        // todo confirm
                        //if (!NebCommandConfirm("Exposing", 1))
                        //    if (!clipboardRetry)
                        //    {
                        //        clipboardRetry = true;
                        //        NebCapture();
                        //    }
                        //    else
                        //        ClipboardFailure("Exposing");
                        Capturing = true;
                        Clipboard.Clear();
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


                    //************** 3-8-13 for meridian flip...cant monitor anyting w/ this while statement is going ******
                    //************  may need to change so sytemfilewatch4 monitors subs and counts rathing than waiting for neb status change ****                   
                    //  List<string> lastKnown = new List<string>();
                    string[] lastKnown = System.IO.Directory.GetFiles(GlobalVariables.Path2, "*.fit", System.IO.SearchOption.TopDirectoryOnly);
                    //  foreach (string file in files)
                    //    Log("files at start of capture" + file);


                    //2 lines below should be deleted if 11-23-13 rem is removed

                    if (backgroundWorker2.IsBusy != true)
                        backgroundWorker2.RunWorkerAsync();
                    //*****  11-23-13 remd
                    //  while (Capturing == true)
                    //   {
                    //****** end
                    //2 lines below should be deleted if 11-23-13 rem is removed

                    //   if (backgroundWorker2.IsBusy != true)
                    //       backgroundWorker2.RunWorkerAsync();
                    // ********* 11-23-13 rem'd
                    //     CheckForFlip();
                    //*************end rem                     
                    /*  

                      //this was added 3-8-13 to monitor the sub directory for new subs....
                      //since filesystemwatcher cannot in this while loop.  
                      string[] files2 = Directory.GetFiles(path2, "*.fit", System.IO.SearchOption.TopDirectoryOnly);
                      List<string> newFiles = new List<string>();

                      foreach (string s in files2)
                      {
                          if (!lastKnown.Contains(s))
                          {
                              newFiles.Add(s);
                            //  Log("added to newfiles" + s);
                          }
                      }

                     // List<string> newFiles = hasNewFiles(path2, lastKnown);
                      foreach (string f in newFiles)
                          Log("Sub " + f);
                      if (newFiles.Count > 0)
                      {
                         // processFiles(newFiles);
                          lastKnown = files2;
                          newFiles.Clear();

//****try add the whole thing****

                          DateTime saveNow = DateTime.Now;
                          GST = CalcGSTFromUT(saveNow);
                          TimeDec = ConvTimeToDec(GST);
                          Urania = ConvDecTUraniaTime(TimeDec);
                          int Usec = Urania.Second;
                          int Umin = Urania.Minute;
                          int Uhr = Urania.Hour;
                          string currentU = Uhr + ":" + Umin + ":" + Usec;
                          DateTime Uhms = Convert.ToDateTime(currentU);
                          textBox44.Text = currentU.ToString();
                          textBox45.Text = saveNow.ToString();
                          textBox46.Text = GetJulianDay(saveNow, -6).ToString();
                          textBox47.Text = CalcUTFromZT(saveNow, -6).ToString();

                          //get current locaton
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
                          CurrentLoc = port2.ReadExisting();
                          ConvertDec(CurrentLoc, out DecDeg, out DecMin, out DecSec);
                          ConvertRA(CurrentLoc, out RAhr, out RAmin, out RAsec);
                          //  Log("Focus Position Set: RA " + RAhr.ToString() + "hr " + RAmin.ToString() + "m " + RAsec.ToString() + "s " + "- Dec " + DecDeg.ToString() + "° " + DecMin.ToString() + "' " + DecSec.ToString() + "\"");
                          Thread.Sleep(100);
                          port2.DiscardInBuffer();
                          Thread.Sleep(10);
                          if (CurrentLoc.Substring(17, 1) == "#")//check for stop bit
                          {
                              // FocusLocObtained = true;
                              //  button32.BackColor = System.Drawing.Color.Lime;
                              //   button32.Text = "Focus Pos Set";

                              string path = textBox11.Text.ToString();
                              string fullpath = path + @"\log.txt";
                              StreamWriter log;
                              //   string string4 = "Focus Position Set: RA " + RAhr.ToString() + "hr " + RAmin.ToString() + "m " + RAsec.ToString() + "s " + "- Dec " + DecDeg.ToString() + "° " + DecMin.ToString() + "' " + DecSec.ToString() + "\"";
                              string currentRA = RAhr.ToString() + ":" + RAmin.ToString() + ":" + Convert.ToInt32(RAsec).ToString();
                              textBox49.Text = currentRA.ToString();
                              DateTime RAtime = Convert.ToDateTime(currentRA); // Converts only the time
                              TimeCompare = DateTime.Compare(RAtime, Uhms);
                              diff = (RAtime - Uhms);
                              FlipTime = (saveNow + diff);
                              TimeSpan TimeToFlip = FlipTime.TimeOfDay - DateTime.Now.TimeOfDay;
                              Log("Meridian Flip in " + TimeToFlip.ToString());
                              textBox50.Text = FlipTime.ToString("HH:mm:ss");
                              textBox48.Text = TimeCompare.ToString();
                              if (!File.Exists(fullpath))
                              {
                                  log = new StreamWriter(fullpath);
                              }
                              else
                              {
                                  log = File.AppendText(fullpath);
                              }
                              log.WriteLine(currentRA);
                              log.Close();
                              port2.Close();
                          }

                          if (TimeCompare == -1 & checkBox23.Checked == true & FlipDone == false & okToFlip == true)
                          {
                              GotoTargetLocation();
                              FlipDone = true;
                              Log("Meridain Flip done at " + DateTime.Now.ToString("HH:mm:ss"));
                          }



                         // button48.PerformClick();
                         // MeridianFlip();
         //******check for meridian flip here ???????******  3-8-13

                          Log("Sub taken");
                      }
                      //  end folder watching addition

                      */

                    /*
                     // **********  remd 11-23-13

                    int StatusstripHandle = FindWindowEx(Handles.NebhWnd, 0, "msctls_statusbar32", null);

                    //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                    IntPtr statusHandle = new IntPtr(StatusstripHandle);
                    StatusHelper sh = new StatusHelper(statusHandle);
                    string[] captions = sh.Captions;

                    if (captions[0] == "Sequence done")
                        Capturing = false;
                     //*******************end 11-23-13 rem*****************
                     //*******************end 3-11-13 rem*****************
                     */
                    //     }  /***** remd 11-23-13
                    // }

                }

                if (!UseClipBoard.Checked) // 11-8-16
                    serverStream.Flush();

                //   toolStripStatusLabel1.Text = " ";

            }
            

            catch (Exception e)
            {
                Log("NebCapture Error" + e.ToString());
                Send("NebCapture Error" + e.ToString());
                FileLog("NebCapture Error" + e.ToString());

            }


        }

        private void ClipboardFailure(string text)
        {
            clipboardRetry = false;
            StackTrace st = new StackTrace();
            Log("Clipboard failure, process aborted from: " + " " + text + " Method: " +st.GetFrame(1).GetMethod().Name.ToString());
            FileLog2("Clipboard failure, process aborted from: " + " " + text + " Method: " + st.GetFrame(1).GetMethod().Name.ToString());
            Send("Clipboard failure, process aborted from: " + " " + text + " Method: " + st.GetFrame(1).GetMethod().Name.ToString());
        }
        private bool NebCommandConfirm(string command, int arrayNo)
        {
            //  // status shows e.g.  captions[0] is Running command: SetDuration 5
            //     captions[1] says exposure done when done.  
            try
            {
                //int number;
                //if (arrayNo == 0)
                //    number = 13;  // for Caption[0] shift the substring so "running commnad:" is not tested.
                //else
                //    number = 0; 

                int StatusstripHandle = FindWindowEx(Handles.NebhWnd, 0, "msctls_statusbar32", null);


                //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                IntPtr statusHandle = new IntPtr(StatusstripHandle);
                StatusHelper sh = new StatusHelper(statusHandle);
                string[] captions = sh.Captions;
                if ((captions[arrayNo] == null) || (captions[arrayNo] == ""))// make sure not empty
                    return false;


                else
                {
                    //  int length = captions[0].Length;
                    if (arrayNo == 0)
                    {
                        if (captions[0].Length == command.Length + 13) // if lengths arent = then false
                        {
                            if (captions[0].Substring(13, command.Length) == command) // if lengths equal, makse sure right command
                                return true;
                            else
                                return false;
                        }
                        else
                            return false;

                    }

                    if (arrayNo == 1)
                    {
                        if (captions[1].Length == command.Length)
                        {
                            if (captions[arrayNo].Substring(0, command.Length) == command) // for "Exposure done"
                                return true;
                            else
                                return false;
                        }
                        else
                        {
                            if ((captions[1].Length > 7) && (captions[1].Substring(0, 8) == command)) // innput exposing to see if exposure started
                            {
                                return true;
                            }
                            if ((captions[1].Length > 6) && (captions[1].Substring(0, 7) == command)) // innput saving to see if done
                            {
                                return true;
                            }
                            else
                                return false;

                        }

                    }
                    // if (arrayNo > 1)
                    return false;



                }    
                                                       
            }
            catch (Exception e)
            {
                
                Log("NebStatusMonitor error " + e.ToString());
                FileLog("NebStatusMonitor error " + e.ToString());
                return false;
            }
        }


        //monitornebstatushere

        private void MonitorNebStatus()
        {
            try
            {
                int StatusstripHandle = FindWindowEx(Handles.NebhWnd, 0, "msctls_statusbar32", null);

                //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                IntPtr statusHandle = new IntPtr(StatusstripHandle);
                StatusHelper sh = new StatusHelper(statusHandle);
                string[] captions = sh.Captions;
                if (captions[0] != "")
                {
                    //if (captions[0].Contains("..."))
                    //{
                    //    MessageBox.Show("NebStatusMonitor error - try increasing the width of the Nebulosity window");
                    //}
                   // FileLog2("Cap[0]: " + captions[0] + "    Cap[1]: " + captions[1]);
                    if (UseClipBoard.Checked)
                    {
                        if (captions[0] == "Script done")
                        {
                            Capturing = false;
                            backgroundWorker2.CancelAsync();
                        }
                    }
                    else
                    {
                        if (captions[0] == "Sequence done")
                        {
                            Capturing = false;
                            backgroundWorker2.CancelAsync();
                        }
                    }
                    // if (captions[0].Length > 8) 11-8-16
                    //if (captions[0] != "")
                    //{
                    if (captions[0].Length > 6)
                    {

                        if (UseClipBoard.Checked)
                        {
                            if (captions[0].Substring(0, 7) == "Running")
                            {
                                Capturing = false;
                                backgroundWorker2.CancelAsync();
                            }
                        }

                    }
                    if (captions[0].Length > 9)
                    {
                        if ((captions[0].Substring(0, 10) == "Requesting") || (captions[0].Substring(0, 10) == "DITHER: Wa") && (!flipCheckDone))
                            CheckForFlip();
                        if ((captions[0].Substring(0, 10) == "Requesting") && (flipCheckDone))
                        {
                            toolStripStatusLabel1.Text = "Waiting for Dither";
                            //toolStripStatusLabel1.BackColor = Color.Yellow;
                        }
                    }
                    //11-8-16 try this

                    //if (captions[0].Substring(0, 10) == "Sequence a") // 10-24-16 added the space_a to eliminate "sequence done" triggering.
                    //{
                    //   toolStripStatusLabel1.Text = "Capturing"; // remd 10-24-16
                    //  //toolStripStatusLabel1.BackColor = Color.Lime;
                    //}

                    // TODO  this way capCurrent can be used for pause/resume, knowing it was done.....
                    if (captions[0].Length > 19)
                    {
                        if (captions[0].Substring(0, 20) == "Sequence acquisition")
                        {
                            int markerSlash = captions[0].IndexOf('/');
                            int markerSpace2 = captions[0].IndexOf(' ', markerSlash - 3);
                            int markerSpace3 = captions[0].IndexOf(' ', markerSlash);
                            int lengthCurrent = markerSlash - markerSpace2 - 1;
                            int lengthTotal = markerSpace3 - markerSlash - 1;
                            GlobalVariables.CapCurrent = Convert.ToInt16(captions[0].Substring(markerSlash - lengthCurrent, lengthCurrent));
                            GlobalVariables.CapTotal = Convert.ToInt16(captions[0].Substring(markerSlash + 1, lengthTotal));
                            //  int capTotal = captions[0].Substring()
                            toolStripStatusLabel1.Text = captions[3] + " " + GlobalVariables.CapCurrent.ToString() + "/" + GlobalVariables.CapTotal.ToString();
                            Capturing = true; // ***  added 11-20-16 for timer2_tick and idle monitor this could screw things up
                        }
                    }
                    
                }

            }
            catch (Exception e)
            {
                Log("NebStatusMonitor error - try increasing the width of the Nebulosity window");
            //    FileLog("NebStatusMonitor error " + e.ToString());
                FileLog2("NebStatusMonitor error " + e.ToString());
            }
        }
        //******** add 3-8-13  *****
        /*
        static List<string> hasNewFiles(string path, List<string> lastKnownFiles)
        {
            List<string> files = Directory.GetFiles(path).ToList();
            List<string> newFiles = new List<string>();

            foreach (string s in files)
            {
                if (!lastKnownFiles.Contains(s))
                    newFiles.Add(s);
            }

            return new List<string>();
        }
        

        private void Listen0()
        {
            if (!UseClipBoard.Checked) // 11-8-16
            {
                if ((IsSlave()) & clientSocket2.Connected == false)
                {
                    NebListenStart(Handles.NebhWnd, SocketPort);
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
            }
            else
            {
                msdelay(750);
                //   Clipboard.SetText("//NEB  Listen 0");
                Clipboard.Clear();
                msdelay(500);
                Clipboard.SetDataObject("//NEB  Listen 0", false, 3, 500);
                msdelay(750);
                Clipboard.Clear();
            }
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
            toolStripProgressBar1.Maximum = TotalSubs();
            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Increment(1);

            toolStripProgressBar1.Value = subCountCurrent;
            toolStripStatusLabel2.Text = "Total Subs " + subCountCurrent.ToString() + " of " + TotalSubs().ToString(); // 11-8-16 added "total"

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

        public void DisableUpDwn()
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


        //   public bool SequenceRunning = false;
        //Go button, goes to starting filter pos, (assumes started at pos 1) and begins script
        //sequencegohere
        private void SequenceGo()
        {
            try
            {
                ResizeNebWindow();
                idleCount = 0;
                FocusGroupCalc();  //redo in case restarting a previously aborted sequence.   11-11-16
                FileLog2("Sequence Go");
                // TODO ...have filelog2 of all filters, exp number, duration and binning.  and if focus between and flats....
                if (comboBox2.SelectedItem == null)
                {
                    DialogResult result = MessageBox.Show("you must enter a discription for filter 1", "scopefocus", MessageBoxButtons.OK);
                    if (result == DialogResult.OK)
                    {
                        comboBox2.Focus();
                        return;
                    }
                }

                if (checkBox13.Checked == true && devId4 == null)
                {
                    MessageBox.Show("Flats selected but not connected to ASCOM switch", "scopefocus");
                    return;
                }


                if ((IsServer()) & (StatusHandleRec == false))
                {
                    GlobalVariables.NebSlavehwnd = Convert.ToInt32(textBox41.Text);
                    Log("NebSlave handle " + GlobalVariables.NebSlavehwnd.ToString());
                    StatusHandleRec = true;
                }
                /*//*************rem'd 6-9
                //  SetForegroundWindow(NebhWnd);
                if ((IsServer()) & (Slavehwnd != 0))
                {

                    SlaveCapture();

                }
                 */
                //   watchforOpenPort();
                //   SetForegroundWindow(NebhWnd);
                //if (GlobalVariables.Nebcamera == "No camera")
                //    NoCameraSelected();
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

                /*
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
                 */
                TotalSubs();
                //totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + (int)numericUpDown32.Value;
                ToolStripProgressBar();

                // subsperfilter = totalsubs / filternumber;
                // textBox19.Text = subsperfilter.ToString();
                if (CountFilterTotal() == 0)
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
                    if (Filter.DevId3 != null)
                        Filter.filterWheel.Position = (short)comboBox2.SelectedIndex;//move to selected filter
                    DisplayCurrentFilter();
                    textBox21.BackColor = System.Drawing.Color.White;
                    Filtertext = comboBox2.Text;
                    textBox21.Text = Filtertext;
                    if (checkBox17.Checked == false)//refocus after N subs
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
                        // filteradvance();
                        if (!checkBox31.Checked)
                            Filter.filterWheel.Position = (short)comboBox3.SelectedIndex;
                        DisplayCurrentFilter();
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
                            //  filteradvance();
                            //Thread.Sleep(4000);
                            //  filteradvance();

                            currentfilter = 3;
                            if (!checkBox31.Checked)
                                Filter.filterWheel.Position = (short)comboBox4.SelectedIndex;
                            DisplayCurrentFilter();
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
                                //  filteradvance();
                                //   filteradvance();
                                //   filteradvance();
                                if (!checkBox31.Checked)
                                    Filter.filterWheel.Position = (short)comboBox5.SelectedIndex;
                                DisplayCurrentFilter();
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
                                    if (comboBox1.SelectedItem.ToString() != "Dark 1")
                                    {
                                        currentfilter = 5;
                                        if ((!checkBox31.Checked) && (Filter.DevId3 != null))
                                            Filter.filterWheel.Position = (short)comboBox1.SelectedIndex;
                                        DisplayCurrentFilter();
                                        if (checkBox17.Checked == false)
                                            subsperfilter = (int)numericUpDown14.Value;
                                        else

                                            subsperfilter = SubsPerFocus[2];

                                        Nebname = comboBox1.Text.ToString();
                                        CaptureTime = (int)numericUpDown17.Value * 1000;
                                        CaptureBin = (int)numericUpDown26.Value;
                                    }
                                    if (comboBox1.SelectedItem.ToString() == "Dark 1")
                                    // if (checkBox5.Checked == true)
                                    {
                                        DarksOn = true;
                                        // currentfilter = 5;
                                        subsperfilter = (int)numericUpDown20.Value;

                                        CaptureTime = (int)numericUpDown19.Value * 1000;
                                        CaptureBin = (int)numericUpDown28.Value;
                                        Nebname = "Dark1_" + CaptureTime.ToString();

                                    }
                                }
                                else
                                {
                                    /*
                                    if (checkBox9.Checked == true)
                                    {
                                        MessageBox.Show("Use Dark1 for single dark frame", "scopefocus");
                                        return;
                                    }

                                    else
                                    {
                                     */
                                    if ((checkBox13.Checked == true) & (checkBox18.Checked == false))//13 flat in sequence, 18 is flat every filter
                                    {
                                        FlatsOn = true;
                                        // currentfilter = 5;
                                        subsperfilter = (int)numericUpDown32.Value;
                                        Nebname = "Flat1";
                                        CaptureTime = (int)numericUpDown34.Value;
                                        CaptureBin = (int)numericUpDown36.Value;
                                    }

                                    // }
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
                fileSystemWatcher4.EnableRaisingEvents = true;
                sequenceRunning = true;
                DisableUpDwn();
                //    fileSystemWatcher4.EnableRaisingEvents = true;
                checkfiltercount();
            }
            catch (Exception ex)
            {
                Log("SequenceGo Error" + ex.ToString());
                Send("SequenceGo Error" + ex.ToString());
                FileLog("SequenceGo Error" + ex.ToString());

            }


        }
        private void button26_Click(object sender, EventArgs e)
        {
            if (paused == true)//*************unfinished *******
            {
                button22.PerformClick();
                return;
            }
            SequenceGo();

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
            textBox23.Text = TotalSubs().ToString();



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
            textBox23.Text = TotalSubs().ToString();
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
        /*
        //filter pos sync to 1
        private void button28_Click_1(object sender, EventArgs e)
        {
            
            try
            {
               
                

                if (checkBox16.Checked == false)
                {
                
                toolStripStatusLabel1.Text = "Filter Moving";
                this.Refresh();
               // focuser.Move(2);
                focuser.Action("FilterSync", "");
               // Thread.Sleep(200);
                    //**********consider re-implamenting a filter moving check*******
              //      toolStripStatusLabel1.Text = "Filter Moving";
              //      this.Refresh();
              //      string movedone;
              // //     port.DiscardOutBuffer();
              // //     port.DiscardInBuffer();
              //  //    Thread.Sleep(20);
              ////      port.Write("Z");
              //      filterMoving = true;
              //      Thread.Sleep(50);

              //      while (filterMoving == true)
              //      {
              //          movedone = port.ReadExisting();
              //          if (movedone == "D")
              //          {
              //              filterMoving = false;
              //              this.Refresh();
              //              toolStripStatusLabel1.Text = "Filter Sync Complete";
              //              //   toolStripStatusLabel1.Text = " ";
              //              //    toolStripStatusLabel1.Text = " ";

              //          }
              //      }

              //      Thread.Sleep(50);
              //      port.DiscardInBuffer();
                }
                this.Refresh();
                toolStripStatusLabel1.Text = "Filter Sync Complete";
                Filtertext = comboBox2.Text;
                currentfilter = 1;
                DisplayCurrentFilter();
                filtersynced = true;
                button28.BackColor = System.Drawing.Color.Lime;
               // gotopos(2);
              

            }

            catch (Exception ex)
            {
                Log("Filter Position Sync Error" + ex.ToString());
            }

        }
        */
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



        private void textBox4_TextChanged(object sender, EventArgs e)
        {

            // posMin = Convert.ToInt32(textBox4.Text);//allows direct entry of close focus position
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
                int metricHFR;
                metricpath = Directory.GetFiles(GlobalVariables.Path2.ToString(), "metric*.fit");
                metricHFR = GetMetric(metricpath, roundto);
                textBox25.Text = metricHFR.ToString();

                AvgMetricHFR[currentmetricN - 1] = metricHFR;

                if (currentmetricN == metricN)
                {
                    int AvgMetric;
                    MetricSample = false;
                    fileSystemWatcher5.EnableRaisingEvents = false;
                    AvgMetric = (AvgMetricHFR.Sum()) / metricN;
                    textBox25.Text = AvgMetric.ToString();
                    Log("Avg Metric HFR= " + AvgMetric.ToString() + " N =  " + metricN.ToString());
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
                    //  SetForegroundWindow(Handles.Aborthwnd);
                    //   PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                    if (UseClipBoard.Checked)
                    {
                        // delay(1);
                        Clipboard.Clear();
                        //delay(1);
                        //SetForegroundWindow(Handles.NebhWnd);
                        ////  Thread.Sleep(1000);
                        //delay(1);
                        //PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                        //  Thread.Sleep(1000);

                        msdelay(750);
                        //Clipboard.SetText("//NEB Listen 0");
                        //msdelay(750);
                    }
                    else
                    {
                        serverStream = clientSocket.GetStream();
                        byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                        serverStream.Write(outStream, 0, outStream.Length);
                        serverStream.Flush();

                        Thread.Sleep(3000);
                        serverStream.Close();

                        SetForegroundWindow(Handles.NebhWnd);
                        Thread.Sleep(1000);
                        PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                        Thread.Sleep(1000);
                        //      NebListenOn = false;
                        // clientSocket.GetStream().Close();//added 5-17-12
                        //  clientSocket.Client.Disconnect(true);//added 5-17-12
                        clientSocket.Close();
                        //  Thread.Sleep(500);
                        //   SendKeys.SendWait("~");
                        //   SendKeys.Flush();
                    }
                    if (metricpath != null)
                        File.Delete(metricpath[0]);
                    currentmetricN = 0;
                    if ((autoMetricVcurve == true) & (FilterFocusOn == false))
                        button11.PerformClick();
                    if ((autoMetricVcurve == true) & (FilterFocusOn == true))
                        gotoFocus();
                    return;
                }

            }
            if (currentmetricN != vN - 1)
            {
                currentmetricN++;
                MetricCapture();
            }
            //if (currentmetricN == vN -1)
            //    currentmetricN = 0;
            //if (vProgress == vN)
            //{
            //    serverStream = clientSocket.GetStream();
            //    byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
            //    serverStream.Write(outStream, 0, outStream.Length);
            //    serverStream.Flush();

            //    Thread.Sleep(3000);
            //    serverStream.Close();
            //    SetForegroundWindow(Handles.NebhWnd);
            //    Thread.Sleep(1000);
            //    PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
            //    Thread.Sleep(1000);
            //    NebListenOn = false;
            //    // clientSocket.GetStream().Close();//added 5-17-12
            //    //  clientSocket.Client.Disconnect(true);//added 5-17-12
            //    clientSocket.Close();
            //    //  Thread.Sleep(500);
            //    //   SendKeys.SendWait("~");
            //    //   SendKeys.Flush();

            //    if (metricpath != null)
            //        File.Delete(metricpath[0]);
            //    currentmetricN = 0;
            //    return;
            //}


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

        // ********removed 11-3-16
        //metric Go button,  start N captures Fine V
        //private void button25_Click(object sender, EventArgs e)
        //{
        //    if (checkBox22.Checked == false)
        //        checkBox22.Checked = true;
        //    currentmetricN = 0;
        //    /*
        //    if (roughvdone == false)//make sure rough v curve done to establish rough focus point
        //    {
        //        DialogResult result2;
        //        result2 = MessageBox.Show("Must perform a rough V-curve first", "scopefocus",
        //            MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //        if (result2 == DialogResult.OK)
        //        {
        //            return;
        //        }
        //        return;
        //    }
        //     */
        //    button3.PerformClick();
        //    /*
        //    if (clientSocket.Connected == false)
        //    {
        //        //clientSocket.Connect("127.0.0.1", SocketPort);//connects to neb
        //        NebListenStart(NebhWnd, SocketPort);
        //    }
        //     */
        //    fileSystemWatcher5.EnableRaisingEvents = true;
        //    MetricCapture();

        //}




        //metriccapturehere
        private void MetricCapture()
        {

            try
            {
                MetricTime = Convert.ToDouble(textBox43.Text);
                //    if (clientSocket.Connected == false)
                //       NebListenStart(Handles.NebhWnd, SocketPort);//remd 7-1-14


                //  NetworkStream serverStream = clientSocket.GetStream();



                // added 10-11-16


                if (UseClipBoard.Checked) // 11-8-16
                {
                  //  int i = 0;
                    NebListenStart(Handles.NebhWnd, SocketPort);
                    delay(1);
                    Clipboard.Clear();
                    msdelay(500);
                    //   Clipboard.SetText("//NEB SetDuration " + MetricTime);
                    Clipboard.SetDataObject("//NEB SetDuration " + MetricTime, false, 3, 500);
                  msdelay(750);
                    //if (!NebCommandConfirm("SetDuration " + MetricTime, 0))
                    //    if (!clipboardRetry)
                    //    {
                    //        clipboardRetry = true;
                    //        MetricCapture();
                    //    }
                    //    else
                    //        ClipboardFailure("SetDuration");
                    //  NebCommandConfirm("SetDuration " + MetricTime, 0);
                    //while (!NebCommandConfirm("SetDuration " + MetricTime, 0))
                    //{
                    //    i++;
                    //    msdelay(100);
                    //    if (i == 30)
                    //    {
                    //        Log("Metric Failed");
                    //        Send("Metric Failed");
                    //        return;
                    //    }
                    //}

                    // 11-21-16 added Bin
                    
                    Clipboard.SetDataObject("//NEB SetBinning " + numericUpDown1.Value.ToString() , false, 3, 500);
                    msdelay(750);



                    //i = 0;
                    //   Clipboard.SetText("//NEB CaptureSingle metric");
                    Clipboard.SetDataObject("//NEB CaptureSingle metric", false, 3, 500);
                    msdelay(750);
                    //if (!NebCommandConfirm("Exposing", 1))
                    //    if (!clipboardRetry)
                    //    {
                    //        clipboardRetry = true;
                    //         MetricCapture();
                    //    }
                    //    else
                    //        ClipboardFailure("Exposing");
                    //*** todo confirm 

                    //   NebCommandConfirm("//NEB CaptureSingle metric");  //***this doesn't happend until the exposure is done.
                    //while (!NebCommandConfirm("Exposure done",1))
                    //{
                    //    i++;
                    //    msdelay(100);
                    //    if (i == 30)
                    //    {
                    //        Log("Metric Capture Failed");
                    //        Send("Metric Capture Failed");
                    //        return;
                    //    }
                    //}
                    Clipboard.Clear();
                }
                else
                {
                    if (clientSocket.Connected == false)
                    {
                        NebListenStart(Handles.NebhWnd, SocketPort);
                        // clientSocket.Connect("127.0.0.1", SocketPort);//connects to neb
                    }
                    serverStream = clientSocket.GetStream();
                    byte[] outStream = System.Text.Encoding.ASCII.GetBytes("SetDuration " + MetricTime + "\n" + "setbinning " + numericUpDown1.Value.ToString() + "\n" + "CaptureSingle metric" + "\n");
                    serverStream.Write(outStream, 0, outStream.Length);
                    serverStream.Flush();
                }
            }
            catch (Exception ex)
            {
                Log("MetricCapture Error line 7492" + ex.ToString());
                Send("MetricCapture Error line 7492" + ex.ToString());
                FileLog("MetricCapture Error line 7492" + ex.ToString());
                return;

            }


        }

        private void checkBox22_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox22.Checked == true)
            {
                Filtertext = Filtertext + "Metric";


            }
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
                metricN = (int)numericUpDown21.Value;
                AvgMetricHFR = new int[metricN];
                MetricTime = Convert.ToDouble(textBox43.Text);
                currentmetricN = 1;
                MetricSample = true;

                fileSystemWatcher5.EnableRaisingEvents = true;
                // NetworkStream serverStream = clientSocket.GetStream();
                if (UseClipBoard.Checked) // 11-8-16
                {
                  //  int i = 0;
                    NebListenStart(Handles.NebhWnd, SocketPort);
                    delay(1);
                    Clipboard.Clear();
                    msdelay(500);
                    //  Clipboard.SetText("//NEB SetDuration 5");
                    Clipboard.SetDataObject("//NEB SetDuration " + MetricTime, false, 3, 500);
                    //    Thread.Sleep(500);
                    msdelay(750);
                    //if (!NebCommandConfirm("SetDuration " + MetricTime, 0))
                    //    if (!clipboardRetry)
                    //    {
                    //        clipboardRetry = true;
                    //        NebCapture();
                    //    }
                    //    else
                    //        SampleMetric();
                    //while (!NebCommandConfirm("SetDuration " + MetricTime, 0))
                    //{
                    //    i++;
                    //    msdelay(100);
                    //    if (i == 30)
                    //    {
                    //        Log("Metric Capture Failed");
                    //        Send("Metric Capture Failed");
                    //        return;
                    //    }
                    //}

                    // 11-21-16 add binning
                    Clipboard.SetDataObject("//NEB SetBinning " + numericUpDown1.Value.ToString(), false, 3, 500);
                    msdelay(750);

                    //i = 0;
                    //    Clipboard.Clear();
                    //   Clipboard.SetText("//NEB CaptureSingle metric");
                    Clipboard.SetDataObject("//NEB CaptureSingle metric", false, 3, 500);
                    msdelay(750);

                    // todo confirm
                    //if (!NebCommandConfirm("Exposing", 1))
                    //    if (!clipboardRetry)
                    //    {
                    //        clipboardRetry = true;
                    //        SampleMetric();
                    //    }
                    //    else
                    //        ClipboardFailure("Exposing");


                    //while (!NebCommandConfirm("Exposure done", 1)) // **** remove this if don't want UI unaccessible. 
                    //{
                    //    i++;
                    //    msdelay(100);
                    //    if (i == 30)
                    //    {
                    //        Log("Metric Capture Failed");
                    //        Send("Metric Capture Failed");
                    //        return;
                    //    }
                    //}

                    //    Thread.Sleep(500);
                    //   delay(1);
                    Clipboard.Clear();
                }
                else
                {
                    if (clientSocket.Connected == false) // moved 11-9-16
                    {
                        NebListenStart(Handles.NebhWnd, SocketPort);
                        // clientSocket.Connect("127.0.0.1", SocketPort);//connects to neb
                    }
                    serverStream = clientSocket.GetStream();
                    byte[] outStream = System.Text.Encoding.ASCII.GetBytes("SetDuration " + MetricTime + "\n" + "setbinning " + numericUpDown1.Value.ToString() + "\n" + "CaptureSingle metric" + "\n");
                    serverStream.Write(outStream, 0, outStream.Length);

                    serverStream.Flush();
                }


            }

            catch (ObjectDisposedException e)
            {
                // MessageBox.Show("Error Communicating with Nebulosity, Make Sure it's open and click Rescan on Setup tab", "scopefocus");
                Log("Error Communicating with Nebulosity");
                Send("Error Communicating with Nebulosity");
                return;
            }

            catch (Exception ex)
            {
                Log("SampleMetric Error Line 8301" + ex.ToString());
                Send("Error Communicating with Nebulosity");
                return;
            }

        }


        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void button36_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    if (GlobalVariables.Portselected2 == null)
            //    {
            //        MessageBox.Show("Port not selected", "scopefocus");
            //        return;
            //    }
            //    if (port2 == null)
            //    {
            //        Port2Open();
            //        Thread.Sleep(200);
            //    }
            //    if (connect2 == 1)
            //    {
            //        DialogResult result2 = MessageBox.Show("Already Connected", "Arduino scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //    }
            //    if (port2.IsOpen == true)
            //    {
            //        Thread.Sleep(500);
            //        CheckNexremoteConnection();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Log("Port 2 Error  Line 4330" + ex.ToString());
            //}


        }


        //private void CheckNexremoteConnection()
        //{
        //    try
        //    {
        //        if (port2.IsOpen == false)
        //            Port2Open();
        //        string CommCheck;
        //        port2.DiscardInBuffer();
        //        port2.DiscardOutBuffer();
        //        Thread.Sleep(50);
        //        port2.Write("Kx");//K is echo command x is an arbitrary char
        //        Thread.Sleep(100);
        //        port2.DiscardOutBuffer();

        //        Thread.Sleep(50);
        //        CommCheck = port2.ReadExisting();
        //        Thread.Sleep(50);
        //        if (CommCheck == "x#") //check for echo of x and stop bit
        //        {
        //            connect2 = 1;
        //            button36.Text = "Mount Connected";
        //            button36.BackColor = System.Drawing.Color.Lime;
        //            port2.DiscardOutBuffer();
        //            port2.DiscardInBuffer();
        //            Log("Connected to Mount on " + GlobalVariables.Portselected2.ToString());

        //        }
        //        else
        //            MessageBox.Show("Nexremote Connection Failed", "scopefocus");

        //        port2.Close();
        //    }
        //    catch (Exception ex)
        //    {
        //        Log("Check Nexremote Connection Error" + ex.ToString());
        //    }
        //}


        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {



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
            double DecTotalSec = (DecDegDec * 256 * 256 + DecMinDec * 256 + DecSecDec) / 12 / 1.078781893;
            deg = ((int)DecTotalSec) / 60 / 60;
            min = ((int)DecTotalSec - (deg * 3600)) / 60;
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


        public double FocusRA;  //for using ASCOM scope
        public double FocusDEC;
        public double TargetRA;
        public double TargetDEC;
        public int RAhr;
        public int RAmin;
        public double RAsec;
        public int DecDeg;
        public int DecMin;
        public int DecSec;


        private void GetSolvedFocusLocation()
        {
            //plate solve the focus image and return FocusRA and FocusDEC


        }

        bool SettingFocusSolve = false;
        bool SettingTargetSolve = false;
        private async void GetFocusLocation() //via scope RA and DEC or prev image or new solve.   //11-9-16 changed to async
        {

            try
            {
                if (checkBox24.Checked) //if use image
                {
                    if (textBox50.Text != "") //use prev saved location image
                    {
                        //  GlobalVariables.FocusImage = file.Name.ToString();
                        GlobalVariables.SolveImage = GlobalVariables.FocusImage;
                        button32.Text = "Solving";
                        button32.BackColor = System.Drawing.Color.Yellow;
                        //  textBox50.Text = GlobalVariables.SolveImage;
                        //  PartialSolve();
                        SettingFocusSolve = true;
                        if (GlobalVariables.LocalPlateSolve)
                            Solve();
                        else
                        {
                            //if (backgroundWorker1.IsBusy != true)
                            //{
                            //    // Start the asynchronous operation.
                            //    backgroundWorker1.RunWorkerAsync();
                            //}
                            Log("Attempt Plate Solve: " + GlobalVariables.SolveImage);
                            FileLog2("Plate Solve: " + GlobalVariables.SolveImage);
                            solveRequested = true;
                            AstrometryRunning = true;
                            // try
                            // {
                            AstrometryNet ast = new AstrometryNet();  // remd for onlyinstance
                                                                      // ast = AstrometryNet.OnlyInstance;
                                                                      //  AstrometryNet.GetForm.Show();
                            await ast.OnlineSolve(GlobalVariables.SolveImage);
                        }

                        // remd 11-9-16
                        //FocusDEC = DEC;
                        //FocusRA = RA;
                        ////Log("Solved Focus Location:  RA = " + FocusRA.ToString() + "   DEC = " + FocusDEC.ToString());
                        ////FileLog2("Solved Focus Location:  RA = " + FocusRA.ToString() + "   DEC = " + FocusDEC.ToString());
                        //button32.Text = "Solved"; // change the set button only...since not necessarily there (could be from prev session) 
                        //button32.BackColor = System.Drawing.Color.Lime;
                    }

                    //manually take neb image w/ FSwatcher 7 on, solve, save FocusRA/DEC
                    else //if need to take image 
                    {
                        button32.Text = "Waiting";
                        button32.BackColor = Color.Yellow;
                        SettingFocusSolve = true;
                        fileSystemWatcher7.Path = GlobalVariables.Path2;
                        fileSystemWatcher7.SynchronizingObject = this;
                        fileSystemWatcher7.EnableRaisingEvents = true;

                    }

                    //remd 11-9-16
                    //Log("Solved Focus Location:  RA = " + FocusRA.ToString() + "   DEC = " + FocusDEC.ToString());
                    //FileLog2("Solved Focus Location:  RA = " + FocusRA.ToString() + "   DEC = " + FocusDEC.ToString());



                    //button32.Text = "Solved"; // change the set button only...since not necessarily there (could be from prev session) 
                    //button32.BackColor = System.Drawing.Color.Yellow;
                }
                else
                {
                    // use mount position.  

                    FocusRA = scope.RightAscension;
                    FocusDEC = scope.Declination;
                    Log("ASCOM Focus Position Set: RA = " + FocusRA.ToString() + "DEC = " + FocusDEC.ToString());
                    FileLog2("ASCOM Focus Position Set: RA = " + FocusRA.ToString() + "DEC = " + FocusDEC.ToString());
                    button33.Text = "At Focus";
                    button33.BackColor = System.Drawing.Color.Lime;
                    button35.Text = "Goto Target";
                    button35.UseVisualStyleBackColor = true;
                    button32.BackColor = System.Drawing.Color.Lime;
                    button32.Text = "Focus Pos Set";

                }
                FocusLocObtained = true;

                //string path = textBox11.Text.ToString();
                //string fullpath = path + @"\log.txt";
                //StreamWriter log;
                //string string4 = "Focus Position Set: RA = " + FocusRA.ToString() + "DEC = " + FocusDEC.ToString();
                //FileLog2(string4);
                //if (!File.Exists(fullpath))
                //{
                //    log = new StreamWriter(fullpath);
                //}
                //else
                //{
                //    log = File.AppendText(fullpath);
                //}
                //log.WriteLine(string4);
                //log.Close();
                //  }

                //   else
                //  {
                //
                //if (port2 == null)
                //{
                //    MessageBox.Show("Not Connected to Nexremote", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    return;
                //}
                //if (port2.IsOpen == false)
                //    Port2Open();

                //port2.DiscardOutBuffer();
                //port2.DiscardInBuffer();
                //Thread.Sleep(20);
                //port2.Write("e");
                //Thread.Sleep(100);
                //FocusLoc = port2.ReadExisting();
                //ConvertDec(FocusLoc, out DecDeg, out DecMin, out DecSec);
                //ConvertRA(FocusLoc, out RAhr, out RAmin, out RAsec);
                //Log("Focus Position Set: RA " + RAhr.ToString() + "hr " + RAmin.ToString() + "m " + RAsec.ToString() + "s " + "- Dec " + DecDeg.ToString() + "° " + DecMin.ToString() + "' " + DecSec.ToString() + "\"");
                //Thread.Sleep(100);
                //port2.DiscardInBuffer();
                //Thread.Sleep(10);
                //if (FocusLoc.Substring(17, 1) == "#")//check for stop bit
                //{
                //    FocusLocObtained = true;
                //    button32.BackColor = System.Drawing.Color.Lime;
                //    button32.Text = "Focus Pos Set";
                //}
                //string path = textBox11.Text.ToString();
                //string fullpath = path + @"\log.txt";
                //StreamWriter log;
                //string string4 = "Focus Position Set: RA " + RAhr.ToString() + "hr " + RAmin.ToString() + "m " + RAsec.ToString() + "s " + "- Dec " + DecDeg.ToString() + "° " + DecMin.ToString() + "' " + DecSec.ToString() + "\"";
                //if (!File.Exists(fullpath))
                //{
                //    log = new StreamWriter(fullpath);
                //}
                //else
                //{
                //    log = File.AppendText(fullpath);
                //}
                //log.WriteLine(string4);
                //log.Close();
                //port2.Close();
                // }
            }
            catch (Exception ex)
            {
                Log("GetFocus Location Error" + ex.ToString());
            }

        }
        //get focus location
        private void button32_Click(object sender, EventArgs e)
        {

            GetFocusLocation();
        }





        private async void GetTargetLocation()
        {
            try
            {
                if (checkBox32.Checked) //if use image
                {

                    if (textBox51.Text != "") //use prev saved location image
                    {
                        GlobalVariables.SolveImage = GlobalVariables.TargetImage;
                        button34.Text = "Solving"; // change the set button only...since not necessarily there (could be from prev session) 
                        button34.BackColor = System.Drawing.Color.Yellow;
                        //  textBox50.Text = GlobalVariables.SolveImage;
                        // PartialSolve();  changed to below 11-9-16
                        SettingTargetSolve = true;

                        if (GlobalVariables.LocalPlateSolve)
                            Solve();
                        else
                        {
                            //if (backgroundWorker1.IsBusy != true)
                            //{
                            //    // Start the asynchronous operation.
                            //    backgroundWorker1.RunWorkerAsync();
                            //}
                            Log("Attempt Plate Solve: " + GlobalVariables.SolveImage);
                            FileLog2("Plate Solve: " + GlobalVariables.SolveImage);
                            solveRequested = true;
                            AstrometryRunning = true;
                            // try
                            // {
                            AstrometryNet ast = new AstrometryNet();  // remd for onlyinstance
                                                                      // ast = AstrometryNet.OnlyInstance;
                                                                      //  AstrometryNet.GetForm.Show();
                            await ast.OnlineSolve(GlobalVariables.SolveImage);
                        }

                        //remd 11-9-16
                        //TargetDEC = DEC;
                        //TargetRA = RA;
                        ////Log("Solved Target Location:  RA = " + TargetRA.ToString() + "   DEC = " + TargetDEC.ToString());
                        ////FileLog2("Solved Target Location:  RA = " + TargetRA.ToString() + "   DEC = " + TargetDEC.ToString());
                        //button34.Text = "Solved"; // change the set button only...since not necessarily there (could be from prev session) 
                        //button34.BackColor = System.Drawing.Color.Lime;
                    }
                    else //if need to take image 
                    {
                        button34.Text = "Waiting";
                        button34.BackColor = Color.Yellow;
                        SettingTargetSolve = true;
                        fileSystemWatcher7.Path = GlobalVariables.Path2;
                        fileSystemWatcher7.SynchronizingObject = this;
                        fileSystemWatcher7.EnableRaisingEvents = true;
                        Log("Attemping to plate solve Nebulosity capture");
                        FileLog2("Plate Solve Start - Nebulosity Capture");

                    }
                    //manually take neb image w/ FSwatcher 7 on, solve, save FocusRA/DEC
                    //fileSystemWatcher7.EnableRaisingEvents = true;
                    //GlobalVariables.SolveImage = GlobalVariables.TargetImage;
                    //PartialSolve();
                    //TargetDEC = DEC;
                    //TargetRA = RA;


                    TargetLocObtained = true;

                    //remd 11-916
                    //Log("Solved Target Location:  RA = " + TargetRA.ToString() + "   DEC = " + TargetDEC.ToString());
                    //FileLog2("Solved Target Location:  RA = " + TargetRA.ToString() + "   DEC = " + TargetDEC.ToString());
                }
                else
                {

                    // use mount position.  

                    TargetRA = scope.RightAscension;
                    TargetDEC = scope.Declination;
                    Log("ASCOM Target Position Set: RA = " + TargetRA.ToString() + "   DEC = " + TargetDEC.ToString());
                    FileLog2("ASCOM Target Position Set: RA = " + TargetRA.ToString() + "   DEC = " + TargetDEC.ToString());

                    TargetLocObtained = true;
                    button34.BackColor = System.Drawing.Color.Lime;
                    button34.Text = "Target Pos Set";
                    button33.Text = "Goto"; //since at target make goto focus button back to normal
                    button33.UseVisualStyleBackColor = true;
                    button35.Text = "At Target";
                    button35.BackColor = Color.Lime;
                }



                //string path = textBox11.Text.ToString();
                //string fullpath = path + @"\log.txt";
                //StreamWriter log;
                //  string string4 = "ASCOM Target Position Set: RA = " + TargetRA.ToString() + "  DEC = " + TargetDEC.ToString();
                //if (!File.Exists(fullpath))
                //{
                //    log = new StreamWriter(fullpath);
                //}
                //else
                //{
                //    log = File.AppendText(fullpath);
                //}
                //  log.WriteLine(string4);
                //  FileLog2(string4);
                //log.Close();
                //    }

                //else
                //{
                //if (port2.IsOpen == false)
                //    Port2Open();
                //if (port2 == null)
                //{
                //    MessageBox.Show("Not Connected to Nexremote", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    return;
                //}
                //port2.DiscardOutBuffer();
                //port2.DiscardInBuffer();
                //Thread.Sleep(20);
                //port2.Write("e");
                //Thread.Sleep(100);
                //TargetLoc = port2.ReadExisting();
                ////  Thread.Sleep(50);
                //port2.DiscardInBuffer();
                ////  textBox28.Clear();
                //Log("targegloc " + TargetLoc.ToString());
                //Thread.Sleep(100);
                //ConvertDec(TargetLoc, out DecDeg, out DecMin, out DecSec);
                //ConvertRA(TargetLoc, out RAhr, out RAmin, out RAsec);
                //Log("Target Position Set: RA " + RAhr.ToString() + "hr " + RAmin.ToString() + "m " + RAsec.ToString() + "s " + "- Dec " + DecDeg.ToString() + "° " + DecMin.ToString() + "' " + DecSec.ToString() + "\"");

                ////     textBox28.Text = TargetLoc.ToString();
                //if (TargetLoc.Substring(17, 1) == "#")//check for stop bit
                //{
                //    TargetLocObtained = true;
                //    button34.BackColor = System.Drawing.Color.Lime;
                //    button34.Text = "Target Pos Set";
                //}
                //string path = textBox11.Text.ToString();
                //string fullpath = path + @"\log.txt";
                //StreamWriter log;
                //string string4 = "Target Position Set: RA " + RAhr.ToString() + "hr " + RAmin.ToString() + "m " + RAsec.ToString() + "s " + "- Dec " + DecDeg.ToString() + "° " + DecMin.ToString() + "' " + DecSec.ToString() + "\"";
                //if (!File.Exists(fullpath))
                //{
                //    log = new StreamWriter(fullpath);
                //}
                //else
                //{
                //    log = File.AppendText(fullpath);
                //}
                //log.WriteLine(string4);
                //log.Close();
                //port2.Close();
                //   }

            }
            catch (Exception ex)
            {
                Log("GetTargetLocation Error" + ex.ToString());

            }
        }



        //    private bool TrackingOn = true;
        //   private int retry = 0;
        /*
        private void StopTracking()
        {
           // try
          //  {

              //  if (TrackingOn == true)
             //   {
                  //  if (usingASCOM)
                  //  {
                        scope.Tracking = false;
                //        TrackingOn = false;
             //   }
                     //   return;//not sure if needed
                  //  }
                  //  else
                  //  {
                        //if (port2.IsOpen == false)
                        //    Port2Open();
                        //port2.DiscardInBuffer();
                        //port2.DiscardOutBuffer();
                        //Thread.Sleep(50);
                        //byte[] newMsg = { 0x54, 0x00 };
                        //port2.Write(newMsg, 0, newMsg.Length);
                        //Thread.Sleep(100);
                        //port2.DiscardOutBuffer();

                        //Thread.Sleep(50);
                        //string Handshake = port2.ReadExisting();
                        //Thread.Sleep(50);
                        //if (Handshake == "#") //check for handshake
                        //{
                        //    Log("Tracking Command Received");
                        //}

                        //port2.DiscardInBuffer();
                        //port2.DiscardOutBuffer();
                        //Thread.Sleep(50);
                        //port2.Write("t");//check tracking mode
                        //Thread.Sleep(100);
                        //port2.DiscardOutBuffer();
                        //// byte[] confirm = new byte[1];
                        //int confirm;
                        //confirm = port2.ReadByte();
                        //Thread.Sleep(20);
                        //if (confirm == 0)
                        //{
                        //    Log("Tracking Off Confirmed");
                        //    TrackingOn = false;
                        //}
                        //else
                        //{
                        //    if (retry < 3)
                        //    {
                        //        retry++;
                        //        Thread.Sleep(1000);
                        //        StopTracking();
                        //    }
                        //    else
                        //    {
                        //        retry = 0;
                        //        MessageBox.Show("Unable to stop tracking", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        //        //*****************eventually add text message here ******************************

                        //    }
                      //  }
                  //  }
               // }
              //  else
              //      MessageBox.Show("Tracking is Already Off", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
             //   port2.Close();
         //   }
            //catch (Exception ex)
            //{
            //    Log("Stop Tracking Error" + ex.ToString());
            //    Send("Stop Tracking Error" + ex.ToString());
            //    FileLog("Stop Tracking Error" + ex.ToString());

            //}
        }
        */
        //Get target location 
        private void button34_Click(object sender, EventArgs e)
        {
            //button35.Text = "At Target";
            //button35.BackColor = System.Drawing.Color.Lime;
            //button33.Text = "Goto";
            //button33.UseVisualStyleBackColor = true;
            GetTargetLocation();

        }
        private bool FocusGotoOn = false;

        bool ConfirmingLocation = false;
        //goto focus
        int solvetry = 0;

        private void GotoFocusLocation()
        {
            try
            {
                /// ***** 4-12 16
                /// add for plate solve...
                ///go to the previously stored FocusRA and FocusDEC then take and image 
                ///plate solve the image
                ///use tolerance from platesolve tab, to ensure is there...if not, sync then take another image and repeat until there.    

                FocusGotoOn = true;

                if (FocusLocObtained == false)
                {
                    MessageBox.Show("Focus Position Not Set", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                //bool PHDpaused = false;

                //if (PHDcommand(17) == "3") //pause guiding if active and gotfocus from button push
                //{
                //    PHDcommand(PHD_PAUSE);
                //    Log("guiding paused for slew");
                //    PHDpaused = true;
                //}  
                StopPHD();
                //   if (usingASCOM)
                // scope.SlewToCoordinates(FocusRA, FocusDEC);
                scope.SlewToCoordinatesAsync(FocusRA, FocusDEC);
                Log("slewing to focus star");
                toolStripStatusLabel1.Text = "Slewing";
                toolStripStatusLabel1.BackColor = Color.Red;
                this.Refresh();
                while (scope.Slewing)
                {

                    Thread.Sleep(50); // pause for 1/20 second
                    System.Windows.Forms.Application.DoEvents();//this makes it necssary to push twice!!!????
                }
                if (checkBox16.Checked == false)//dont resume if guide on focus disabled
                {
                    //if (Handles.PHDVNumber == 2)
                    Log("Guide at focus postion disabled");

                    //else
                    //    ResumePHD();
                }
                else
                    resumePHD2();
                // add 4-12 16

                // remd 4-14-16  confirming each goto focus slew is a problem since sends focus time before confimration is comple
                //if (checkBox24.Checked)
                //{
                //    toolStripStatusLabel1.Text = "Soving";
                //    toolStripStatusLabel1.BackColor = Color.Lime;
                //    this.Refresh();
                //    button33.Text = "Pending"; // change the set button only...since not necessarily there (could be from prev session) 
                //    button33.BackColor = System.Drawing.Color.Yellow;
                //    Log("Confirming Focus plate solve");
                //    ConfirmingLocation = true;
                //    fileSystemWatcher7.Path = GlobalVariables.Path2;
                //    fileSystemWatcher7.SynchronizingObject = this;
                //    if (fileSystemWatcher7.EnableRaisingEvents == false)
                //        fileSystemWatcher7.EnableRaisingEvents = true;

                //    ///  ****   4-12-16  todo send neb command to caputre 1 image at exp time per textbox65.  
                //    ///  cant just push cpautre button since the setting wont necessarily be correct if doing focus in middle of sequence.  


                //    SolveCapture();
                //}

                ///  ********** need to wait until capture AND solve done before continuing.            






                ////SetForegroundWindow(Handles.NebhWnd);  //rem'd to testing 
                ////PostMessage(Handles.CaptureMainhWnd, BN_CLICKED, 0, 0);//rem'd for testing



                ////if (checkBox25.Checked == true)//repeat until tolerance met
                ////{
                ////  Thread.Sleep(5000);
                ////     Log("calibrating");
                ////scope.SlewToCoordinates(RA, DEC);
                //CurrentRA = scope.RightAscension;
                //CurrentDEC = scope.Declination;
                //// first attempts at comparing parse solve coords w/ scope coords.
                ////need to FIX...seems to maintain RA and DEC from first plate solve after the second one. 
                ////if (usingASCOM == true)
                ////{
                ////    //  scope = new ASCOM.DriverAccess.Telescope(devId);
                ////    scope.SlewToCoordinates(CurrentRA, CurrentDEC);//go back to where originally though it was supposed to be
                ////                                                   //should be closer after the sync
                ////}
                ////else
                ////    MessageBox.Show("Must use ASCOM mount connection", "scopefocus");

                //if ((Math.Abs(CurrentRA - RA) * 60 > Convert.ToDouble(textBox59.Text)) || (Math.Abs(CurrentDEC - DEC) * 60 > Convert.ToDouble(textBox59.Text)))//*************untested!!****
                //{
                //    scope.SyncToCoordinates(RA, DEC);//sync to parsed(solve) location 
                //    Log("repeating" + "DeltaRA = " + ((Math.Abs(CurrentRA - RA) * 60).ToString()) + "     DeltaDEC = " + ((Math.Abs(CurrentDEC - DEC) * 60).ToString()));
                //    if (solvetry < 3)
                //    {
                //        solvetry++;
                //        SolveCapture();
                //    }
                //    if (solvetry == 3)
                //    {
                //        Log("Solve Confirmation failed, using mount coordinates");

                //    }
                //    //     button55.PerformClick();//prob dont need since fsw7 still on
                //    //if (fileSystemWatcher7.EnableRaisingEvents == false)
                //    //    fileSystemWatcher7.EnableRaisingEvents = true;
                //    //SetForegroundWindow(Handles.NebhWnd);  //rem'd to testing 
                //    //PostMessage(Handles.CaptureMainhWnd, BN_CLICKED, 0, 0);//rem'd for testing
                //}

                //else
                //{
                //    Log("Confimred: At Focus Location");
                //    //solvetry = 0;
                //    //ConfirmingLocation = false;
                //}

                //       }
                //else  // if not using plate solve    remd w/ 4-14-16 rem above.
                //{
                Log("at focus star");
                toolStripStatusLabel1.Text = "Ready";
                toolStripStatusLabel1.BackColor = Color.WhiteSmoke;
                button33.Text = "At Focus";
                button33.BackColor = Color.Lime;
                button35.Text = "Goto";
                button35.UseVisualStyleBackColor = true;
                if (checkBox16.Checked == true)//resume if guide on focus enabled
                {
                    //if (Handles.PHDVNumber == 2)
                    resumePHD2();
                    //else
                    //    ResumePHD();
                }
                else
                    Log("Guide at focus postion disabled");

                // ****   4-12 16 begin addition to use plate solve confirmation


                //  }
            }

            // *************   need to reset confirminglocation and solvetry at some point.......

            catch (Exception ex)
            {






                Log("GotoFocus Location Error" + ex.ToString());
                Send("GotoFocus Location Error" + ex.ToString());
                FileLog("GotoFocus Location Error" + ex.ToString());

            }

        }


        private void SolveCapture()
        {
            NebListenStart(Handles.NebhWnd, SocketPort);

            //   }
            // Thread.Sleep(1000);//was 3000 5-29
            delay(1);
            if (FlatsOn == true)
                toolStripStatusLabel1.Text = "Capturing";
            //   string prefix = textBox19.Text.ToString();
            //   NetworkStream serverstream;
            string Solvetime = textBox65.Text;
            if (!UseClipBoard.Checked) // 11-8-16 and else below
            {
                try
                {
                    serverStream = clientSocket.GetStream();
                }
                catch (IOException e)
                {
                    Log("Neb socket failure " + e.ToString());
                    NebCapture();
                }
                byte[] outStream = System.Text.Encoding.ASCII.GetBytes("setname " + "Confirm_Solve_Location" + "\n" + "setbinning " + CaptureBin + "\n" + "SetShutter 0" + "\n" + "SetDuration " + Solvetime + "\n" + "Capture " + "1" + "\n");
                try
                {
                    serverStream.Write(outStream, 0, outStream.Length);
                    Log("Getting plate solve image confirmation");
                }
                catch
                {
                    MessageBox.Show("Error sending command", "scopefocus");
                    return;

                }
            }
            else
            {
               // int i = 0;
                delay(1);
                Clipboard.Clear();
                msdelay(500);
                // Clipboard.SetText("//NEB SetName " + "Confirm_Solve_Location");
                Clipboard.SetDataObject("//NEB SetName " + "Confirm_Solve_Location", false, 3, 500);
                msdelay(750);
                //if (!NebCommandConfirm("SetName " + "Confirm_Solve_Location", 0))
                //    if (!clipboardRetry)
                //    {
                //        clipboardRetry = true;
                //        SolveCapture();
                //    }
                //    else
                //        ClipboardFailure("SetName");
                //while (!NebCommandConfirm("SetName " + "Confirm_Solve_Location", 0))
                //{
                //    i++;
                //    msdelay(100);
                //    if (i == 30)
                //    {
                //        Log("Solve Capture Failed");
                //        Send("Solve Caputre Failed");
                //        return;
                //    }
                //}
                //i = 0;
                //   Clipboard.SetText("//NEB SetBinning " + CaptureBin);
                Clipboard.SetDataObject("//NEB SetBinning " + CaptureBin, false, 3, 500);
                msdelay(750);
                //if (!NebCommandConfirm("SetBinning " + CaptureBin, 0))
                //    if (!clipboardRetry)
                //    {
                //        clipboardRetry = true;
                //        SolveCapture();
                //    }
                //    else
                //        ClipboardFailure("SetBinning");
                //while (!NebCommandConfirm("SetBinning " + CaptureBin , 0))
                //{
                //    i++;
                //    msdelay(100);
                //    if (i == 30)
                //    {
                //        Log("Solve Capture Failed");
                //        Send("Solve Caputre Failed");
                //        return;
                //    }
                //}
                //i = 0;
                Clipboard.SetDataObject("//NEB SetShutter 0", false, 3, 500);
                //    Clipboard.SetText("//NEB SetShutter 0");
                msdelay(750);
                //if (!NebCommandConfirm("SetShutter 0", 0))
                //    if (!clipboardRetry)
                //    {
                //        clipboardRetry = true;
                //        SolveCapture();
                //    }
                //    else
                //        ClipboardFailure("SetShutter");
                //while (!NebCommandConfirm("SetShutter 0", 0))
                //{
                //    i++;
                //    msdelay(100);
                //    if (i == 30)
                //    {
                //        Log("Solve Capture Failed");
                //        Send("Solve Caputre Failed");
                //        return;
                //    }
                //}
                //i = 0;
                Clipboard.SetDataObject("//NEB SetDuration " + Solvetime, false, 3, 500);
                // Clipboard.SetText("//NEB SetDuration " + Solvetime);
                msdelay(750);
                //if (!NebCommandConfirm("SetDuration " + Solvetime, 0))
                //    if (!clipboardRetry)
                //    {
                //        clipboardRetry = true;
                //        SolveCapture();
                //    }
                //    else
                //        ClipboardFailure("setDuration");
                //while (!NebCommandConfirm("SetDuration " + Solvetime, 0))
                //{
                //    i++;
                //    msdelay(100);
                //    if (i == 30)
                //    {
                //        Log("Solve Capture Failed");
                //        Send("Solve Caputre Failed");
                //        return;
                //    }
                //}
                //i = 0;
                // Clipboard.SetText("//NEB Capture 1");
                Clipboard.SetDataObject("//NEB Capture 1", false, 3, 500);
                msdelay(750);
                //if (!NebCommandConfirm("Exposing", 1))
                //    if (!clipboardRetry)
                //    {
                //        clipboardRetry = true;

                //        SolveCapture();
                //    }
                //    else
                //        ClipboardFailure("Exposing");
                // todo confirm

                //while (!NebCommandConfirm("Exposure done", 1)) //holds UI until done. 
                //{
                //    i++;
                //    msdelay(100);
                //    if (i == 30)
                //    {
                //        Log("Solve Capture Failed");
                //        Send("Solve Caputre Failed");
                //        return;
                //    }
                //}
                //i = 0;
                Clipboard.Clear();
            }
            // add 4-13-16 try to hold progress until capture is done.  
            //if (backgroundWorker2.IsBusy != true)
            //    backgroundWorker2.RunWorkerAsync();
            //tryadd setforeground ***** 3-8-13  ****


        }



        //private void CheckForSlewDone()
        //{
        //    try
        //    {
        //        // ***************** this is doing nothing *************

        //        if (backgroundWorker1.IsBusy != true)
        //        {
        //            // Start the asynchronous operation.
        //            backgroundWorker1.RunWorkerAsync();
        //        }


        //        /*
        //        if (port2.IsOpen == false)
        //            port2.Open();
        //        while (MountMoving == true)
        //        {
        //            port2.DiscardInBuffer();
        //            port2.Write("L");
        //            Thread.Sleep(20);
        //            port2.DiscardOutBuffer();

        //          //  textBox13.Clear();
        //            Thread.Sleep(20);

        //            //  
        //            GotoDoneCommand = port2.ReadExisting();
        //            Thread.Sleep(10);
        //          //  port2.DiscardOutBuffer();
        //            port2.DiscardInBuffer();
        //           // textBox13.Text = GotoDoneCommand.ToString();
        //            Thread.Sleep(20);
        //            if (GotoDoneCommand == "0#")
        //            {
        //              //  textBox13.Text = "Goto Done";
        //                MountMoving = false;
        //                port2.DiscardInBuffer();
        //                port2.DiscardOutBuffer();
        //               // break;
        //                return;
        //            }
        //        }
        //       */
        //    }
        //    catch (Exception ex)
        //    {
        //        Log("CheckForSlewDone Error" + ex.ToString());
        //        Send("CheckForSlewDone Error" + ex.ToString());
        //        FileLog("CheckForSlewDone Error" + ex.ToString());

        //    }
        //}
        //goto focus location button
        private void button33_Click(object sender, EventArgs e)
        {
            try
            {
                if (SettingFocusSolve)
                    SettingFocusSolve = false;
                GotoFocusLocation();
                //  if (usingASCOM)
                //    {
                //while (scope.Slewing)
                //{
                //    Thread.Sleep(50); // pause for 1/20 second
                //    System.Windows.Forms.Application.DoEvents();
                //}
                ///    }
                //else
                //{

                //    while (MountMoving == true)
                //    {
                //        Thread.Sleep(50); // pause for 1/20 second
                //        System.Windows.Forms.Application.DoEvents();
                //    }
                //}
                //button33.Text = "At Focus";
                //button33.BackColor = System.Drawing.Color.Lime;
                //button35.Text = "Goto";
                //button35.UseVisualStyleBackColor = true;
            }
            catch (Exception ex)
            {
                Log("GotoFocus locaton Button Error" + ex.ToString());
            }

        }

        private bool TargetGotoOn = false;
        private void GotoTargetLocation()  //gototargetlocationhere
        {
            try
            {
                /// ***** 4-12 16
                ///  add for plate solve
                /// slew to prev solved TargetRA and TargetDEC
                /// take an image and plate solve
                /// compare to previously defined target location, if not within tolerance defined on plate solve tab repeat.  


                if (TargetLocObtained == false)
                {
                    MessageBox.Show("Target Position Not Set", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                //bool PHDpaused = false;
                //if (PHDcommand(17) == "3")//pause if goto target from button push
                //{
                //    PHDcommand(PHD_PAUSE);
                //    Log("Guiding Paused for slew");
                //    PHDpaused = true;
                //}
                if (backgroundWorker2.IsBusy == false)
                {
                    StopPHD();//already paused if doing meridian flip
                    //   if (usingASCOM)
                    //   scope.SlewToCoordinates(TargetRA, TargetDEC);
                    scope.SlewToCoordinatesAsync(TargetRA, TargetDEC);//allows abort
                }
                else
                    scope.SlewToCoordinates(TargetRA, TargetDEC);

                if (backgroundWorker2.IsBusy == false)//cant call log if running from bkgnd worker
                    Log("slewing to target");
                toolStripStatusLabel1.Text = "Slewing";
                toolStripStatusLabel1.BackColor = Color.Red;
                this.Refresh();
                while (scope.Slewing)
                {
                    Thread.Sleep(50); // pause for 1/20 second
                    System.Windows.Forms.Application.DoEvents();//this makes it necssary to push twice!!!????
                }




                // add 4-12 16

                // remd 4-14-16 not confirming w/ platesolve not working.
                //if (checkBox32.Checked)
                //{
                //    toolStripStatusLabel1.Text = "Soving";
                //    toolStripStatusLabel1.BackColor = Color.Lime;               
                //    button35.Text = "Pending"; // change the set button only...since not necessarily there (could be from prev session) 
                //    button35.BackColor = System.Drawing.Color.Yellow;
                //    this.Refresh();
                //    Log("Confirming Focus plate solve");
                //    ConfirmingLocation = true;
                //    fileSystemWatcher7.Path = GlobalVariables.Path2;
                //    fileSystemWatcher7.SynchronizingObject = this;

                //    if (fileSystemWatcher7.EnableRaisingEvents == false)
                //        fileSystemWatcher7.EnableRaisingEvents = true;

                //    ///  ****   4-12-16  todo send neb command to caputre 1 image at exp time per textbox65.  
                //    ///  cant just push cpautre button since the setting wont necessarily be correct if doing focus in middle of sequence.  


                //    SolveCapture();

                //}
                //    //SetForegroundWindow(Handles.NebhWnd);  //rem'd to testing 
                //    //PostMessage(Handles.CaptureMainhWnd, BN_CLICKED, 0, 0);//rem'd for testing



                //    //if (checkBox25.Checked == true)//repeat until tolerance met
                //    //{
                //    //  Thread.Sleep(5000);
                //    //     Log("calibrating");
                //    //scope.SlewToCoordinates(RA, DEC);
                //    CurrentRA = scope.RightAscension;
                //    CurrentDEC = scope.Declination;
                //    // first attempts at comparing parse solve coords w/ scope coords.
                //    //need to FIX...seems to maintain RA and DEC from first plate solve after the second one. 
                //    //if (usingASCOM == true)
                //    //{
                //    //    //  scope = new ASCOM.DriverAccess.Telescope(devId);
                //    //    scope.SlewToCoordinates(CurrentRA, CurrentDEC);//go back to where originally though it was supposed to be
                //    //                                                   //should be closer after the sync
                //    //}
                //    //else
                //    //    MessageBox.Show("Must use ASCOM mount connection", "scopefocus");

                //    if ((Math.Abs(CurrentRA - RA) * 60 > Convert.ToDouble(textBox59.Text)) || (Math.Abs(CurrentDEC - DEC) * 60 > Convert.ToDouble(textBox59.Text)))//*************untested!!****
                //    {
                //        scope.SyncToCoordinates(RA, DEC);//sync to parsed(solve) location 
                //        Log("repeating" + "DeltaRA = " + ((Math.Abs(CurrentRA - RA) * 60).ToString()) + "     DeltaDEC = " + ((Math.Abs(CurrentDEC - DEC) * 60).ToString()));
                //        if (solvetry < 3)
                //        {
                //            solvetry++;
                //            SolveCapture();
                //        }
                //        else
                //        {
                //            Log("Solve Confirmation failed, using mount coordinates");

                //        }
                //        //     button55.PerformClick();//prob dont need since fsw7 still on
                //        //if (fileSystemWatcher7.EnableRaisingEvents == false)
                //        //    fileSystemWatcher7.EnableRaisingEvents = true;
                //        //SetForegroundWindow(Handles.NebhWnd);  //rem'd to testing 
                //        //PostMessage(Handles.CaptureMainhWnd, BN_CLICKED, 0, 0);//rem'd for testing
                //    }

                //    else
                //    {
                //        Log("Confimred: At Target Location");
                //        solvetry = 0;
                //        ConfirmingLocation = false;
                //    }

                //}

                // ******  ??  this is different the goto focus location   4-12 16   ****8  ??? 
                // ???? move this to after solve complete...???   4-12 16




                if (backgroundWorker2.IsBusy == false)
                {
                    Log("At target");
                    toolStripStatusLabel1.Text = "Ready";
                    toolStripStatusLabel1.BackColor = Color.WhiteSmoke;
                    button35.Text = "At Target";
                    button35.BackColor = Color.Lime;
                    button33.Text = "Goto";
                    button33.UseVisualStyleBackColor = true;
                }
                // if (PHDpaused)//resume for button push.  
                if (Handles.PHDVNumber == 2)
                    resumePHD2();
                else
                    ResumePHD();

                // ****   4-12 16 begin addition to use plate solve confirmation


                //if (checkBox24.Checked)
                //{
                //    Log("Confirming Focus plate solve");

                //    fileSystemWatcher7.Path = GlobalVariables.Path2;
                //    if (fileSystemWatcher7.EnableRaisingEvents == false)
                //        fileSystemWatcher7.EnableRaisingEvents = true;

                //    ///  *****  4-12-16  todo send neb command to caputre 1 image at exp time per textbox65.  

                //    SolveToleranceCheck();
                //    //SetForegroundWindow(Handles.NebhWnd);  //rem'd to testing 
                //    //PostMessage(Handles.CaptureMainhWnd, BN_CLICKED, 0, 0);//rem'd for testing


                //    //if (checkBox25.Checked == true)//repeat until tolerance met
                //    //{
                //    //  Thread.Sleep(5000);
                //    //     Log("calibrating");
                //    //scope.SlewToCoordinates(RA, DEC);
                //    CurrentRA = scope.RightAscension;
                //    CurrentDEC = scope.Declination;
                //    // first attempts at comparing parse solve coords w/ scope coords.
                //    //need to FIX...seems to maintain RA and DEC from first plate solve after the second one. 
                //    //if (usingASCOM == true)
                //    //{
                //    //    //  scope = new ASCOM.DriverAccess.Telescope(devId);
                //    //    scope.SlewToCoordinates(CurrentRA, CurrentDEC);//go back to where originally though it was supposed to be
                //    //                                                   //should be closer after the sync
                //    //}
                //    //else
                //    //    MessageBox.Show("Must use ASCOM mount connection", "scopefocus");

                //    if ((Math.Abs(CurrentRA - RA) * 60 > Convert.ToDouble(textBox59.Text)) || (Math.Abs(CurrentDEC - DEC) * 60 > Convert.ToDouble(textBox59.Text)))//*************untested!!****
                //    {
                //        scope.SyncToCoordinates(RA, DEC);//sync to parsed(solve) location 
                //        Log("repeating" + "DeltaRA = " + ((Math.Abs(CurrentRA - RA) * 60).ToString()) + "     DeltaDEC = " + ((Math.Abs(CurrentDEC - DEC) * 60).ToString()));
                //        //     button55.PerformClick();//prob dont need since fsw7 still on
                //        if (fileSystemWatcher7.EnableRaisingEvents == false)
                //            fileSystemWatcher7.EnableRaisingEvents = true;
                //        //SetForegroundWindow(Handles.NebhWnd);  //rem'd to testing 
                //        //PostMessage(Handles.CaptureMainhWnd, BN_CLICKED, 0, 0);//rem'd for testing
                //    }


                //    // }

                //}






            }
            catch (Exception ex)//*********remd 5-5-14
            {
                //Log("GotoTargetLocation Error" + ex.ToString());
                //   Send("GotoTargetLocation Error" + ex.ToString());
                //   FileLog("GotoTargetLocation Error" + ex.ToString());

            }
        }

        //goto target
        private void button35_Click(object sender, EventArgs e)
        {
            try
            {
                if (SettingTargetSolve)
                    SettingTargetSolve = false;
                GotoTargetLocation();
                //   if (usingASCOM)
                //   {
                //while (scope.Slewing)
                //{
                //    Thread.Sleep(50); // pause for 1/20 second
                //    System.Windows.Forms.Application.DoEvents();//this makes it necssary to push twice!!!????
                //}
                //}
                //else
                //{
                //    while (MountMoving == true)
                //    {
                //        Thread.Sleep(50); // pause for 1/20 second
                //        System.Windows.Forms.Application.DoEvents();
                //    }
                ////}
                //button35.Text = "At Target";
                //button35.BackColor = System.Drawing.Color.Lime;
                //button33.Text = "Goto";
                //button33.UseVisualStyleBackColor = true;
            }
            catch (Exception ex)
            {
                Log("gotoarget location Button error" + ex.ToString());
            }
        }
        //Abort slew
        private void button37_Click(object sender, EventArgs e)
        {
            try
            {
                if (usingASCOM)
                {
                    scope.AbortSlew();
                    FocusGotoOn = false;
                    TargetGotoOn = false;
                    toolStripStatusLabel1.Text = "Slew Aborted";

                }
                else
                {

                    //if (backgroundWorker1.WorkerSupportsCancellation == true)
                    //{
                    //    // Cancel the asynchronous operation.
                    //    backgroundWorker1.CancelAsync();
                    //}

                    //if (port2.IsOpen == false)
                    //    Port2Open();
                    //port2.DiscardOutBuffer();
                    //port2.DiscardInBuffer();
                    //Thread.Sleep(20);
                    //port2.Write("M");
                    //Thread.Sleep(50);
                    //port2.DiscardOutBuffer();
                    //port2.DiscardInBuffer();

                    //FocusGotoOn = false;
                    //TargetGotoOn = false;
                    //toolStripStatusLabel1.Text = "Slew Aborted";
                    ////   port2.Close();

                }
            }
            catch (Exception ex)
            {
                Log("Abort Slew Button Error" + ex.ToString());

            }

        }

        //try send focus time via script  *****Close****
        private void SendFocusTime(float dur)
        {
            if (!UseClipBoard.Checked) // 11-8-16
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
                            NebListenStart(Handles.NebhWnd, SocketPort);
                        }
                    }
                    else
                    {

                        if (clientSocket.Connected == false)
                        {
                            Log("Socket 1 was not connected");
                            NebListenStart(Handles.NebhWnd, SocketPort);
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
                    Thread.Sleep(3000); //**** was 1000 4-22-14
                                        //*********************goes farther without this NebListenStop but doesn't hit enter 
                                        // for client socket closed message boxes
                                        //*************need to eleminiate mult. socket opwn/close  *******************

                    //   NebListenStop();
                }
                catch (Exception ex)
                {
                    Log("SendFocutTime Error" + ex.ToString());
                    Send("SendFocutTime Error" + ex.ToString());
                    FileLog("SendFocutTime Error" + ex.ToString());

                }
            }
            else // 11-8-16
            {
                NebListenStart(Handles.NebhWnd, SocketPort);
                delay(1);
                Clipboard.Clear();
                msdelay(500);
                Clipboard.SetDataObject("//NEB SetDuration " + dur.ToString(), false, 3, 500);
                //   Clipboard.SetText("//NEB SetDuration " + dur.ToString());
                msdelay(750);
                //int i = 0;
               //if (!NebCommandConfirm("SetDuration " + dur.ToString(), 0))
               //         if (!clipboardRetry)
               //         {
               //             clipboardRetry = true;
               //             SendFocusTime(dur);
               //         }
               //         else
               //             ClipboardFailure("SetDuration");
                //{
                //    i++;
                //    msdelay(100);
                //    if (i == 30)
                //    {
                //        Log("Solve Capture Failed");
                //        Send("Solve Caputre Failed");
                //        return;
                //    }
                //}

                Log("FocusTime sent " + dur.ToString());
                Clipboard.Clear();
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
            //  System.Object lockThis = new System.Object();


            try
            {
                FileLog2("filterfocus called");
                if ((IsServer()) & (checkBox8.Checked == true))
                    SlaveFocus();

                //if ((radioButton14.Checked == false) && (radioButton15.Checked == false) && (radioButton16.Checked == false))
                //{
                //    MessageBox.Show("Must select PHD pause method", "scopefocus");
                //    return;
                //}
                FilterFocusOn = true;
                fileSystemWatcher4.EnableRaisingEvents = false;//**************added2_29
             
                /*  NebListenStop();   **************remd 2_29  next 1 rem as well
                  Thread.Sleep(500);
                  SetForegroundWindow(NebhWnd);
                  SendKeys.SendWait("~");
                  SendKeys.Flush();
                  Thread.Sleep(1000);
                 */
                Handles.Editfound = 0;

                toolStripStatusLabel1.Text = "Filter Focus On";
                //   clientSocket = new TcpClient();
                FocusTime = (float)numericUpDown22.Value;
                //  if (radioButton12.Checked == true)\
                if (Handles.NebVNumber == 2)
                    FocusTime = FocusTime * 1000;
                //     textBox22.Text = FocusTime.ToString();
                this.Refresh();
                if (checkBox10.Checked == false) // Need slew to focus -- check to make sure focus/target locations have been set
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
                        //Thread.Sleep(1500);
                        delay(1);
                        NebFineFocus();
                        // Thread.Sleep(1000);
                        delay(1);
                        //  lock (lockThis)
                        //   {
                        gotoFocus();//***************** 8-26-13 ********* this needs to finish before 4279 can continue
                                    //   }
                    }

                    

                }
                else    //need to slew to focus
                {
                    // SlewDelay = (int)numericUpDown32.Value * 1000;
                    //{pause phd, goto focus....(pause) --> Frame/focus --> fine focus(autoclick center) --> goto focus-->
                    // Abort neb--> goto target --> resume PHD} --> capture series

                    //   StopPHD();
                    //Thread.Sleep(500);
                    FileLog2("slew to focus location");
                    delay(1);
                    toolStripStatusLabel1.Text = "Slewing to Focus Location";
                    this.Refresh();
                    //   while (MountMoving == true)
                    GotoFocusLocation();
                    //  *****rem for debugging******
                    // MessageBox.Show("simulate slew");
                    //while (MountMoving == true)
                    //{
                    //    Thread.Sleep(50); // pause for 1/20 second
                    //    System.Windows.Forms.Application.DoEvents();
                    //}

                    //    Thread.Sleep(SlewDelay);
                    // CheckForSlewDone();
                    //if (MountMoving == false)
                    //{

                    //if (checkBox16.Checked == true)
                    //{
                    //    if (Handles.PHDVNumber == 2)
                    //        resumePHD2();
                    //    else
                    //        ResumePHD();
                    //}
                    //else


                    button33.Text = "At Focus";
                    button33.BackColor = System.Drawing.Color.Lime;
                    button35.Text = "Goto Target";
                    button35.UseVisualStyleBackColor = true;
                    // Thread.Sleep(1000);
                    delay(1);
                    SendFocusTime(FocusTime);
                    // Thread.Sleep(1000);
                    delay(1);
                    NebFineFocus();//has 1 sec delay at end
                                   // Thread.Sleep(1000);
                    delay(1);
                    //   clientSocket = null;*************remd 2_29??not sure if needed, not tested yet
                    gotoFocus();

                    //    ResumePHD();
                    //   }




                }
            }
            catch (Exception ex)
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
        // [DllImport("User32.dll")]
        //public static extern Int32 GetWindowText(int hWnd, StringBuilder s, int nMaxCount);
        // [DllImport("User32.dll")]
        // public static extern Int32 GetClassName(int hWnd, StringBuilder s, int nMaxCount);//added to try to get edit box 
        //  [DllImport("User32.dll")]
        //   public static extern Int32 GetWindowTextLength(int hwnd);
        [DllImport("user32.dll", EntryPoint = "GetDesktopWindow")]
        public static extern int GetDesktopWindow();
        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")]
        public static extern int ChildWindowFromPointEx(int hWnd, Point pt, uint uFlags);

        public delegate int Callback(int hWnd, int lParam);


        //   public int hwndDuraton2;
        //   public int LoadScripthwnd2;
        //   public int Framehwnd2;
        //   public int Finehwnd2;
        ////   public int Advhwnd2;
        ////   public int Advancedhwnd2;
        //   public int Aborthwnd2;
        //   public int hwndDuration2;
        //   public int panelhwnd2;
        //   public int SlaveStatushwnd;
        //   public int Gotofocushwnd;//slave mode handles(2nd instance of scopefocus)
        //   public int Capturehwnd;
        //   public int Flathwnd;
        //   public int Pausehwnd;
        //   public int EnumChildGetValue(int hWnd, int lParam)
        //   {

        //       StringBuilder formDetails = new StringBuilder(256);
        //       StringBuilder formClass = new StringBuilder(256);
        //       int txtValue;
        //       int txtValue2;
        //       string editText = "";
        //       string classtext = "";
        //       txtValue = GetWindowText(hWnd, formDetails, 256);
        //       editText = formDetails.ToString().Trim();

        //       txtValue2 = GetClassName(hWnd, formClass, 256);
        //       classtext = formClass.ToString().Trim();
        //       if (lParam == 0)
        //       {
        //           if (editText == "panel")//doesn't work w/ msctls_statusbar32  either
        //           {
        //               panelhwnd = hWnd;

        //           }
        //           if (classtext == "Edit")
        //           {
        //               editfound++;
        //               if (editfound == 2)
        //               {
        //                   hwndDuration = hWnd;
        //               }
        //           }

        //           if (editText == "Advanced")
        //           {
        //               int Advancedhwnd = hWnd;
        //           }
        //           //*****************************need to test after writing auto camera find stuff ***********************************         
        //           if (editText == NebCamera + " Setup")
        //           {
        //               if (setupWindowFound == false)//picks first one
        //               {
        //                  int Advhwnd = hWnd;
        //                   //   Log("Adv" + Advhwnd.ToString());
        //                   setupWindowFound = true;
        //               }
        //           }
        //           if (editText == "Capture Series")
        //               CaptureMainhWnd = hWnd;
        //           if (editText == "Abort")
        //           {

        //               Aborthwnd = hWnd;
        //               //   Log("Abort " + Aborthwnd.ToString());
        //           }
        //           if (editText == "Frame and Focus")
        //           {

        //               Framehwnd = hWnd;
        //               // Log("Frame " + Framehwnd.ToString());
        //           }
        //           if (editText == "Fine Focus")
        //           {

        //               Finehwnd = hWnd;
        //               //  Log("Fine " + Finehwnd.ToString());
        //           }
        //           if (editText == "Load script")//not finding it
        //           {
        //               LoadScripthwnd = hWnd;
        //               //   Log("LoadScript " + LoadScripthwnd.ToString());
        //           }
        //       }
        //       if (lParam == 2)//added to find handles of second instence of Neb
        //       {
        //           if (editText == "panel")//doesn't work w/ msctls_statusbar32  either
        //           {
        //               panelhwnd2 = hWnd;

        //           }
        //           if (classtext == "Edit")
        //           {
        //               editfound++;
        //               if (editfound == 2)
        //               {
        //                   hwndDuration2 = hWnd;
        //               }
        //           }

        //           if (editText == "Advanced")
        //           {
        //               int Advancedhwnd2 = hWnd;
        //           //    Advancedhwnd2 = hWnd;
        //           }
        //           //*****************************need to test after writing auto camera find stuff ***********************************         
        //           if (editText == NebCamera + " Setup")
        //           {
        //               if (setupWindowFound == false)//picks first one
        //               {
        //                  int Advhwnd2 = hWnd;
        //                   //   Log("Adv" + Advhwnd.ToString());
        //                   setupWindowFound = true;
        //               }
        //           }

        //           if (editText == "Abort")
        //           {

        //               Aborthwnd2 = hWnd;
        //               //   Log("Abort " + Aborthwnd.ToString());
        //           }
        //           if (editText == "Frame and Focus")
        //           {

        //               Framehwnd2 = hWnd;
        //               // Log("Frame " + Framehwnd.ToString());
        //           }
        //           if (editText == "Fine Focus")
        //           {

        //               Finehwnd2 = hWnd;
        //               //  Log("Fine " + Finehwnd.ToString());
        //           }
        //           if (editText == "Load script")//not finding it
        //           {
        //               LoadScripthwnd2 = hWnd;
        //               //   Log("LoadScript " + LoadScripthwnd.ToString());
        //           }
        //       }

        //       if (checkBox20.Checked == true)
        //       {
        //           if (editText == "GotoFocus")
        //           {
        //               Gotofocushwnd = hWnd;
        //               Log("Slave Gotofocus handle found  " + Gotofocushwnd.ToString());
        //           }
        //           if (editText == "Capture")
        //           {
        //               Capturehwnd = hWnd;
        //               Log("Slave Capture Handle Found  " + Capturehwnd.ToString());

        //           }
        //           if (editText == "Flat")
        //           {
        //               Flathwnd = hWnd;
        //               Log("Slave Flat handle found  " + Flathwnd.ToString());
        //           }
        //           if (editText == "Pause")
        //           {
        //               Pausehwnd = hWnd;
        //               Log("SlavePause Handle found  " + Pausehwnd.ToString());
        //           }
        //           if (editText == "Not Connected")
        //           {
        //               SlaveStatushwnd = hWnd;
        //               Log("Status Handle " + SlaveStatushwnd.ToString());
        //           }

        //       }





        //       //MessageBox.Show("Contains text of control "+ editText);
        //       return 1;

        //   }

        private const int SW_SHOW = 5;
        private const int SW_SHOWMAXIMIZED = 3;
        private const int SW_RESTORE = 9;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(int hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(int hWnd, int nCmdShow);
        //neblistenstarthere
        private int retry = 0;
        private void NebListenStart(int hwnd, int port)//was blank
        {
            FileLog2("NebListenStart");
            if (retry < 3)
            {
                try
                {

                    int hwndChild;
                    int hwndChild1;
                    int hwndChild2;
                    if (IsSlave())
                        Thread.Sleep(4000);//*****6-17 was 3000
                    ShowWindowAsync(hwnd, SW_SHOW);
                    Thread.Sleep(200);
                    ShowWindowAsync(hwnd, SW_RESTORE);
                    Thread.Sleep(500);
                    //    ShowWindow(hwnd, SW_SHOW);
                    Thread.Sleep(500);
                    SetForegroundWindow(hwnd);//was NebhWnnd;
                    Thread.Sleep(500);
                    SendKeys.SendWait("^r");
                    SendKeys.Flush();
                    Thread.Sleep(1000);
                    Handles.LoadScripthwnd = FindWindow(null, "Load script");
                    //    GetHandles();
                    if (Handles.LoadScripthwnd == 0)
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
                        Handles.LoadScripthwnd = FindWindow(null, "Load script");
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
                    SetForegroundWindow(Handles.LoadScripthwnd);
                    //   int indexhandle = FindWindowByIndex(LoadScripthwnd, 1, "ComboBoxEx32");
                    //   Log("Indexhandle " + indexhandle.ToString());
                    //   EnumChildGetValue(LoadScripthwnd, 0);
                    hwndChild = FindWindowEx(Handles.LoadScripthwnd, 0, "ComboBoxEx32", null);
                    //  Log("cbex32" + hwndChild.ToString());
                    hwndChild1 = FindWindowEx(hwndChild, 0, "ComboBox", null);
                    //   Log("cb" + hwndChild1.ToString());
                    hwndChild2 = FindWindowEx(hwndChild1, 0, "edit", null);
                    //  Log("enterscript name box" + hwndChild2.ToString());
                    if (hwndChild == 0 || hwndChild1 == 0 || hwndChild2 == 0)
                        return;
                    if (UseClipBoard.Checked)
                    {
                        Clipboard.Clear();
                        Thread.Sleep(250); // added 12-29-16
                        ScriptName = "\\Listen.neb";
                    }
                    if ((checkBox14.Checked == true) && (!UseClipBoard.Checked))
                    {
                        port = 4302;
                        ScriptName = "\\listenPort2.neb";
                    }

                    // else 11-8-16
                    if ((checkBox14.Checked == false) && (!UseClipBoard.Checked))
                    {
                        port = 4301;
                        ScriptName = "\\listenPort.neb";
                    }

                    //need to get combobox then edit.....
                    //  string send = "";
                    // string send = NebPath + @"\listenPort.neb";
                    //   send = "C:\\Program Files (x86)\\Nebulosity3\\listenPort.neb";
                    //   string ScriptName = "\\listenPort.neb";
                    string send = GlobalVariables.NebPath + ScriptName;

                    StringBuilder sb = new StringBuilder(send);
                    //  sb = send;
                    SendMessage(hwndChild2, WM_SETTEXT, 0, sb);
                    Thread.Sleep(1000);//******5-29 was 250
                                       //  SendKeys.SendWait("~"); // 10-25-16 changed to  send not sendwait below
                    SendKeys.Send("~");
                    Thread.Sleep(250); // added 10-25-16
                    SendKeys.Flush();
                    Thread.Sleep(1000);


                    // make sure it script filename got sent
                    // ???



                    if (Handles.LoadScripthwnd != 0)
                    {
                        if ((IsServer()) & (SlaveFlatOn == false))
                        {

                            Log("Waiting for Slave connection");
                            toolStripStatusLabel1.Text = "Waiting for slave connection";
                            // this.Refresh();
                            while (working == true)
                            {

                                WaitForSequenceDone("Waiting for connection", GlobalVariables.NebSlavehwnd);
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
                            WaitForSequenceDone("Waiting for connection", Handles.NebhWnd);
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
                    if (!UseClipBoard.Checked)
                    {
                        if ((clientSocket.Connected == false) & (port == 4301))
                            Connect1(port);
                        if ((clientSocket2.Connected == false) & (port == 4302))
                            Connect2(port);
                    }
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
                    //  NebListenOn = true;

                    // added 11-8-16
                    if (!UseClipBoard.Checked)
                    {
                        if (CheckConnection() == "Waiting for connection")
                        {
                            retry = 0;
                            return;
                        }
                        else

                        {
                            Log("retry NebListenStart");
                            retry++;
                            NebListenStart(Handles.NebhWnd, SocketPort);
                        }
                        // if the load script windows still up try sending cancel and repeat.  
                        if (Handles.LoadScripthwnd != 0)
                        {
                            SendMessage(hwndChild2, WM_SETTEXT, 0, sb);
                            Thread.Sleep(1000);//******5-29 was 250
                                               //  SendKeys.SendWait("~"); // 10-25-16 changed to  send not sendwait below
                            SendKeys.Send("{ESC}");
                            Thread.Sleep(250); // added 10-25-16
                            SendKeys.Flush();
                            Thread.Sleep(1000);
                            retry++;
                            NebListenStart(Handles.NebhWnd, SocketPort);
                            Log("retry NebListenStart");
                        }

                    }
                    //if (UseClipBoard.Checked) // 11-12-16
                    //{
                    //    int i = 0;
                    //    while (!NebCommandConfirm("Listen 1", 0))
                    //    {
                    //        i++;
                    //        msdelay(100);
                    //        if (i == 30)
                    //        {
                    //            Log("NebListen failed");
                    //            Send("NebListen Failed");
                    //            return;
                    //        }
                    //    }
                    //}

                }
                catch (Exception ex)
                {
                    Log("NebListenStart Error" + ex.ToString());
                    Send("NebListenStart Error" + ex.ToString());
                    FileLog("NebListenStart Error" + ex.ToString());

                }
            }
            //else
            //    Log("NebListenStart Failed");
            //Send("NebListenStart Failed");
            //FileLog("NebListenStart failed");

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
                //    Log(port.ToString() + " Conneceted");
                //  }
            }
            catch (Exception e)
            {
                Log("Connect2 Error" + e.ToString());
                Send("Connect2 Error" + e.ToString());
                FileLog("Connect2 Error" + e.ToString());

            }
        }
        // added 11-8-16
        private string CheckConnection()
        {
            int StatusstripHandle = FindWindowEx(Handles.NebhWnd, 0, "msctls_statusbar32", null);

            //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
            IntPtr statusHandle = new IntPtr(StatusstripHandle);
            StatusHelper sh = new StatusHelper(statusHandle);
            string[] captions = sh.Captions;
            if (captions[0] != "")
            {
                return captions[0];
            }
            else
                return "0";

        }



        private void Connect1(int port)
        {
            //bool tryagain = true;
            //while (tryagain)
            //{
            while (CheckConnection() == "Waiting for connection")
                try
                {
                    //  if (clientSocket.Connected == false)//**************try adding 2_29
                    //  { 
                    clientSocket = new TcpClient();
                    LingerOption lingerOption = new LingerOption(true, 1); // 10-25-16 was 1
                    clientSocket.LingerState = lingerOption;
                    clientSocket.Connect("127.0.0.1", port);//*************try adding and red above)
                    Thread.Sleep(1000);
                    //   tryagain = false;
                    //   Log(port.ToString() + " Conneceted");
                    //   }
                }
                catch (SocketException e) // added 10-25-16
                {
                    Log("A connection error occured, rescan for Neb and PHD handles");
                    FileLog("A connection error occured, rescan for Neb and PHD handles");
                    Send("A connection error occured, rescan for Neb and PHD handles");

                    // remd 11-8-16
                    //MessageBox.Show("Unable to communicate with Nebulosity, Open Neb and click Rescan on Setup tab", "scopefocus");
                    //        return;

                    //DialogResult result1;
                    //result1 = MessageBox.Show("A connection error occured, click 'retry' to rescan or 'abort' to quit", "scopefocus",
                    //         MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                    //if (result1 == DialogResult.Retry)
                    //{
                    //    button51.PerformClick();
                    //    delay(1);
                    //}
                    //if (result1 == DialogResult.Abort)
                    //{
                    //    DialogResult result2;
                    //            result2 = MessageBox.Show("Dou you really wnat to quit?", "scopefocus",
                    //                 MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    //            if (result2 == DialogResult.Yes)
                    //                System.Environment.Exit(0);
                    //}
                    //Handles H = new Handles();
                    //Callback myCallBack = new Callback(H.EnumChildGetValue);

                    //H.FindHandles();
                    //if (Handles.PHDhwnd == 0)//this can really just serve as reminded that PHD is needed for some functions...not necessary for focusing.  
                    //{
                    //    DialogResult result1;
                    //    result1 = MessageBox.Show("PHD Not Found - Open and 'Retry', 'Ignore' or 'Abort' to Close", "scopefocus",
                    //         MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);

                    //    if (result1 == DialogResult.Ignore)
                    //        this.Refresh();
                    //    if (result1 == DialogResult.Retry)
                    //    {
                    //        H.FindHandles();
                    //        this.Refresh();
                    //    }
                    //    if (result1 == DialogResult.Abort)
                    //    {
                    //        DialogResult result2;
                    //        result2 = MessageBox.Show("Dou you really wnat to quit?", "scopefocus",
                    //             MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    //        if (result2 == DialogResult.Yes)
                    //            System.Environment.Exit(0);

                    //    }

                    //}
                    //else
                    //    Log("PHD version " + Handles.PHDVNumber.ToString());


                    //if (Handles.NebhWnd == 0)
                    //{
                    //    DialogResult result;
                    //    result = MessageBox.Show("Nebulosity Not Found - Open or Close & Restart and hit 'Retry' or 'Ignore' to continue",
                    //       "scopefocus", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);//change so ok moves focus
                    //    if (result == DialogResult.Ignore)
                    //        this.Refresh();
                    //    if (result == DialogResult.Retry)
                    //    {
                    //        H.FindHandles();
                    //        this.Refresh();
                    //    }
                    //    if (result == DialogResult.Abort)
                    //    {
                    //        DialogResult result2;
                    //        result2 = MessageBox.Show("Dou you really wnat to quit?", "scopefocus",
                    //             MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    //        if (result2 == DialogResult.Yes)
                    //            System.Environment.Exit(0);
                    //    }

                    //    //Log("Connect Error Line 10172" + e.ToString());
                    //    //Send("Connect Error Line 10172" + e.ToString());
                    //    //FileLog("Connect Error Line 10172" + e.ToString());

                    //}

                }
                catch (Exception e)
                {
                    Log("Connect Error Line 10172" + e.ToString());
                    Send("Connect Error Line 10172" + e.ToString());
                    FileLog("Connect Error Line 10172" + e.ToString());
                    return;
                    //    tryagain = false;

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
        private void NebFineFocus()//abort - frame - abort - fine - click star at position    nebfinefocushere
        {
            try
            {
                FileLog2("NebFine Focus on");
                toolStripStatusLabel1.Text = "Neb Fine Focus On";
                this.Refresh();
                SetForegroundWindow(Handles.NebhWnd);
                //   Thread.Sleep(3000);
                delay(3);
                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);//*******added 3_13
                                                                 //  Thread.Sleep(3000);
                delay(3);
                PostMessage(Handles.Framehwnd, BN_CLICKED, 0, 0);
                // Thread.Sleep(3000);
                delay(3);
                //  Thread.Sleep((int)FocusTime * 2000);//need to make sure at least one frame is done
                delay((int)FocusTime * 2);
                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                // Thread.Sleep(3000);
                delay(3);
                PostMessage(Handles.Finehwnd, BN_CLICKED, 0, 0);
                // Thread.Sleep(5000); 
                delay(5);

                Point xxx = new Point();
                xxx.X = Convert.ToInt32(textBox29.Text.ToString());//default is center
                xxx.Y = Convert.ToInt32(textBox30.Text.ToString());
                int starpos = ((xxx.Y << 16) | (xxx.X & 0xffff));
                int panelhwnd1 = FindWindowByIndexName(Handles.NebhWnd, 1, "panel");
                int panelhwnd3 = FindWindowByIndexName(Handles.NebhWnd, 3, "panel");//index(2) may change w. neb versions???
                int panelhwnd2 = FindWindowByIndexName(Handles.NebhWnd, 2, "panel");
                //******************* added 3_13
                int panel2Vis = IsWindowVisible(panelhwnd2);
                if (Handles.NebVNumber == 4)
                    Handles.Panelhwnd = panelhwnd1;
                else
                {
                    if (panel2Vis == 1)
                    {
                        Handles.Panelhwnd = panelhwnd2;
                        //  Log("panel2vis " + panel2Vis.ToString() + " # " + panelhwnd.ToString());
                    }
                    else
                    {
                        Handles.Panelhwnd = panelhwnd3;
                    }
                }
                int panel3Vis = IsWindowVisible(panelhwnd3);//WORKS  visable =1
                // Log("panel3vis " + panel3Vis.ToString() + " # " + panelhwnd.ToString());
                Log("Focus star click pos = " + starpos.ToString() + "  X = " + xxx.X.ToString() + "  Y = " + xxx.Y.ToString());



                //*****************end addition
                SetForegroundWindow(Handles.Panelhwnd);
                //  Log("panel2 for fine Focus" + panelhwnd.ToString());//index 3 now for some reason(3/9/12)
                // Thread.Sleep(500);
                delay(1);
                PostMessage(Handles.Panelhwnd, WM_LBUTTONDOWN, 0, starpos);//was SendMessage
                PostMessage(Handles.Panelhwnd, WM_LBUTTONUP, 0, starpos);//was SendMessage
                                                                         //  SendMessage(panelhwnd, WM_LBUTTONDOWN, 0, starpos);//was SendMessage
                                                                         //   SendMessage(panelhwnd, WM_LBUTTONUP, 0, starpos);//was SendMessage
                                                                         //  Thread.Sleep(1000);
                delay(1);
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



            catch (Exception ex)
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
                if (!UseClipBoard.Checked)
                {
                    serverStream = clientSocket.GetStream();
                    byte[] outStream3 = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                    serverStream.Write(outStream3, 0, outStream3.Length);
                    serverStream.Flush();
                    outStream3 = null;
                }
                else
                {
                    Clipboard.Clear();
                    // delay(1);
                    // SetForegroundWindow(Handles.NebhWnd);
                    // delay(1);
                    //// Thread.Sleep(1000);
                    // PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                    //// Thread.Sleep(1000);

                    msdelay(750);
                    //Clipboard.SetText("//NEB Listen 0");
                    //msdelay(750);
                }
            }
            catch (Exception ex)
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


        //[DllImport("user32.dll", SetLastError = true)]
        //internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(int hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

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

        //private int GetNebWindowSize ()
        //{
        //    RECT NebRect = new RECT();
        //    GetWindowRect(Handles.NebhWnd, out NebRect);
        //    int width = NebRect.right - NebRect.left;
        //    return width;

        //}

        private void ResizeNebWindow()
        {

            RECT NebRect = new RECT();
            if (GetWindowRect(Handles.NebhWnd, out NebRect) < 820)
            {
                MoveWindow(Handles.NebhWnd, NebRect.left, NebRect.top, 820, NebRect.bottom - NebRect.top, true);
                Log("Nebulosity window resized");
                FileLog2("Neb Window resized from width of " + (NebRect.right - NebRect.left).ToString());
            }
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
            catch (Exception ex)
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
        public static int PHD_STOP = 18;
        private void StopPHD()
        {

            // PHDcommand(PHD_PAUSE);

            //12-7-16
            ph.StopCapture();

            Log("guiding paused");
            //            try
            //            {
            //                if (radioButton14.Checked == true || radioButton16.Checked == true)
            //                {

            //                    PHDcommand(PHD_PAUSE);
            //                    /*
            //                    phdsocket = new TcpClient();
            //                    phdsocket.Connect("127.0.0.1", 4300);
            //                    byte[] buf = new byte[1];
            //                    buf[0] = (byte)((char)PHD_STOP);
            //                    phdsocket.Client.Send(buf);
            //                    phdsocket.Client.Receive(buf);
            //                    */
            ///*

            //                    SetForegroundWindow(Handles.PHDhwnd);
            //                    int phdchildStop = FindWindowByIndex(Handles.PHDhwnd, 5, "button");
            //                    Log("PHD stop button " + phdchildStop.ToString());
            //                    SetForegroundWindow(phdchildStop);
            //                    Thread.Sleep(500);
            //                    PostMessage(phdchildStop, BN_CLICKED, 0, 0);
            // */
            //                }
            //                if (radioButton15.Checked == true)
            //                {
            //                    PHDSocketPause(true);

            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                Log("StopPHD Error" + ex.ToString());
            //                Send("StopPHD Error" + ex.ToString());
            //                FileLog("StopPHD Error" + ex.ToString());
            //            }

        }
        public void PHDread()
        {
            //future read msg from phd to ensure guiding.  
        }



        // TcpClient phdsocket = new TcpClient();
        private static int PHD_GETSTATUS = 17;

        public string PHDcommand(int command)//phdcommandhere
        {






            // all below remd 12-7-16 with new PHC2Comm class

           // FileLog2("PHD Command: " + command.ToString());
            string responseData = String.Empty;
            bool success = false;
            int retry = 0;
            while (!success && retry < 4)
            {
                try
                {

                    //   while (SocketExtensions.IsConnected(phdsocket.Client))//****** 4-24-14 this might not be needed since catch and retry
                    //   {
                    //  Log("command " + command.ToString() + phdsocket.Connected.ToString());
                    //    Log("IsConnected = " + (SocketExtensions.IsConnected(phdsocket.Client).ToString()));

                    // if (SocketExtensions.IsConnected(phdsocket.Client)) 
                    if (phdsocket.Connected == false)
                    {
                        phdsocket = new TcpClient("127.0.0.1", 4300);
                        // phdsocket.Connect("127.0.0.1", 4300);
                    }
                    NetworkStream stream = phdsocket.GetStream();
                    stream.Flush();
                    byte[] buf = new byte[1];
                    byte[] data = new Byte[1];
                    buf[0] = (byte)((char)command);
                    // phdsocket.Client.Send(buf);
                    //    Log("PHDcommand send = " + command.ToString());
                    stream.WriteTimeout = 2000;
                    stream.Write(buf, 0, buf.Length);
                    //    Log("buf " + buf[0].ToString());
                    //    Thread.Sleep(200); 
                    //  data = new Byte[1];
                    // string responseData = String.Empty;//*******moved up***
                    stream.ReadTimeout = 2000;
                    Int32 bytes = stream.Read(data, 0, data.Length);
                    // responseData = System.Text.Encoding.ASCII.GetString(data, 0, bytes);
                    int response = Convert.ToInt32(data[0]);
                    responseData = response.ToString();
                    //  string responseData = String.Empty;
                    //   phdsocket.Client.Receive(data);
                    //     responseData = System.Text.Encoding.ASCII.GetString(data, 0, bytes);
                    //   Thread.Sleep(200);
                    //  phdsocket.Client.Receive(buf2);
                    //    Log("buf2 " + data[0].ToString());
                    //      Log("PHDCommand responce = " + responseData);
                    stream.Flush();

                    //return Convert.ToInt32(responseData);
                    //   return Convert.ToInt32(data[0]);
                    //     stream.Close();
                    return responseData;

                    //   }
                    //   Log("PHDsocket connection lost");
                    // return "0";
                }

                catch (IOException e)
                {
                    //   Log("IO exception" + e); //***** remd this and 8978 5-29-14
                    retry++;
                    //  Log("repeat command " + retry.ToString());
                    delay(2);
                    // return "0";

                }
                catch (SocketException e)
                {
                    //  Log("socket exception: {0}" + e);
                    retry++;
                    //    Log("reapeat command " + retry.ToString());
                    // return "0";
                }
            }
            Log("PHDCommand failed");
            return "99";
            //  return responseData;
        }
        public static int PHD_LOOP = 19;
        public static int PHD_START = 20;
        // private void PHDAutoStar()
        // {

        //     //needs cleaning up...the status are returning 2 for some reason. 
        //     try
        //     {
        //         //seems like once autostar....always autostar???"
        //         Log("Loop command " + PHDcommand(PHD_LOOP).ToString());
        //         Thread.Sleep(200);
        //      //   while (PHDcommand(PHD_GETSTATUS) != 1 || PHDcommand(PHD_GETSTATUS) != 101)
        //      //       Thread.Sleep(100);
        //     //    Log("Looping");

        //        Thread.Sleep(2000);
        //         if (PHDcommand(PHD_GETSTATUS) == "101")
        //         {
        //             Log("autostar command " + PHDcommand(PHD_AUTOSTAR).ToString());
        //             Thread.Sleep(200);
        //          //   Log("PHD Autostar");
        //         }

        //  //     while(PHDcommand(PHD_GETSTATUS) !=1)
        //        //     Thread.Sleep(100);

        //         Thread.Sleep(2000);
        //         Log("start command " + PHDcommand(PHD_START).ToString());
        //         Thread.Sleep(200);
        ////       while (PHDcommand(PHD_GETSTATUS) != 3);
        //         Log("status " + PHDcommand(PHD_GETSTATUS).ToString());
        //             Thread.Sleep(100);
        //      //   Log("PHD Started");



        //     /*
        //         phdsocket = new TcpClient();
        //         phdsocket.Connect("127.0.0.1", 4300);
        //         byte[] buf = new byte[1];
        //         buf[0] = (byte)((char)PHD_AUTOSTAR);
        //         phdsocket.Client.Send(buf);
        //         phdsocket.Client.Receive(buf);

        //         phdsocket = new TcpClient();
        //         phdsocket.Connect("127.0.0.1", 4300);
        //         byte[] buf2 = new byte[1];
        //         buf2[0] = (byte)((char)PHD_START);
        //         phdsocket.Client.Send(buf2);
        //         phdsocket.Client.Receive(buf2);
        //         */
        //         /*
        //         SetForegroundWindow(Handles.PHDhwnd);
        //         SendKeys.SendWait("%s");
        //         SendKeys.Flush();

        //         int phdcapture = FindWindowByIndex(Handles.PHDhwnd, 3, "button");
        //         Log("PHD capture button " + phdcapture.ToString());
        //         SetForegroundWindow(phdcapture);

        //         PostMessage(phdcapture, BN_CLICKED, 0, 0);
        //         Thread.Sleep(2000);

        //         int phdguide = FindWindowByIndex(Handles.PHDhwnd, 4, "button");
        //         Log("PHD guide button " + phdguide.ToString());
        //         SetForegroundWindow(phdguide);
        //         Thread.Sleep(500);
        //         PostMessage(phdguide, BN_CLICKED, 0, 0);
        //          */
        //     }
        //     catch (Exception ex)
        //     {
        //         Log("PHDAutostar Error" + ex.ToString());
        //         Send("PHDAutostar Error" + ex.ToString());
        //         FileLog("PHDAutostar Error" + ex.ToString());
        //     }
        // }
        /*
                private void PHDCoordStar()
                {
                    try
                    {
                        //This works w/ auto star select.   below is attempt to select star based on coords..works using screen coords, need
                        // to change to active window coords.

                        SetForegroundWindow(Handles.PHDhwnd);
                        int phdcapture = FindWindowByIndex(Handles.PHDhwnd, 3, "button");
                        Log("PHD capture button " + phdcapture.ToString());
                        SetForegroundWindow(phdcapture);

                        PostMessage(phdcapture, BN_CLICKED, 0, 0);
                        Thread.Sleep(1000);//time to place cursor
                        Point pt = new Point();
                        GetCursorPos(ref pt);

                        RECT rc;
                        //  Rectangle rect = new Rectangle();


                        //***************WORKS!*************
                        int phdpanel = FindWindowEx(Handles.PHDhwnd, 0, null, "panel");
                        Log("Phd panel " + phdpanel.ToString());
                        SetForegroundWindow(phdpanel);
                        GetWindowRect(phdpanel, out rc);//find upper left corner to reference location on screen.  
                        //allows for converion of position to "acitve window"
                        //      textBox35.Text = rc.left.ToString();
                        //      textBox36.Text = rc.top.ToString();
                        Point xxx = new Point();
                        //  xxx.X = Convert.ToInt32(textBox33.Text.ToString()) - Convert.ToInt32(textBox35.Text.ToString());
                        //   xxx.Y = Convert.ToInt32(textBox34.Text.ToString()) - Convert.ToInt32(textBox36.Text.ToString());
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



                        int phdguide = FindWindowByIndex(Handles.PHDhwnd, 4, "button");
                        Log("PHD guide button " + phdguide.ToString());
                        SetForegroundWindow(phdguide);
                        Thread.Sleep(500);
                        PostMessage(phdguide, BN_CLICKED, 0, 0);

                    }
                    catch (Exception ex)
                    {
                        Log("PHDCoordStar Error" + ex.ToString());
                        Send("PHDCoordStar Error" + ex.ToString());
                        FileLog("PHDCoordStar Error" + ex.ToString());
                    }

                }
         */
        //radiobutton 14 was use autostar
        private void ResumePHD()
        {
            try
            {
                //if (radioButton14.Checked == true)//auto star= hit stop, slew to focus, slew back then hit guide
                //{
                //    PHDAutoStar();
                //}

                //if (radioButton16.Checked == true)//use coordinates to select star , stop, slew, slew back, capture, select star, guide
                //{
                //    PHDCoordStar();
                //    //*************************untested***************** 4-9-12
                //    Thread.Sleep(4000);//wait then check to make sure ther is no LOW SNR msg (see  below)
                //    int StatusstripHandle = FindWindowEx(Handles.PHDhwnd, 0, "msctls_statusbar32", null);

                //    //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                //    IntPtr statusHandle = new IntPtr(StatusstripHandle);
                //    StatusHelper sh = new StatusHelper(statusHandle);
                //    string[] captions = sh.Captions;
                //    //******************not sure of exact messages....may just be albe to look for "guiding" OR
                //    //may say "guiding" even if no star, thus need to look for LOW SNR"
                //    if (captions[0] == "Guiding")//******must rem for sim debugging *****
                //        if (captions[1] != null)
                //            if (captions[1].Substring(0, 3) == "LOW")
                //            {
                //                radioButton15.Checked = false;
                //                radioButton14.Checked = true;
                //                if (Handles.PHDVNumber == 2)
                //                    resumePHD2();
                //                else
                //                    ResumePHD();
                //            }
                //}


                //   if (radioButton15.Checked == true)
                //  {

               ph.Guide();
// remd 12-7-16  replaced with above. 

                //if (Handles.PHDVNumber == 2)
                //    resumePHD2();
                //else
                //{
                //    PHDcommand(PHD_RESUME);



                    //*******************untested************** 4-9-12 (see comments above)
                    //Thread.Sleep(4000);//wait 4 seconds then make sure there is a guide star
                    ////if no guidestar, repeat using autostar select.  
                    ////will catch if no star after that
                    //int StatusstripHandle = FindWindowEx(Handles.PHDhwnd, 0, "msctls_statusbar32", null);

                    ////    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                    //IntPtr statusHandle = new IntPtr(StatusstripHandle);
                    //StatusHelper sh = new StatusHelper(statusHandle);

                    //string[] captions = sh.Captions;
                    //if (captions[0] == "Guiding")//******must rem for sim debugging *****
                    //    if (captions[1] != null)
                    //        if (captions[1].Substring(0, 3) == "LOW")
                    //        {
                    //            radioButton15.Checked = false;
                    //            radioButton14.Checked = true;
                    //          //  if (Handles.PHDVNumber == 2)
                    //          //      resumePHD2();
                    //         //   else
                    //                ResumePHD();
                    //        }
                    //********also there may be new addition to send autostar by socket command*****************    
                    //  ResumePHD();
              //  }
                // }

            }
            catch (Exception ex)
            {
                Log("ResumePHD Error" + ex.ToString());
                Send("ResumePHD Error" + ex.ToString());
                FileLog("ResumePHD Error" + ex.ToString());
            }
        }

        public static int PHD_PAUSE = 1;
        public static int PHD_RESUME = 2;
        public static int PHD_AUTOSTAR = 14;


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

            catch (Exception ex)
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

        //public static IntPtr SearchForWindow(string wndclass, string title)
        //{
        //    SearchData sd = new SearchData { Wndclass = wndclass, Title = title };
        //    EnumWindows(new EnumWindowsProc(EnumProc), ref sd);
        //    return sd.hWnd;
        //}

        //public static bool EnumProc(IntPtr hWnd, ref SearchData data)
        //{
        //    // Check classname and title 
        //    // This is different from FindWindow() in that the code below allows partial matches
        //    StringBuilder sb = new StringBuilder(1024);
        //    GetClassName(hWnd, sb, sb.Capacity);
        //    if (sb.ToString().StartsWith(data.Wndclass))
        //    {
        //        sb = new StringBuilder(1024);
        //        GetWindowText(hWnd, sb, sb.Capacity);
        //        if (sb.ToString().StartsWith(data.Title))
        //        {
        //            data.hWnd = hWnd;
        //            return false;    // Found the wnd, halt enumeration
        //        }
        //    }
        //    return true;
        //}

        //public class SearchData
        //{
        //    // You can put any vars in here...
        //    public string Wndclass;
        //    public string Title = "Nebulosity";
        //    public IntPtr hWnd;
        //}

        //      private delegate bool EnumWindowsProc(IntPtr hWnd, ref SearchData data);

        //      [DllImport("user32.dll")]
        //      [return: MarshalAs(UnmanagedType.Bool)]
        //      private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, ref SearchData data);

        //      [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        //      public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        // //     [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        ////      public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);



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
        //Radiobutton 15 was resume guiding on same star
        //private void radioButton15_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (radioButton15.Checked == true)
        //        MessageBox.Show("Ensure PHD Server is enabled", "scopefocus");
        //}

        private void numericUpDown22_ValueChanged(object sender, EventArgs e)
        {
            FocusTime = (float)numericUpDown22.Value;

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
        /*
        private void button31_Click(object sender, EventArgs e)
        {
            if (paused == true)
            {
            //    button12.PerformClick();
                return;
            }
            SequenceGo();
            button31.BackColor = System.Drawing.Color.Lime;
        }
        */
        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            CountFilterTotal();
            if (checkBox9.Checked == false)
            {
                numericUpDown29.Value = 0;
                numericUpDown30.Value = 0;
            }
            //if (checkBox5.Checked == false)
            //{
            //    MessageBox.Show("Use Dark 5 for single dark frame", "scopefocus");
            //    checkBox9.Checked = false;
            //}
        }


        private void textBox22_TextChanged(object sender, EventArgs e)
        {
            //  camera = textBox22.Text.ToString();
        }

        // public void TextCSV_FromDataTable(DataTable dt)






        //export
        public void TextCSV_FromDataTable(DataTable dt, string filename)
        {
            // Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            // Create an Excel object and add workbook...
            //orig   ApplicationClass excel = new ApplicationClass();

            //   Workbook workbook = excel.Application.Workbooks.Add(true); // true for object template???
            //   string path = textBox11.Text.ToString();  //worked

            //   string fullpath = path + @"\scopefocusData_" + DateTime.Now.ToString("yyy-M-dd_HHmmss") + ".txt";  //works
            //   string fullpath = filename +"_" + DateTime.Now.ToString("yyy-M-dd_HHmmss") + ".txt";..this writes filename.txt_datstuff.txt...need to fix
            //    string fullpath = filename + "_" + DateTime.Now.ToString("yyy-M-dd_HHmmss") + ".txt";
            StreamWriter dataBU;
            dataBU = new StreamWriter(filename);


            //using (SqlCeConnection con = new SqlCeConnection(conString))
            //{
            //    con.Open();
            //    using (SqlCeDataAdapter a = new SqlCeDataAdapter("SELECT * FROM table1", con))
            //    {


            //DataTable dt = new DataTable();
            //a.Fill(t);
            //dataGridView1.DataSource = t;
            //a.Update(t);

            foreach (DataRow row in dt.Rows)
            {
                bool firstCol = true;
                foreach (DataColumn col in dt.Columns)
                {
                    if (!firstCol) dataBU.Write(",");
                    dataBU.Write(row[col].ToString());
                    firstCol = false;
                }
                dataBU.WriteLine();
            }
            dataBU.Close();
            MessageBox.Show("Export to textCSV file complete", "scopefocus");

        }
        //   con.Close();
        //  }

        //DataTable t = new DataTable();
        //   a.Fill(t);
        //     dataGridView1.DataSource = t;
        //   a.Update(t);




        //  }


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





        ////************ 7-25-14 try this......


        //public static string SQLGetType(object type, int columnSize, int numericPrecision, int numericScale)
        //{
        //    switch (type.ToString())
        //    {
        //        case "System.Byte[]":
        //            return "VARBINARY(MAX)";

        //        case "System.Boolean":
        //            return "BIT";

        //        case "System.DateTime":
        //            return "DATETIME";

        //        case "System.DateTimeOffset":
        //            return "DATETIMEOFFSET";

        //        case "System.Decimal":
        //            if (numericPrecision != -1 && numericScale != -1)
        //                return "DECIMAL(" + numericPrecision + "," + numericScale + ")";
        //            else
        //                return "DECIMAL";

        //        case "System.Double":
        //            return "FLOAT";

        //        case "System.Single":
        //            return "REAL";

        //        case "System.Int64":
        //            return "BIGINT";

        //        case "System.Int32":
        //            return "INT";

        //        case "System.Int16":
        //            return "SMALLINT";

        //        case "System.String":
        //            return "NVARCHAR(" + ((columnSize == -1 || columnSize > 8000) ? "MAX" : columnSize.ToString()) + ")";

        //        case "System.Byte":
        //            return "TINYINT";

        //        case "System.Guid":
        //            return "UNIQUEIDENTIFIER";

        //        default:
        //            throw new Exception(type.ToString() + " not implemented.");
        //    }
        //}

        //// Overload based on row from schema table
        //public static string SQLGetType(DataRow schemaRow)
        //{
        //    int numericPrecision;
        //    int numericScale;

        //    if (!int.TryParse(schemaRow["NumericPrecision"].ToString(), out numericPrecision))
        //    {
        //        numericPrecision = -1;
        //    }
        //    if (!int.TryParse(schemaRow["NumericScale"].ToString(), out numericScale))
        //    {
        //        numericScale = -1;
        //    }

        //    return SQLGetType(schemaRow["DataType"],
        //                        int.Parse(schemaRow["ColumnSize"].ToString()),
        //                        numericPrecision,
        //                        numericScale);
        //}
        //// Overload based on DataColumn from DataTable type
        //public static string SQLGetType(DataColumn column)
        //{
        //    return SQLGetType(column.DataType, column.MaxLength, -1, -1);
        //}




        //public static string GetCreateFromDataTableSQL(string tableName, DataTable table)
        //{
        //    string sql = "CREATE TABLE [" + tableName + "] (\n";
        //    // columns
        //    foreach (DataColumn column in table.Columns)
        //    {
        //        sql += "[" + column.ColumnName + "] " + SQLGetType(column) + ",\n";
        //    }
        //    sql = sql.TrimEnd(new char[] { ',', '\n' }) + "\n";
        //    // primary keys
        //    if (table.PrimaryKey.Length > 0)
        //    {
        //        sql += "CONSTRAINT [PK_" + tableName + "] PRIMARY KEY CLUSTERED (";
        //        foreach (DataColumn column in table.PrimaryKey)
        //        {
        //            sql += "[" + column.ColumnName + "],";
        //        }
        //        sql = sql.TrimEnd(new char[] { ',' }) + "))\n";
        //    }

        //    return sql;
        //}


        //public static bool TableExists(this SqlCeConnection connection, string tableName)
        //{
        //    if (tableName == null) throw new ArgumentNullException("tableName");
        //    if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Invalid table name");
        //    if (connection == null) throw new ArgumentNullException("connection");
        //    if (connection.State != ConnectionState.Open)
        //    {
        //        throw new InvalidOperationException("TableExists requires an open and available Connection. The connection's current state is " + connection.State);
        //    }

        //    using (SqlCeCommand command = connection.CreateCommand())
        //    {
        //        command.CommandType = CommandType.Text;
        //        command.CommandText = "SELECT 1 FROM Information_Schema.Tables WHERE TABLE_NAME = @tableName";
        //        command.Parameters.AddWithValue("tableName", tableName);
        //        object result = command.ExecuteScalar();
        //        return result != null;
        //    }
        //}





        //  ***  end 7-25-14 try
        public void importDataFromExcel()
        {

            string sSQLTable = "table1";
            //this works
            string myExcelDataQuery = "SELECT Date, PID, SlopeDWN, SlopeUP, Number, Equip, BestHFR, FocusPos FROM [Sheet1$]";
            try
            {
                if (_importPath == null)
                {
                    DialogResult result = openFileDialog1.ShowDialog();
                    _importPath = openFileDialog1.FileName.ToString();
                    //  textBox34.Clear();
                    textBox34.Text = _importPath.ToString();
                    if (_importPath == "")
                        return;
                }
                string sSqlConnectionString = conString;
                string sExcelConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + _importPath + "';Extended Properties=Excel 8.0;"; //this works for VS 2013 and .xls.  for .xlsx may need if .xlsx then line below 
                                                                                                                                                     // see http://www.codeproject.com/Questions/227945/Import-Excel-data-in-textbox-using-oledb-conn-in-C


                string sClearSQL = "DELETE FROM " + sSQLTable;

                //  string SQLcreate = "CREATE TABLE table1 (Date datetime, PID int, SlopeDWN float(8,15), SlopeUP float(8,15), Number int, Equip nvarchar(100), BestHFR int, FocusPos int)";

                //   string sqlExists = "SELECT 1 FROM Information_Schema.Tables WHERE TABLE_NAME = @table1";


                SqlCeConnection SqlConn = new SqlCeConnection(conString);
                SqlCeCommand SqlCmd = new SqlCeCommand(sClearSQL, SqlConn);

                SqlConn.Open();
                SqlCmd.ExecuteNonQuery();

                SqlConn.Close();

                if (Path.GetExtension(_importPath.Substring(1)) == ".txt")
                {
                    string filepath = _importPath;
                    StreamReader sr = new StreamReader(filepath);
                    string line = sr.ReadLine();
                    string[] value = line.Split(',');
                    DataTable dt = new DataTable();
                    DataRow row;
                    foreach (string dc in value)
                    {
                        dt.Columns.Add(new DataColumn(dc));
                    }
                    using (SqlCeConnection con = new SqlCeConnection(conString))
                    {
                        con.Open();
                        //SqlCeCommand SqlCmd3 = new SqlCeCommand(sqlExists, con);
                        //if (SqlCmd3.ExecuteNonQuery() == 0)
                        //{
                        while (!sr.EndOfStream)
                        {
                            value = sr.ReadLine().Split(',');
                            if (value.Length == dt.Columns.Count)
                            {
                                row = dt.NewRow();
                                row.ItemArray = value;
                                dt.Rows.Add(row);



                                DateTime d = Convert.ToDateTime(row.ItemArray[0]);
                                int p = Convert.ToInt32(row.ItemArray[1]);
                                //  int num2 = _apexHFR;
                                //  int num4 = _posminHFR;
                                float down = Convert.ToSingle(row.ItemArray[2]);//********1-23-15  these looked backwards this line was up and next WAS down
                                float up = Convert.ToSingle(row.ItemArray[3]);
                                string test = row.ItemArray[5].ToString();
                                int hfrtest = Convert.ToInt16(row.ItemArray[6]);
                                int testintpos = Convert.ToInt32(row.ItemArray[7]);

                                using (SqlCeCommand com = new SqlCeCommand("INSERT INTO table1 (Date, PID, SlopeDWN, SlopeUP, Number, Equip, BestHFR, FocusPos) VALUES (@Date, @PID, @SlopeDWN, @SlopeUP, @Number, @equip, @BestHFR, @FocusPos)", con))
                                {

                                    com.Parameters.AddWithValue("@Date", d);
                                    com.Parameters.AddWithValue("@PID", p);
                                    com.Parameters.AddWithValue("@SlopeDWN", down);
                                    com.Parameters.AddWithValue("@SlopeUP", up);
                                    com.Parameters.AddWithValue("@Number", rows + 1);
                                    com.Parameters.AddWithValue("@equip", test);
                                    com.Parameters.AddWithValue("@BestHFR", hfrtest);
                                    com.Parameters.AddWithValue("@FocusPos", testintpos);
                                    com.ExecuteNonQuery();
                                    rows++;
                                }


                            }



                            // }
                        }
                        //else
                        //    MessageBox.Show("sql datatable does not exist, try reinstalling", "scopefocus");

                        con.Close();
                    }

                  //  MessageBox.Show("Data Import Successful", "scopefocus");
                    Log(Path.GetFileName(_importPath) + " import success" );
                    FileLog(Path.GetFileName(_importPath) + " import success");
                }
                else
                {



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
                 //   MessageBox.Show("Data Import Successful", "scopefocus");
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Import Failed - check file for errors", "scopefocus");
                FileLog2("Data import failed " + ex.ToString());
                return;
            }
        }

        private void button29_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox34_Click(object sender, EventArgs e)
        {
            if (textBox34.Text == "")
            {
                DialogResult result = openFileDialog1.ShowDialog();
                _importPath = openFileDialog1.FileName.ToString();
                textBox34.Text = _importPath.ToString();
                //Data d = new Data();
                // Filename = openFileDialog1.ToString();
            }
            else
                _importPath = textBox34.Text;

        }


        private void textBox33_Click(object sender, EventArgs e)
        {
            if (textBox33.Text == "")
            {
                DialogResult result = saveFileDialog1.ShowDialog();
                ExcelFilename = saveFileDialog1.FileName.ToString();
                textBox33.Text = ExcelFilename.ToString();
            }
            else
                ExcelFilename = textBox33.Text;
        }






        private void FindNebCamera()
        {
            try
            {

                //    try to get from the camera selection combobox...not working.  Captions for this windows is "" fro some reason
                //int panel = FindWindowByIndexName(Handles.NebhWnd, 3, "panel");

                //int camera = FindWindowByIndex(panel, 1, "combobox");//Finds thrid edit panel under parent panel under main window
                //IntPtr statusHandle = new IntPtr(camera);
                //StatusHelper sh = new StatusHelper(statusHandle);
                //string[] captions = sh.Captions;        






                // remd 11-8-16
                int StatusstripHandle = FindWindowEx(Handles.NebhWnd, 0, "msctls_statusbar32", null);

                IntPtr statusHandle = new IntPtr(StatusstripHandle);
                StatusHelper sh = new StatusHelper(statusHandle);
                // string[] captions = sh.Captions;
                if (sh.Captions[0] == "" || sh.Captions[0].Substring(0, 6) == "No cam") //"" if non selected after opening neb, the later is no camera specifically selected
                    NoCameraSelected();
                else
                {
                    Log(sh.Captions[0]);


                    /*
                    ******2-12-14 this no longer works for some reason.  changes to use statusbar instead  ********
                                    int panel3 = FindWindowByIndexName(Handles.NebhWnd, 4, "panel");// try panelhwnd
                                    //  Log("panel3 found" + panel3.ToString());
                                    int panel4 = FindWindowByIndexName(panel3, 1, "panel");//panel 4 is small one in panel 3, camera is 1st child in this
                                    //   Log("found panel4" + panel4);
                                    int camera = FindWindowByIndex(panel4, 1, null);
                                    //    Log("Found camera" + camera.ToString());
                                    StringBuilder sb = new StringBuilder(1024);
                                    SendMessage(camera, WM_GETTEXT, 1024, sb);
                                    //  GetWindowText(camera, sb, sb.Capacity);
                                    Log("Camera " + sb.ToString());
                     */
                    if (Handles.NebVNumber == 4)
                    {
                        int length = sh.Captions[0].Length;
                        string sb3 = sh.Captions[0].Remove(length - 10, 10);
                        GlobalVariables.Nebcamera = sb3.ToString();
                        textBox22.Text = GlobalVariables.Nebcamera.ToString();
                    }
                    else
                    {
                        string sb1 = sh.Captions[0].Remove(0, 7);
                        int indexofSpace = sb1.IndexOf(' ');
                        string sb2 = sb1.Substring(0, indexofSpace);
                        GlobalVariables.Nebcamera = sb2.ToString();
                        textBox22.Text = GlobalVariables.Nebcamera.ToString();
                    }
                    FileLog2("NebCamera: " + GlobalVariables.Nebcamera.ToString());
                }
                /*
               
                if (GlobalVariables.Nebcamera == null)
                {
                    NoCameraSelected();
                }
                 */


            }
            catch (Exception ex)
            {
                Log("Error finding Neb Camera");
                //  Send("FindNebCamera Error" + ex.ToString());
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
            {
                GlobalVariables.Nebcamera = "none"; // added 11-3-16
                return;
            }
        }
        public int MaxADU;
        private void FindADU()
        {
            try
            {
                msdelay(500);
                SetForegroundWindow(Handles.NebhWnd);
                msdelay(500);
                MaxADU = 0; //need to rest for seuqential determiniations otherwise will break out of loop evertyine after 1st MaxADU found
                for (int i = 1; i < 8; i++)
                {
                    int ADUpanel = FindWindowByIndexName(Handles.NebhWnd, i, "panel");// try panelhwnd  (was 5 for Neb v3)  ***  11-13-16 might be different on differt computers, cycle through til found.  
                                                                                      //    Log("panel5 found" + ADUpanel.ToString());
                    int MaxADUbox = FindWindowByIndex(ADUpanel, 2, "edit");//Finds second edit panel under parent panel under main window
                                                                           //    Log("found MaxADU box" + MaxADUbox.ToString());


                    StringBuilder sb = new StringBuilder(1024);
                    SendMessage(MaxADUbox, WM_GETTEXT, 1024, sb);
                    //  GetWindowText(camera, sb, sb.Capacity);

                    if (sb.Length != 0)
                    { 
                    string ADU = sb.ToString();
                    MaxADU = Convert.ToInt32(ADU);
                   

                    }
                    // Log("MaxADU: " + MaxADU.ToString() + "Exposure Time: " + FlatExp.ToString());
                    if ((MaxADU < 65000) && (MaxADU > 1))
                        break;
                }
            }

            catch (Exception ex)
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
        // public bool FlatCalcDone = false;
        //   public bool SlaveFlatCalcDone = false;
        bool clipboardRetry = false;
        private void CalculateFlatExp()//calcflatexphere
        {
            FileLog2("Flat Calc");
            if (!clipboardRetry) // try 11-15-16
            {
                if (IsSlave())
                    textBox41.Clear();
                toolStripStatusLabel1.Text = "Calculating Flat Exposure";
                this.Refresh();
                if ((IsSlave()) & (FlatExp == 0))//allows second set of flats to start at calculated time
                    FlatExp = Convert.ToDouble(textBox40.Text) / 1000;
                if (FlatExp == 0)
                    FlatExp = (double)numericUpDown34.Value / 1000;
                MaxADU = 0;

                fileSystemWatcher4.EnableRaisingEvents = false;
                /*
                if (NebVNumber == 3)
                    CaptureTime3 = (double)CaptureTime / 1000;//need to allow fractions for flats.
               */
                // CaptureTime3 = FlatExp;


                tol = 1 + (Convert.ToDouble(textBox39.Text.ToString()) / 100);
                FlatGoal = Convert.ToInt32(textBox38.Text.ToString()) * 1000;
                //    if (NebListenOn == false)
                NebListenStart(Handles.NebhWnd, SocketPort);
            }
            msdelay(750); // 11-8-16
            while ((MaxADU > FlatGoal * tol) || (MaxADU < FlatGoal / tol))
            {
                textBox41.Refresh();
               // Log("while ADU not found");
                //  toolStripStatusLabel1.Text = "Capturing";
                //   this.Refresh();
                string prefix = textBox19.Text.ToString();
                //  NetworkStream serverStream;
                if (!UseClipBoard.Checked) // 11-8-16
                {

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
                            int StatusstripHandle = FindWindowEx(Handles.NebhWnd, 0, "msctls_statusbar32", null);

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
                }

                else
                {
                    int subs = 1;
                    string name = "FlatCalc";
                    // byte[] outStream = System.Text.Encoding.ASCII.GetBytes("setname " + prefix + name + "\n" + "setbinning " + CaptureBin + "\n" + "SetShutter 0" + "\n" + "SetDuration " + FlatExp + "\n" + "Capture " + subs + "\n");
                    try
                    {
                        Capturing = true;  // remd 11-13-16 with rem below.  
                    //    int i = 0;
                        //  serverStream.Write(outStream, 0, outStream.Length);
                        msdelay(750);
                        Clipboard.Clear();
                        msdelay(500);
                        //   Clipboard.SetText("//NEB SetName " + prefix + name);
                        Clipboard.SetDataObject("//NEB SetName " + prefix + name, false, 3, 500);
                        msdelay(750);
                        //if (!NebCommandConfirm("SetName " + prefix + name, 0))
                        //    if (!clipboardRetry)
                        //    {
                        //        clipboardRetry = true;
                        //        CalculateFlatExp();
                        //    }
                        //    else
                        //        ClipboardFailure("SetName");
                      

                        //   Clipboard.SetText("//NEB SetBinning " + CaptureBin);
                        Clipboard.SetDataObject("//NEB SetBinning " + CaptureBin, false, 3, 500);
                        msdelay(750);
                        //if (!NebCommandConfirm("SetBinning " + CaptureBin, 0))
                        //    if (!clipboardRetry)
                        //    {
                        //        clipboardRetry = true;
                        //        CalculateFlatExp();
                        //    }
                        //    else
                        //        ClipboardFailure("SetBinning");
                        //while (!NebCommandConfirm("SetBinning " + CaptureBin, 0))
                        //{
                        //    i++;
                        //    msdelay(100);
                        //    if (i == 30)
                        //    {
                        //        Log("FlatCalc Capture Failed");
                        //        Send("FlatCalc Caputre Failed");
                        //        return;
                        //    }
                        //}

                        //   Clipboard.SetText("//NEB SetShutter 0");
                        Clipboard.SetDataObject("//NEB SetShutter 0", false, 3, 500);
                        msdelay(750);
                        //if (!NebCommandConfirm("SetShutter 0", 0))
                        //    if (!clipboardRetry)
                        //    {
                        //        clipboardRetry = true;
                        //        CalculateFlatExp();
                        //    }
                        //    else
                        //        ClipboardFailure("SetShutter");
                        //while (!NebCommandConfirm("SetShutter 0", 0))
                        //{
                        //    i++;
                        //    msdelay(100);
                        //    if (i == 30)
                        //    {
                        //        Log("FlatCalc Capture Failed");
                        //        Send("FlatCalc Caputre Failed");
                        //        return;
                        //    }
                        //}

                        //     Clipboard.SetText("//NEB SetDuration " + FlatExp);
                        Clipboard.SetDataObject("//NEB SetDuration " + FlatExp, false, 3, 500);
                        msdelay(750);
                        //if (!NebCommandConfirm("SetDuration " + FlatExp, 0))
                        //    if (!clipboardRetry)
                        //    {
                        //        clipboardRetry = true;
                        //        CalculateFlatExp();
                        //    }
                        //    else
                        //        ClipboardFailure("SetDuration");
                        //while (!NebCommandConfirm("SetDuration " + FlatExp, 0))
                        //{
                        //    i++;
                        //    msdelay(100);
                        //    if (i == 30)
                        //    {
                        //        Log("FlatCalc Capture Failed");
                        //        Send("FlatCalc Caputre Failed");
                        //        return;
                        //    }
                        //}
                        //i = 0;
                        //  Clipboard.SetText("//NEB Capture " + subs);
                        Clipboard.SetDataObject("//NEB Capture " + subs, false, 3, 500);
                        msdelay(200);

                       
                        //if (!NebCommandConfirm("Exposing", 1))
                        //    if (!clipboardRetry)
                        //    {
                        //        clipboardRetry = true;
                        //        CalculateFlatExp();
                        //    }
                        //    else
                        //        ClipboardFailure("Exposing");
                        // todo confirm 

                        Clipboard.Clear();

                        // wait here until done.....
                        while (!NebCommandConfirm("Saving:", 1))
                        {
                            msdelay(200);
                          //  Log("Waiting for exposure");
                        }

                        //   msdelay(750);
                       System.Threading.Tasks.Task.Delay(5000).Wait();
                      //  delay(5);
                       // Thread.Sleep(5000);
                        // 11-13-16 try this instead of below.  
                        //while (!NebCommandConfirm("Capture " + subs, 0))
                        //{
                        //    i++;
                        //    msdelay(100);
                        //    if (i == 30)
                        //    {
                        //        Log("FlatCalc Capture Failed");
                        //        Send("FlatCalc Caputre Failed");
                        //        return;
                        //    }
                        //}



                        //Thread.Sleep(500);//wait to check until after starts capturing 
                        //while (Capturing == true)  //**** this no longer works w/ clipborad
                        //{
                        //    int StatusstripHandle = FindWindowEx(Handles.NebhWnd, 0, "msctls_statusbar32", null);

                        //    //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
                        //    IntPtr statusHandle = new IntPtr(StatusstripHandle);
                        //    StatusHelper sh = new StatusHelper(statusHandle);
                        //    string[] captions = sh.Captions;
                        //    if (captions[0] == "Sequence done")
                        //        Capturing = false;
                        //}




                    }

                    catch
                    {
                        MessageBox.Show("Error sending command", "scopefocus");
                        return;

                    }
                }

              
                FindADU();
                Log("MaxADU " + MaxADU.ToString());
                FileLog2("MaxADU: " + MaxADU.ToString() + "    Exp time:  " + Math.Round(FlatExp, 3).ToString());
                if ((MaxADU < FlatGoal * tol) && (MaxADU > FlatGoal / tol))//this gets out fo the loop before adusting exp time
                {
                    if (IsSlave())
                    {
                        textBox41.Text = "Slave FlatCalc Complete";
                        textBox41.Refresh();
                    }
                    Log("Flat Calc Complete--MaxADU: " + MaxADU.ToString() + "   Exposure Time: " + Math.Round(FlatExp,3).ToString());
                    FileLog2("Flat Calc Complete--MaxADU: " + MaxADU.ToString() + "   Exposure Time: " + Math.Round(FlatExp, 3).ToString());
                    CaptureTime3 = FlatExp;
                    FlatCalcDone = true;
                    break;
                }
                ADUratio = (double)FlatGoal / (double)MaxADU;
                FlatExp = FlatExp * ADUratio;


            } // end while
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
                    SendMessage(Handles.SlaveStatushwnd, WM_GETTEXT, 1024, sb);
                    //  GetWindowText(camera, sb, sb.Capacity);
                    //  Log("SlaveStatus= " + sb.ToString());
                    if (sb.ToString() == "Slave FlatCalc Complete")
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
            GlobalVariables.NebPath = folderBrowserDialog2.SelectedPath.ToString();
            textBox35.Text = GlobalVariables.NebPath;
        }
        //std dev button // remd 10-25-16
        //private void Button2000_Click(object sender, EventArgs e)
        //{
        //    //   port.DiscardInBuffer();
        //    // port.DiscardOutBuffer();
        //    fileSystemWatcher2.EnableRaisingEvents = false;
        //    fileSystemWatcher5.EnableRaisingEvents = false; //added to test metricHFR
        //    fileSystemWatcher3.EnableRaisingEvents = true;
        //    fileSystemWatcher1.EnableRaisingEvents = false;
        //    list = new int[arraysize2];
        //    abc = new double[arraysize2];

        //    posMin = 0;
        //    min = 1;
        //    sum = 0;
        //    avg = 0;
        //    vDone = 0;
        //    vProgress = 0;
        //}



        private void button19_Click_1(object sender, EventArgs e)
        {
            //  ResumePHD();
        }
        /*
        private void button22_Click_1(object sender, EventArgs e)
        {
            StopTracking();
        }
         */
        /*
        private void button30_Click(object sender, EventArgs e)
        {
            ResumeTracking();
        }
        */
        /*
        private void ResumeTracking()
        {
            try
            {
                if (TrackingOn == true)
                {
                    MessageBox.Show("Tracking Already On", "scoprefocus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                scope.Tracking = true;
                //if (port2.IsOpen == false)
                //    Port2Open();
                //port2.DiscardInBuffer();
                //port2.DiscardOutBuffer();
                //Thread.Sleep(50);
                //byte[] newMsg = { 0x54, 0x02 };
                //port2.Write(newMsg, 0, newMsg.Length);
                //Thread.Sleep(100);
                //port2.DiscardOutBuffer();

                //Thread.Sleep(50);
                //string Handshake = port2.ReadExisting();
                //Thread.Sleep(50);
                //if (Handshake == "#") //check for handshake
                //{
                //    Log("Tracking Command Received");
                //}

                //port2.DiscardInBuffer();
                //port2.DiscardOutBuffer();
                //Thread.Sleep(50);
                //port2.Write("t");//check tracking mode
                //Thread.Sleep(100);
                //port2.DiscardOutBuffer();
                //// byte[] confirm = new byte[1];
                //int confirm;
                //confirm = port2.ReadByte();
                //Thread.Sleep(20);
                //if (confirm == 2)
                //{
                //    Log("Tracking EQ-North confirmed");
                //    TrackingOn = true;
                //}
                //else
                //{
                //    if (retry < 3)
                //    {
                //        retry++;
                //        Thread.Sleep(1000);
                //        ResumeTracking();
                //    }
                //    else
                //    {
                //        retry = 0;
                //        MessageBox.Show("Unable to Resume Tracking", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    }
                //}
                //port2.Close();
            }
            catch (Exception ex)
            {
                Log("Resume Tracking Error" + ex.ToString());
                Send("Resume Tracking Error" + ex.ToString());
                FileLog("Resume Tracking Error" + ex.ToString());
            }
        }
        */
        //  AstrometryNet ast = new AstrometryNet();
        public async void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {

                //AstrometryNet ast = new AstrometryNet();
                //await ast.OnlineSolve(GlobalVariables.SolveImage);

            }
            catch (Exception ex)
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


        public void Send(string msg) //sendhere
        {
            try
            {
                if (checkBox12.Checked == true || checkBox30.Checked == true)
                {
                    this.user = WindowsFormsApplication1.Properties.Settings.Default.user;
                    if (user == "")
                    {
                        if (textBox27.Text == "")
                        {
                            MessageBox.Show("Username is blank, enter username and try again", "scopefocus");
                            return;
                        }
                        else
                            user = textBox27.Text;
                    }

                    MailMessage Mail = new MailMessage();

                    Mail.Subject = ("scopefocus - Error");

                    Mail.Body = (msg);

                    Mail.BodyEncoding = Encoding.GetEncoding("Windows-1254"); // Turkish Character Encoding

                    Mail.From = new MailAddress(user);

                    Mail.To.Add(new MailAddress(to));

                    System.Net.Mail.SmtpClient Smtp = new SmtpClient();

                    Smtp.Host = (server); // for example gmail smtp server

                    Smtp.EnableSsl = true;
                    Smtp.Port = 587;
                    textBox27.Text = user.ToString();
                    // Smtp.Credentials = new System.Net.NetworkCredential("ksipp911@gmail.com", "cloe$1124");
                    Smtp.Credentials = new System.Net.NetworkCredential(user, pswd);

                    Smtp.Send(Mail);
                    FileLog("message sent to " + to.ToString());
                    Log("message sent to " + to.ToString());
                }
                else
                    return;
            }
            catch (SmtpFailedRecipientException ex)
            {
                Log("Send Mail Error" + ex.ToString());
                //  Send("Send Mail Error" + ex.ToString());
                FileLog("Send Mail Error" + ex.ToString());
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
            if (checkBox12.Checked == false)
            {
                checkBox12.Checked = true;
                Send("test message");
                Log("test message sent");
                checkBox12.Checked = false;
            }
            else
            {
                Send("test message");
                Log("test message sent");
            }

        }

        public void FileLog2(string textlog)  //for non-error file log
        {

            if (checkBox33.Checked)
            {
                //  string strLogText = "Std Dev" + "\t  " + abc[vProgress].ToString() + "\t  " + avg.ToString() + "\t" + (vProgress + 1).ToString() + "\t" + stdev.ToString();
                //  string path = textBox11.Text.ToString();

                string fullpath = GlobalVariables.Path2 + @"\scopefocusLog_" + DateTime.Now.ToString("yyy-M-dd") + ".txt";
                StreamWriter log;
                if (!File.Exists(fullpath))
                {
                    log = new StreamWriter(fullpath);
                }
                else
                {
                    log = File.AppendText(fullpath);
                }

                //  log.WriteLine(DateTime.Now);
                //  log.WriteLine("scopefocus - Error");
                log.WriteLine(textlog);
                log.Close();
            }
            return;
        }


        public void FileLog(string textlog)
        {
            //  string strLogText = "Std Dev" + "\t  " + abc[vProgress].ToString() + "\t  " + avg.ToString() + "\t" + (vProgress + 1).ToString() + "\t" + stdev.ToString();
            //  string path = textBox11.Text.ToString();

            string fullpath = GlobalVariables.Path2 + @"\scopefocusLog_" + DateTime.Now.ToString("yyy-M-dd") + ".txt";
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

        //     private void Bitstuff()
        //      {
        //  Bitmap bmp = new Bitmap("c:\\atest3\\Focus_3556_9685_1.76_1.04.bmp");
        //    int height = bmp.Size.Height;
        //    int width = bmp.Size.Width;
        //textBox37.Text = height.ToString();


        //       }

        private void button21_Click(object sender, EventArgs e)
        {
            ToggleFlat();
        }

        private void ToggleFlat()
        {
            if (devId4 != null)
            {
                if (!FlatFlap.GetSwitch(0))
                {
                    FlatFlap.SetSwitch(0, true);
                    FlatsOn = true;
                    button21.Text = "Flat On";
                    button21.BackColor = System.Drawing.Color.Lime;
                }

                else
                {
                    FlatFlap.SetSwitch(0, false);
                    FlatsOn = false;
                    button21.Text = "Flat Off";
                    button21.BackColor = System.Drawing.Color.Red;
                }
            }
            else
            {

                // gotopos(3);
                //focuser.Move(3);
                focuser.Action("FlatToggle", "");
                Thread.Sleep(200);

                //***************need to add to driver ***********

                //toolStripStatusLabel1.Text = "Flat Panel Moving";
                //this.Refresh();
                //string movedone;
                //port.DiscardOutBuffer();
                //port.DiscardInBuffer();
                //Thread.Sleep(20);
                //port.Write("R");
                //filterMoving = true;
                //Thread.Sleep(50);

                //while (filterMoving == true)
                //{


                //    movedone = port.ReadExisting();
                //    if (movedone == "D")
                //    {

                //        filterMoving = false;
                //        this.Refresh();
                //        toolStripStatusLabel1.Text = "Flat Move Complete";
                //        this.Refresh();
                //        //   toolStripStatusLabel1.Text = " ";
                //        //    toolStripStatusLabel1.Text = " ";

                //    }
                //}
                //Thread.Sleep(50);
                //port.DiscardInBuffer();
                //if (FlatsOn == false)
                //{
                //    FlatsOn = true;
                //    button21.Text = "Flat On";
                //    button21.BackColor = System.Drawing.Color.Lime;


                //}

                //else
                //{
                //    FlatsOn = false;
                //    button21.Text = "Flat Off";
                //    button21.BackColor = System.Drawing.Color.Red;
                //}
            }
        }

        //Add PHD monitor for lost guidesstar
        //can use as built in soundalert finction AND compensate for meridian flip lost guide star
        //added to resumePHd() 4-9-12.  This is just for testing purpose w/ button 39 o/w not used. 

        //**********also consider this:
        //  - (1.13.5) New server command (id=14) to auto-find star (same as pulling down this from the menu)

        //  public double PHDtime = 0.0;
        //  public double DX = 0.0;
        //  public double DY = 0.0;
        //private void PHDCheckSNR()
        //{

        //    try
        //    {
        //        int StatusstripHandle = FindWindowEx(Handles.PHDhwnd, 0, "msctls_statusbar32", null);

        //        //    from   http://www.pinvoke.net/default.aspx/user32/SB_GETTEXT.html 
        //        IntPtr statusHandle = new IntPtr(StatusstripHandle);
        //        StatusHelper sh = new StatusHelper(statusHandle);
        //        string[] captions = sh.Captions;
        //        //   if (captions[0] == "Guiding")//******must rem for sim debugging *****
        //        if (captions[1] != null)
        //            if (captions[1].Substring(0, 3) == "LOW")
        //            {
        //                MessageBox.Show("LOW SNR");
        //            }
        //    }
        //    catch (Exception ex)
        //    {

        //        Log("PHD Low SNR error" + ex.ToString());
        //        Send("PHD Low SNR error" + ex.ToString());
        //        FileLog("PHD Low SNR error" + ex.ToString());
        //    }
        //}

        private void textBox4_Leave(object sender, EventArgs e)
        {
            posMin = Convert.ToInt32(textBox4.Text);//allows direct entry of close focus position
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
        /*
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
         */
        //pause  doesn't work
        /*
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
         */
        //stop doesn't work?? 
        /*
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
         */
        //pause filter tab
        bool paused;
        //this won't work until nebcapture is in background
        private void button22_Click_2(object sender, EventArgs e)
        {
            if (paused == false)
            {
                fileSystemWatcher4.EnableRaisingEvents = false;
                button22.BackColor = System.Drawing.Color.Lime;
                paused = true;
                button22.Text = "Resume";

                // PHDSocketPause(true);
                // 12-7-16 
                //PHDcommand(PHD_PAUSE);
                ph.StopCapture();
                Log("guiding paused");

                // 11-8-16
                //TODO 
                //hit NEb abort button to stop.  
                // determine howm many done gobal.capcurrent versus global.capTotal
                // also if fucusing at intervals need to figure remander and when to focus again.  
                // then put remining in the numericupdown next to the filte selection so hitting resume is just like hitting start again.
                //focusing interval is biggest problem......

            }

            else
            {
                fileSystemWatcher4.EnableRaisingEvents = true;
                button22.UseVisualStyleBackColor = true;
                paused = false;
                button22.Text = "Pause";
                if (Handles.PHDVNumber == 2)
                    resumePHD2();
                else
                    ph.Guide();
                // 12-7-16
                  //  PHDcommand(PHD_RESUME);
            }
        }
        //stop filter tab
        private void button30_Click_1(object sender, EventArgs e)
        {
            FileLog2("Abort pushed");
            DialogResult result;
            result = MessageBox.Show("Abort the current sequence?", "scopefocus",
                            MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
            if (result == DialogResult.OK)
            {
                backgroundWorker2.CancelAsync();// added 1-23-13


                ShowWindowAsync(Handles.NebhWnd, SW_SHOW);
                Thread.Sleep(200);
                SetForegroundWindow(Handles.Aborthwnd);
                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);

                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(500);
                SendKeys.SendWait("~"); // this enter key clears frequent failed to read clipboard error
                SendKeys.Flush();


                fileSystemWatcher4.EnableRaisingEvents = false;
                filterCountCurrent = 0;
                subCountCurrent = 0;
                standby();

                //  MessageBox.Show("Close and restart Nebulosity, then click rescan on setup tab");

                // 10-25-16 try (doesn't work get socket closed then error on restart seqeunce

                //serverStream = clientSocket.GetStream();
                //byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                //serverStream.Write(outStream, 0, outStream.Length);
                //serverStream.Flush();

                //Thread.Sleep(3000);
                //serverStream.Close();
                //SetForegroundWindow(Handles.NebhWnd);
                //Thread.Sleep(1000);
                //PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                //Thread.Sleep(1000);
                //NebListenOn = false;
                //// clientSocket.GetStream().Close();//added 5-17-12
                ////  clientSocket.Client.Disconnect(true);//added 5-17-12
                //clientSocket.Close();





                //clientSocket2.Close();
                //Thread.Sleep(250);
                //SendKeys.SendWait("~");
                //SendKeys.Flush();


                // 10-25-16 addition


                //serverStream = clientSocket.GetStream();
                //byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                //serverStream.Write(outStream, 0, outStream.Length);
                //serverStream.Flush();
                toolStripStatusLabel1.Text = "Sequence Aborted";
                toolStripStatusLabel1.BackColor = Color.WhiteSmoke;
                button26.UseVisualStyleBackColor = true;
                this.Refresh();




                //fileSystemWatcher4.EnableRaisingEvents = false;
                ////NetworkStream serverStream = clientSocket.GetStream();
                //serverStream = clientSocket.GetStream();
                //byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                //serverStream.Write(outStream, 0, outStream.Length);
                //serverStream.Flush();
                //toolStripStatusLabel1.Text = "Sequence Aborted";
                //toolStripStatusLabel1.BackColor = Color.WhiteSmoke;
                //this.Refresh();
                //if (DarksOn == true)
                //    DarksOn = false;
                //currentfilter = 1;
                //DisplayCurrentFilter();
                ////      currentfilter = 0;
                //subCountCurrent = 0;
                //filterCountCurrent = 0;
                //Thread.Sleep(3000);//prevent overlapping sounds
                //serverStream.Close();
                //SetForegroundWindow(Handles.NebhWnd);
                //Thread.Sleep(1000);
                //PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                //Thread.Sleep(3000);
                //NebListenOn = false;
                //// clientSocket.GetStream().Close();//added 5-17-12
                ////  clientSocket.Client.Disconnect(true);//added 5-17-12
                //clientSocket.Close();
                //toolStripStatusLabel4.BackColor = System.Drawing.Color.LightGray;
                //toolStripStatusLabel4.Text = Filtertext.ToString();
                ////  button12.UseVisualStyleBackColor = true;
                ////   button31.UseVisualStyleBackColor = true;








                // ************ end 11-23-13 addt

            }
            if (result == DialogResult.Cancel)
                return;
            EnableAllUpDwn(); //3-2-16
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
            //    ******   11-6-13   ********
            //*************this will need to be changed to universal application**************
            //********currently just a work around for my setup. *****
            //could setup table like maxtravel for each item in the cobobox7 (epquip) list
            //   OR save Maxtravel as string, add R to the end to signal reverse.  would need
            //to parse for the R, trim it, then convert to Int.....


            if (comboBox7.SelectedItem.ToString() == "FSQ-85")
                checkBox28.Checked = true;



        }

        int checkboxs;
        public int TotalSubs()
        {
            int totalsubs;
            checkboxs = 0;
            if (checkBox1.Checked == true)
                checkboxs++;
            if (checkBox2.Checked == true)
                checkboxs++;
            if (checkBox3.Checked == true)
                checkboxs++;
            if (checkBox4.Checked == true)
                checkboxs++;
            if (checkBox5.Checked == true)
                if (comboBox1.SelectedItem.ToString() != "Dark 1")
                    checkboxs++;
            if (checkBox9.Checked == true)//*** 4-8-14 was 6 ****
                if ((comboBox8.SelectedItem.ToString() != "Dark 1") || (comboBox1.SelectedItem.ToString() != "Dark 2")) // changed to 'or' 11-16-16
                    checkboxs++;
            if (checkBox18.Checked == true)
                totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + ((int)numericUpDown32.Value * checkboxs);
            else
                totalsubs = (int)numericUpDown12.Value + (int)numericUpDown13.Value + (int)numericUpDown14.Value + (int)numericUpDown15.Value + (int)numericUpDown20.Value + (int)numericUpDown29.Value + (int)numericUpDown32.Value;
            textBox23.Text = totalsubs.ToString();
            
            return totalsubs;
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
        private double[] FocusGroup = new double[6];
        private int[] SubsPerFocus = new int[6];
        //this likely needs to be removed  *******   4-8-14  *******
        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            /*
            if ((checkBox8.Checked == true) & (checkBox17.Checked == true))
                MessageBox.Show("Confirm refocus after filter change AND " + numericUpDown38.Value.ToString() + " subs?"
                    , "scopefocus");
             */
            if (checkBox17.Checked == true)
            {
                checkBox8.Checked = true;
                MessageBox.Show("Filter change focus implied with N-subs focus", "scopefocus");
            }
        }






        //this likely needs to be removed  *****  4-8-14  ******
        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {
           // if (checkBox17.Checked)
           //     FocusGroupCalc(); //shouel do it either way that way gets reset if unchecked
           // 11-18-16 try just doing at sequence go.....

            
            // 11-18-16 rem'd all of this.  does it when numericupdown38.value changed

            //if (checkBox17.Checked)
            //{
            //    //if (numericUpDown38.Value == 0) // remd 10-24-16
            //    //    MessageBox.Show("Value cannot be zero", "scopefoucs");
            //    //if (numericUpDown38.Value < 4) // remd 10-24-16
            //    //    MessageBox.Show("Confrim low refocus per sub value of " + numericUpDown38.Value.ToString(), "scopefocus");
            //    FocusPerSub = (int)numericUpDown38.Value;
            //    if (checkBox1.Checked == true)
            //    {
            //        if (numericUpDown12.Value == 0)
            //        {
            //            MessageBox.Show("Position 1 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            return;
            //        }
            //        FocusGroup[0] = (double)numericUpDown12.Value / FocusPerSub;
            //        if (FocusGroup[0] != Math.Round(FocusGroup[0]))
            //        {
            //            MessageBox.Show("Total pos. 1 subs must be multiple of Focus per sub", "scopefocus");
            //            numericUpDown38.Value = 1;
            //        }
            //        else
            //            SubsPerFocus[0] = (int)numericUpDown12.Value / (int)FocusGroup[0];
            //        if (FocusGroup[0] == 1)//then just use the filter change to refocus
            //        {
            //            SubsPerFocus[0] = 0;
            //            FocusGroup[0] = 0;
            //        }
            //    }
            //    if (checkBox2.Checked == true)
            //    {
            //        if (numericUpDown13.Value == 0)
            //        {
            //            MessageBox.Show("Position 2 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            return;
            //        }
            //        FocusGroup[1] = (double)numericUpDown13.Value / FocusPerSub;
            //        if (FocusGroup[1] != Math.Round(FocusGroup[1]))
            //        {
            //            MessageBox.Show("Total pos. 2 subs must be multiple of Focus per sub", "scopefocus");
            //            numericUpDown38.Value = 1;
            //        }
            //        else
            //            SubsPerFocus[1] = (int)numericUpDown13.Value / (int)FocusGroup[1];
            //        if (FocusGroup[1] == 1)//then just use the filter change to refocus
            //        {
            //            SubsPerFocus[1] = 0;
            //            FocusGroup[1] = 0;
            //        }
            //    }
            //    if (checkBox3.Checked == true)
            //    {
            //        if (numericUpDown14.Value == 0)
            //        {
            //            MessageBox.Show("Position 3 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            return;
            //        }
            //        FocusGroup[2] = (double)numericUpDown14.Value / FocusPerSub;
            //        if (FocusGroup[2] != Math.Round(FocusGroup[2]))
            //        {
            //            MessageBox.Show("Total pos.3 subs must be multiple of Focus per sub", "scopefocus");
            //            numericUpDown38.Value = 1;
            //        }
            //        else
            //            SubsPerFocus[2] = (int)numericUpDown14.Value / (int)FocusGroup[2];
            //        if (FocusGroup[2] == 1)//then just use the filter change to refocus
            //        {
            //            SubsPerFocus[2] = 0;
            //            FocusGroup[2] = 0;
            //        }
            //    }
            //    if (checkBox4.Checked == true)
            //    {
            //        if (numericUpDown15.Value == 0)
            //        {
            //            MessageBox.Show("Position 4 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            return;
            //        }
            //        FocusGroup[3] = (double)numericUpDown15.Value / FocusPerSub;
            //        if (FocusGroup[3] != Math.Round(FocusGroup[3]))
            //        {
            //            MessageBox.Show("Total pos. 4 subs must be multiple of Focus per sub", "scopefocus");
            //            numericUpDown38.Value = 1;
            //        }
            //        else
            //            SubsPerFocus[3] = (int)numericUpDown15.Value / (int)FocusGroup[3];
            //        if (FocusGroup[3] == 1)//then just use the filter change to refocus
            //        {
            //            SubsPerFocus[3] = 0;
            //            FocusGroup[3] = 0;
            //        }
            //    }
            //    if (checkBox5.Checked == true)
            //    {
            //        if (numericUpDown20.Value == 0)
            //        {
            //            MessageBox.Show("Position 5 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            return;
            //        }
            //        FocusGroup[4] = (double)numericUpDown20.Value / FocusPerSub;
            //        if (FocusGroup[4] != Math.Round(FocusGroup[4]))
            //        {
            //            MessageBox.Show("Total pos. 4 subs must be multiple of Focus per sub", "scopefocus");
            //            numericUpDown38.Value = 1;
            //        }
            //        else
            //            SubsPerFocus[4] = (int)numericUpDown20.Value / (int)FocusGroup[4];
            //        if (FocusGroup[4] == 1)//then just use the filter change to refocus
            //        {
            //            SubsPerFocus[4] = 0;
            //            FocusGroup[4] = 0;
            //        }
            //    }
            //    if (checkBox9.Checked == true)
            //    {
            //        if (numericUpDown29.Value == 0)
            //        {
            //            MessageBox.Show("Position 6 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            return;
            //        }
            //        FocusGroup[5] = (double)numericUpDown29.Value / FocusPerSub;
            //        if (FocusGroup[5] != Math.Round(FocusGroup[3]))
            //        {
            //            MessageBox.Show("Total pos. 4 subs must be multiple of Focus per sub", "scopefocus");
            //            numericUpDown38.Value = 1;
            //        }
            //        else
            //            SubsPerFocus[5] = (int)numericUpDown29.Value / (int)FocusGroup[5];
            //        if (FocusGroup[5] == 1)//then just use the filter change to refocus
            //        {
            //            SubsPerFocus[5] = 0;
            //            FocusGroup[5] = 0;
            //        }
            //    }
            //    FileLog2("FocusSubsPerGroup: " + FocusGroup[0] + FocusGroup[1] + FocusGroup[2] + FocusGroup[3] + FocusGroup[4] + FocusGroup[5]);
            //    FocusPerSubGroupCount = (int)(FocusGroup[0] + FocusGroup[1] + FocusGroup[2] + FocusGroup[3] + FocusGroup[4] + FocusGroup[5]);
            //}


            ///*
            //if ((checkBox8.Checked == true) & (checkBox17.Checked == true))
            //    MessageBox.Show("Confirm refocus after filter change AND " + numericUpDown38.Value.ToString() + " subs?"
            //        , "scopefocus");
            //*/
            //if (checkBox17.Checked == true)
            //    checkBox8.Checked = true;
        }

        private void numericUpDown38_Leave(object sender, EventArgs e)
        {
         //   FocusGroupCalc();  // 11-18-16 try just doing at sequence go
            // 10-24-16 moved to value changed method at end
        }
        //  doflathere
        private void DoFlat()
        {
            FileLog2("DoFlat");
            if ((IsSlave()) & (numericUpDown37.Value == 0))
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
            Handles.Editfound = 0;
            Nebname = Nebname + "_Flat";
            if (IsSlave())
            {
                subsperfilter = (int)numericUpDown37.Value;
                CaptureTime = Convert.ToInt32(textBox40.Text);//starting exposure time
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
                        WaitForSequenceDone("Sequence done", GlobalVariables.NebSlavehwnd);
                        // Thread.Sleep(50);
                    }
                    working = true;
                    Log("Slave flat done");

                }

            }
            else
            {
                //**** 4-22-14  this was added since nebcapture is now in background, need to wait until flats are done then filter change

                this.Refresh();
                while (working == true)
                {
                    WaitForSequenceDone("Sequence done", Handles.NebhWnd);
                    // Thread.Sleep(50);
                }
                working = true;
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
                if (Handles.PHDVNumber == 2)
                    resumePHD2();
                else
                    ResumePHD();
            if (!IsSlave())
            {
                subCountCurrent = subCountCurrent + (int)numericUpDown32.Value;
                fileSystemWatcher4.EnableRaisingEvents = true;

                checkfiltercount();// added 5-23 this is needed to end sequence if no darks selected
            }
        }

        //import button
        private void button29_Click(object sender, EventArgs e)
        {
            if (textBox34.Text == "")
            {
                DialogResult result = openFileDialog1.ShowDialog();
                _importPath = openFileDialog1.FileName.ToString();
                textBox34.Text = _importPath.ToString();
                return;
            }
            else
            {
                //  Data d = new Data();
                _importPath = textBox34.Text;
                importDataFromExcel();
                button17.PerformClick();//update
            }
        }
        //export
        private void button42_Click_1(object sender, EventArgs e)
        {
            try
            {
                //if (ExcelFilename == null)
                //{
                //    DialogResult result = saveFileDialog1.ShowDialog();
                //    ExcelFilename = saveFileDialog1.FileName.ToString();
                //    textBox33.Text = ExcelFilename.ToString();
                //}
                if (textBox33.Text == "")
                {
                    DialogResult result = saveFileDialog1.ShowDialog();
                    //   ExcelFilename = saveFileDialog1.FileName.ToString();
                    //  textBox33.Text = ExcelFilename.ToString();
                    return;
                }
                else
                    ExcelFilename = textBox33.Text;
                using (SqlCeConnection con = new SqlCeConnection(conString))
                {
                    con.Open();
                    using (SqlCeDataAdapter a = new SqlCeDataAdapter("SELECT * FROM table1", con))
                    {
                        DataTable t = new DataTable();
                        a.Fill(t);
                        // dataGridView1.DataSource = t;
                        a.Update(t);
                        object missing = System.Reflection.Missing.Value;
                        if (ExcelFilename.IndexOf(@".") == 0)//if extension not on filename
                        {
                            MessageBox.Show("Filename extension not selected, must select text or Excel file.  Export Aborted", "scopefocus");
                            return;
                        }

                        if (Path.GetExtension(ExcelFilename.Substring(1)) == ".txt" || Path.GetExtension(ExcelFilename.Substring(1)) == ".TXT")
                            TextCSV_FromDataTable(t, ExcelFilename);
                        else
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
            //   string[] portlist = SerialPort.GetPortNames();
            //  foreach (String s in portlist)
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
                if (devId4 == null)
                {
                    MessageBox.Show("Flats selected but not connected to ASCOM switch", "scopefocus");
                    checkBox18.Checked = false;
                    return;
                }
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
        //    int TotalTime;

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

                //   CaptureTime = Convert.ToInt32(textBox37.Text) * 1000;
                //   subsperfilter = ((Convert.ToInt32(textBox41.Text)/CaptureTime));//calculate number of subs based on server total time
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
                ShowWindowAsync(Handles.NebhWnd, SW_SHOW);
                Thread.Sleep(200);
                SetForegroundWindow(Handles.Aborthwnd);
                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);

                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);
                //********added 6-6
                //   serverStream.Dispose();
                //   serverStream.Close();
                //  clientSocket2.GetStream().Dispose();
                //  clientSocket2.GetStream().Close();

                if (!UseClipBoard.Checked)
                {



                    clientSocket2.Close();
                    Thread.Sleep(250);
                    SendKeys.SendWait("~");
                    SendKeys.Flush();
                    //*********

                    //   NebListenOn = false;
                    if (clientSocket2.Connected == false)
                        Log("Slave Disconnected");
                    else
                        Log("slave still connected");
                }
            }

        }


        int ServerStatusHandle;
        //  int NebhWnd2;
        int Slavehwnd;
        //    int SocketPort2;
        //enable server
        private void checkBox20_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox20.Checked)
            {
                Handles H = new Handles();
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
                IntPtr SlavehwndPtr = Handles.SearchForWindow("WindowsForms10", "scopefocus - slave mode");
                Log("scopefocus-slave handle found --  " + SlavehwndPtr.ToInt32());
                Slavehwnd = SlavehwndPtr.ToInt32();
                //full class name   WindowsForms10.Window.8.app.0.201d787_r15_ad1
                if (Slavehwnd != 0)
                {
                    Callback myCallBack2 = new Callback(H.EnumChildGetValue);
                    EnumChildWindows(Slavehwnd, myCallBack2, 0);//this gets the slave capture, pause, flat and gotofocus buttons
                                                                //this sends the server status box handle to the slave textbox
                    string send = textBox41.Handle.ToString();
                    StringBuilder sb = new StringBuilder(send);
                    SendMessage(Handles.SlaveStatushwnd, WM_SETTEXT, 0, sb);
                }



                //make sure it's not the spawned copy  

                MainWindowName = "Nebulosity v3.0";//change the mainwindow label in neb to signify sencond copy
                IntPtr hWnd2 = Handles.SearchForWindow("wxWindow", MainWindowName);//must open file named test.fit to specifiy second copu
                Log("Ensure not slave handle -- " + hWnd2.ToInt32());
                Handles.NebhWnd = hWnd2.ToInt32();
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
                GlobalVariables.ServerEnabled = true;  // 12-8-16 untested
            }
            else
               GlobalVariables.ServerEnabled = false;   // 12-8-16 untested was a method in main window 
            
        }
        // public int Serverhwnd;
        //enable slave mode
        private void checkBox14_CheckedChanged_1(object sender, EventArgs e)
        {
           
            Handles H = new Handles();
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
                Callback myCallBack3 = new Callback(H.EnumChildGetValue);

                //   int hWnd;
                IntPtr hWnd2 = Handles.SearchForWindow("wxWindow", MainWindowName);//must open file named test.fit to specifiy second copu
                Log("Neb(spawned) handle -- " + hWnd2.ToInt32());
                Handles.NebhWnd = hWnd2.ToInt32();
                // FindHandles();
                if (Handles.NebhWnd == 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Nebulosity slave Not Found - Open or Close & Restart and hit 'Retry' or 'Ignore' to continue",
                       "scopefocus", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);//change so ok moves focus
                    if (result == DialogResult.Ignore)
                        this.Refresh();
                    if (result == DialogResult.Retry)
                    {
                        H.FindHandles();
                        this.Refresh();
                    }
                    if (result == DialogResult.Abort)
                        System.Environment.Exit(0);

                }
                else
                {

                    EnumChildWindows(Handles.NebhWnd, myCallBack3, 0);
                }
                //  FindNebCamera();
                //IntPtr ServerhwndPtr = Handles.SearchForWindow("WindowsForms10", "scopefocus - Main");
                //Log("scopefocus-server handle found --  " + ServerhwndPtr.ToInt32());
                //Serverhwnd = ServerhwndPtr.ToInt32();
                //  this.Refresh();
                GlobalVariables.SlaveModeEnabled = true; // 12-8-16
            }
            else
            {
                GlobalVariables.SlaveModeEnabled = false;
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
                string send = Handles.NebhWnd.ToString();
                StringBuilder sb = new StringBuilder(send);
                //  sb = send;
                SendMessage(ServerStatusHandle, WM_SETTEXT, 0, sb);
                StatusHandleSend = true;
            }
            if ((IsServer()) & (StatusHandleRec == false))
            {
                GlobalVariables.NebSlavehwnd = Convert.ToInt32(textBox41.Text);
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

            SetForegroundWindow(Handles.Pausehwnd);
            PostMessage(Handles.Pausehwnd, BN_CLICKED, 0, 0);
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
            SendMessage(Handles.SlaveStatushwnd, WM_SETTEXT, 0, sb);
            Thread.Sleep(500);
        }
        private void SendtoServer(string message)
        {
            Handles H = new Handles();
            StringBuilder sb = new StringBuilder(message);
            ShowWindowAsync(H.Serverhwnd(), SW_SHOW);
            Thread.Sleep(200);
            SetForegroundWindow(H.Serverhwnd());
            SendMessage(ServerStatusHandle, WM_SETTEXT, 0, sb);
        }
        private void SlaveCapture()
        {
            // SendtoSlave((subsperfilter * CaptureTime3).ToString());
            SlavePaused = false;
            ShowWindowAsync(Slavehwnd, SW_SHOW);
            ShowWindowAsync(Slavehwnd, SW_RESTORE);
            SetForegroundWindow(Slavehwnd);
            PostMessage(Handles.Capturehwnd, BN_CLICKED, 0, 0);
            // PostMessage(Capturehwnd, BN_CLICKED, 0, 0);
            Thread.Sleep(500);
        }

        private void button40_Click_1(object sender, EventArgs e)
        {
            SequenceGo();
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
            PostMessage(Handles.Flathwnd, BN_CLICKED, 0, 0);
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
                PostMessage(Handles.Flathwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);
            }


        }
        string[] captions;
        bool working = true;
        //   int Workerhwnd;
        //    string StatusLabel;

        private void SlaveFocus()
        {
            FilterFocusOn = true;//??unsure if needed was slaveflat on=true (
            ShowWindow(Handles.Gotofocushwnd, SW_SHOW);
            ShowWindow(Handles.Gotofocushwnd, SW_RESTORE);
            SetForegroundWindow(Slavehwnd);
            PostMessage(Handles.Gotofocushwnd, BN_CLICKED, 0, 0);
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
           
            GlobalVariables.Path2 = textBox11.Text.ToString();
            Log("Path set to " + GlobalVariables.Path2.ToString());
            fileSystemWatcher1.Path = GlobalVariables.Path2.ToString();
            fileSystemWatcher2.Path = GlobalVariables.Path2.ToString();
            fileSystemWatcher5.Path = GlobalVariables.Path2.ToString();//added to test metricHFR
            fileSystemWatcher3.Path = GlobalVariables.Path2.ToString();
            fileSystemWatcher4.Path = GlobalVariables.Path2.ToString();
        }


        private void textBox42_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This value is calcuated from server subs * capture time", "scopefocus");
        }

        private void checkBox22_CheckedChanged_1(object sender, EventArgs e)
        {
            EquipRefresh();
            metricpath = Directory.GetFiles(GlobalVariables.Path2.ToString(), "metric*.fit");
            if (checkBox22.Checked == true)
            {
                if (metricpath.Length > 0) // delete all prev metric files.  
                {
                    for (int x = 0; x < metricpath.Length; x++)
                    {
                        File.Delete(metricpath[x]);
                    }

                }
            }
            //this stuff not used w/ full frame metric
            //   checkBox10.Checked = false;  rem'd 8-26-13.  this should stay checked o/w will try to slew to target when not needed.
            //   checkBox10.Enabled = false;

            //******rem'd 5-22-14
            //   label37.Enabled = false;
            //  textBox29.Enabled = false;
            //   textBox30.Enabled = false;
            //    label23.Enabled = false;
            //    label28.Enabled = false;
        }

        private void button27_Click(object sender, EventArgs e)
        {
            SampleMetric();
        }
        /*
        //Start Addt to cal sideral time
        public static DateTime CalcGSTFromUT(DateTime dDate)
        {
           
            double fJD;
            double fS;
            double fT;
            double fT0;
            DateTime dGST;
            double fUT;
            double fGST;

            fJD = GetJulianDay(dDate.Date, 0);
            fS = fJD - 2451545.0;
            fT = fS / 36525.0;
            fT0 = 6.697374558 + (2400.051336 * fT) + (0.000025862 * fT * fT);
            fT0 = PutIn24Hour(fT0);
            fUT = ConvTimeToDec(dDate);
            fUT = fUT * 1.002737909;
            fGST = fUT + fT0;
            while (fGST > 24)
            {
                fGST = fGST - 24;
            }
            dGST = ConvDecTUraniaTime(fGST);
            return dGST;
        }

        public static DateTime ConvDecTUraniaTime(double fTime)
        {
            DateTime dDate;

            dDate = new DateTime();
            dDate = dDate.AddHours(fTime);
            return (dDate);
        }

        public static double ConvTimeToDec(DateTime dDate)
        {
            double fHour;

            fHour = dDate.Hour + (dDate.Minute / 60.0) + (dDate.Second / 60.0 / 60.0) + (dDate.Millisecond / 60.0 / 60.0 / 1000.0);
            return fHour;
        }
        public static double GetJulianDay(DateTime dDate, int iZone)
        {
            double fJD;
            double iYear;
            double iMonth;
            double iDay;
            double iHour;
            double iMinute;
            double iSecond;
            double iGreg;
            double fA;
            double fB;
            double fC;
            double fD;
            double fFrac;

            dDate = CalcUTFromZT(dDate, iZone);

            iYear = dDate.Year;
            iMonth = dDate.Month;
            iDay = dDate.Day;
            iHour = dDate.Hour;
            iMinute = dDate.Minute;
            iSecond = dDate.Second;
            fFrac = iDay + ((iHour + (iMinute / 60) + (iSecond / 60 / 60)) / 24);
            if (iYear < 1582)
            {
                iGreg = 0;
            }
            else
            {
                iGreg = 1;
            }
            if ((iMonth == 1) || (iMonth == 2))
            {
                iYear = iYear - 1;
                iMonth = iMonth + 12;
            }

            fA = (long)Math.Floor(iYear / 100);
            fB = (2 - fA + (long)Math.Floor(fA / 4)) * iGreg;
            if (iYear < 0)
            {
                fC = (int)Math.Floor((365.25 * iYear) - 0.75);
            }
            else
            {
                fC = (int)Math.Floor(365.25 * iYear);
            }
            fD = (int)Math.Floor(30.6001 * (iMonth + 1));
            fJD = fB + fC + fD + 1720994.5;
            fJD = fJD + fFrac;
            return fJD;
        }

        public static DateTime getDateFromJD(double fJD, int iZone)
        {
            DateTime dDate;
            int iYear;
            int fMonth;
            int iDay;
            int iHour;
            int iMinute;
            int iSecond;
            double fFrac;
            double fFracDay;
            int fI;
            int fA;
            int fB;
            int fC;
            int fD;
            int fE;
            int fG;

            fJD = fJD + 0.5;
            fI = (int)Math.Floor(fJD);
            fFrac = fJD - fI;

            if (fI > 2299160)
            {
                fA = (int)Math.Floor((fI - 1867216.25) / 36524.25);
                fB = fI + 1 + fA - (int)Math.Floor((double)(fA / 4));
            }
            else
            {
                fA = 0;
                fB = fI;
            }
            fC = fB + 1524;
            fD = (int)Math.Floor((fC - 122.1) / 365.25);
            fE = (int)Math.Floor(365.25 * fD);
            fG = (int)Math.Floor((fC - fE) / 30.6001);
            fFracDay = fC - fE + fFrac - (long)Math.Floor((double)(30.6001 * fG));
            iDay = (int)Math.Floor(fFracDay);
            fFracDay = (fFracDay - iDay) * 24;
            iHour = (int)Math.Floor(fFracDay);
            fFracDay = (fFracDay - iHour) * 60;
            iMinute = (int)Math.Floor(fFracDay);
            fFracDay = (fFracDay - iMinute) * 60;
            iSecond = (int)Math.Floor(fFracDay);

            if (fG < 13.5)
            {
                fMonth = fG - 1;
            }
            else
            {
                fMonth = fG - 13;
            }

            if (fMonth > 2.5)
            {
                iYear = (int)Math.Floor((double)(fD - 4716.0));
            }
            else
            {
                iYear = (int)Math.Floor((double)(fD - 4715.0));
            }


            dDate = new DateTime(iYear, (int)Math.Floor((double)fMonth), iDay, iHour, iMinute, iSecond);
            dDate = CalcZTFromUT(dDate, iZone);
            return dDate;
        }
        public static DateTime CalcUTFromGST(DateTime dGSTDate, DateTime dCalDate)
        {
            double fJD;
            double fS;
            double fT;
            double fT0;
            double fGST;
            double fUT;
            DateTime dUT;

            bool bPrevDay;
            bool bNextDay;

            bPrevDay = false;
            bNextDay = false;

            fJD = GetJulianDay(dCalDate.Date, 0);
            fS = fJD - 2451545.0;
            fT = fS / 36525.0;
            fT0 = 6.697374558 + (2400.051336 * fT) + (0.000025862 * fT * fT);

            fT0 = PutIn24Hour(fT0);

            fGST = ConvTimeToDec(dGSTDate);
            fGST = fGST - fT0;

            while (fGST > 24)
            {
                fGST = fGST - 24;
                bNextDay = true;
            }
            while (fGST < 0)
            {
                fGST = fGST + 24;
                bPrevDay = true;
            }

            fUT = fGST * 0.9972695663;
            // fUT = fGST

            dUT = dCalDate.Date;
            dUT = dUT.AddHours(fUT);
            fUT = dUT.Millisecond;

            if (bNextDay == true)
            {
                dUT = dUT.AddDays(1);
            }

            if (bPrevDay == true)
            {
                dUT = dUT.Subtract(new TimeSpan(1, 0, 0, 0, 0));
            }
            return dUT;
        }
        public static DateTime CalcUTFromZT(DateTime dDate, int iZone)
        {
            if (iZone >= 0)
            {
                return dDate.Subtract(new TimeSpan(iZone, 0, 0));
            }
            else
            {
                return dDate.AddHours(Math.Abs(iZone));
            }
        }

        public static DateTime CalcZTFromUT(DateTime dDate, int iZone)
        {
            if (iZone >= 0)
            {
                return dDate.AddHours(iZone);
            }
            else
            {
                return dDate.Subtract(new TimeSpan(Math.Abs(iZone), 0, 0));
            }
        }
        public static DateTime CalcLMTFromUT(DateTime dDate, double fLong)
        {
            bool bAdd = false;
            DateTime dLongDate;

            dLongDate = ConvLongTUraniaTime(fLong, ref bAdd);

            if (bAdd == true)
            {
                dDate = dDate.Add(dLongDate.TimeOfDay);
            }
            else
            {
                dDate = dDate.Subtract(dLongDate.TimeOfDay);
            }

            return dDate;
        }

        public static DateTime ConvLongTUraniaTime(double fLong, ref bool bAdd)
        {
            //double fHours;
            double fMinutes;
            //double fSeconds;
            DateTime dDate;
            //DateTime dTmpDate;

            fMinutes = fLong * 4;
            if (fMinutes < 0)
            {
                bAdd = false;
            }
            else
            {
                bAdd = true;
            }
            fMinutes = Math.Abs(fMinutes);

            dDate = new DateTime();
            dDate = dDate.AddMinutes(fMinutes);
            return dDate;
        }
        public static double PutIn24Hour(double pfHour)
        {
            while (pfHour >= 24)
            {
                pfHour = pfHour - 24;
            }
            while (pfHour < 0)
            {
                pfHour = pfHour + 24;
            }
            return pfHour;
        }
        
        TimeSpan diff;
        int TimeCompare;
        public static DateTime GST;
        double TimeDec;
        public static DateTime Urania;
        string CurrentLoc;
        double Longitude;//use - for west
        int zone;
        double correction;
         */
        //  DateTime FlipTime;
        bool flipCheckDone = false;
        private void CheckForFlip()
        {
            //Must use dithering to perform meridian flip without loosing a sub
            //pausing PHD while slewing will allow neb pause.  once slew is done, phd resumes, then 
            //when settled, Neb resumes
            if (checkBox23.Checked == true)
            {
                toolStripStatusLabel1.Text = "Flip Check- Done";
                // TimeToFlip = Math.Round(Math.Abs(scope.SiderealTime - scope.RightAscension), 2) - .5; //remd 5-5-14
                //      textBox57.Text = TimeToFlip.ToString();
                //if (TimeToFlip < 0 && FlipDone == false)//remd 5-5-14 plus next 3
                //    FlipNeeded = true;
                //else
                //    FlipNeeded = false;

                if (scope.SideOfPier == scope.DestinationSideOfPier(scope.RightAscension, scope.Declination))
                    FlipNeeded = false;
                else
                    FlipNeeded = true;

                //   Log("Flip Needed = " + FlipNeeded.ToString() + "    Time to flip " + TimeSpan.FromHours((double)TimeToFlip).ToString());
                if (FlipNeeded == true && FlipDone == false)
                {
                    // 12-7-16 
                    ph.StopCapture();
                   // PHDcommand(PHD_PAUSE);

                    //     Log("guiding paused for flip");
                    GotoTargetLocation();
                    while (scope.Slewing)
                        Thread.Sleep(100);
                    FlipDone = true;
                    FlipNeeded = false;
                    if (Handles.PHDVNumber == 2)
                        resumePHD2();
                    else
                        ph.Guide();
                    // 12-7-16
                      //  PHDcommand(PHD_RESUME);
                    //  Log("flip done");
                }
            }

            flipCheckDone = true;// allows for only checking once per dither. 

        }

        //if t1 is less than t2 then result is Less than zero


        //if t1 equals t2 then result is Zero

        //if t1 is greater than t2 then result isGreater zero 

        bool FlipDone = false;
        bool okToFlip = false;

        private void checkBox23_CheckedChanged(object sender, EventArgs e)
        {
            //  MessageBox.Show("Meridian Flip can only happen during dither pause, filter change or periodic re-focus");
            // button48.PerformClick();
            CheckForFlip();
            if (TimeToFlip < 0)
            {
                DialogResult result1 = MessageBox.Show("Target already past Meridian", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (result1 == DialogResult.OK)
                {
                    checkBox23.Checked = false;
                    return;
                }
            }
            else
            {
                okToFlip = true;
                DialogResult result1 = MessageBox.Show("Merdian Flip in " + (TimeSpan.FromHours((double)TimeToFlip).ToString()), "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (result1 == DialogResult.OK)
                {
                    return;
                }
            }
            //prob fileSytemwatcher 4 that can monitor times for meridian flip


        }

        //****************future*************************
        // work on "center here" function.  figure out how to translate neb cursor position to mount movement. 
        //use conversion factor much like dx/dy.  need to determine movement vectors then send as mount slew
        // may be able to use the solved tables solve.xyls and solve.rdls, extract values of the xy coords, and RA/DEC for 
        //a few stars using ./tablist for each of those files.  then use the astrometric position calculator
        // www.gyes.eu/calculator/calculator_page2.htm
        //get center info from solve.wcs using wcsinfo
        //will need to convert from RA Dec in decimal to hh:mm:ss.ss and +-/dd:mm:ss.ss for the calculator
        //calculator on web seems to miss the neg Dec when calculating



        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e) // update mount info every second
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            if (worker.CancellationPending == true)
            {
                e.Cancel = true;
            }
            else
                //  NebCapture();
                while (Capturing == true)
                {
                    MonitorNebStatus();
                    msdelay(750);

                }
        }

        bool usingASCOM = false;
        private void Chooser()
        {
            ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
            chooser.DeviceType = "Telescope";
            devId = chooser.Choose();
            if (devId != "")
                scope = new ASCOM.DriverAccess.Telescope(devId);
            else
                return;
            //  ASCOM.DriverAccess.Telescope scope = new ASCOM.DriverAccess.Telescope(devId);
            Log("connected to " + devId);
            FileLog2("connected to " + devId);
            scope.Connected = true;
            if (scope.Connected)
            {
                timer2.Enabled = true;
                timer2.Start();
            }
            usingASCOM = true;
            button49.BackColor = System.Drawing.Color.Lime;
            //   button49.Text = "Connected";
        }
      //  bool usingASCOMFocus = false;
        private static string devId2;

        public void FocusChooser()
        {

            ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
            chooser.DeviceType = "Focuser";
            devId2 = chooser.Choose();
            //  ASCOM.DriverAccess.Focuser focuser = new ASCOM.DriverAccess.Focuser(devId2);
            if (devId2 != "")
                focuser = new Focuser(devId2);
            else
                return;



            focuser.Connected = true;
            //***************** I think this needs to be changes so it SETs the value not GETs***************
            //go back to previous method for storing the maxtravel in settings, then when gets it 
            //after selecting equipement it sets the value. 
            //****************************************************************************************
            travel = focuser.MaxStep;
            textBox2.Text = travel.ToString();
            count = focuser.Position;
            textBox1.Text = count.ToString();
            Log("connected to " + devId2);
            FileLog2("connected to " + devId2);
            button8.BackColor = System.Drawing.Color.Lime;
            //  button8.Text = "Connected";
            /*
            if (focuser.Connected)
            {
                MessageBox.Show("ASCOM Focuser connected");
            }
             */
            numericUpDown6.Value = focuser.Position;
        //    usingASCOMFocus = true;

            //   focuser.CommandString("C", true);


        }
        // ASCOM.DriverAccess.Telescope scope;
        public static string devId;
        private void button49_Click(object sender, EventArgs e)
        {

            if (button49.BackColor != Color.Lime)
                Chooser();
            else
            {
                scope.Connected = false;
                button49.BackColor = Color.WhiteSmoke;
                scope.Dispose();
                Log(devId + " disconnected");
                devId = "";
            }
        }
        // added 3-12-16 for online solve  
        private static double _rA;
        public static double RA
        {
            get { return _rA; }
            set { _rA = value; }
        }
        private static double _dEC;
        public static double DEC
        {
            get { return _dEC; }
            set { _dEC = value; }
        }











        int IndexOfSecond(string theString, string toFind)
        {
            int first = theString.IndexOf(toFind);

            if (first == -1) return -1;

            // Find the "next" occurrence by starting just past the first
            return theString.IndexOf(toFind, first + 1);
        }

        //tablist test view  use w a button put fit file in for string
        private void tablistViewer(string file)
        {
            string commandXY = "./tablist " + file;
            tabList(commandXY);

        }






        //private void parseXY()
        //{
        //    /*
        //    var destDir = @"c:\cygwin\lib\astrometry\bin";
        //    // var pattern = "*.csv";
        //    // var file = solveImage;
        //    //   var sourceDir = @"c:\cygwin\home\astro";
        //    var destfile = "solve.fit";
        //    if (solveImage != Path.Combine(destDir, Path.GetFileName(solveImage)))
        //    {
        //        foreach (var files in new DirectoryInfo(destDir).GetFiles("*.*"))
        //        {
        //            files.Delete();//empty the directory
        //        }
        //        // File.Copy(solveImage, Path.Combine(destDir, Path.GetFileName(solveImage)));
        //        File.Copy(solveImage, Path.Combine(destDir, destfile));//copy to cygwin
        //    }
        //     */

        //    try
        //    {
        //        //to view a fits file, change above to parseXY(string file) then add file n
        //        //    string commandXY = "./tablist " + "solve.corr";  //just to see what is in the solve.corr file
        //        string commandXY = "./tablist " + "solve-indx.xyls";
        //        tabList(commandXY);
        //        StreamReader reader = new StreamReader(@"c:\cygwin\home\astro\text2.txt"); //read the cygwin log file
        //                                                                                   //  reader = FileInfo.OpenText("filename.txt");
        //        string line;
        //        // while ((line = reader.ReadToEnd()) != null) {
        //        line = reader.ReadToEnd();
        //        string[] items = line.Split('\n');

        //        string[,] FieldLine = new String[4, 2];
        //        string find = "   1 ";
        //        int pos = 0;
        //        bool found = false;


        //        int count = 0;
        //        FileLog2("    X     " + "     Y     ");
        //        foreach (string item in items)
        //        {

        //            if (item.Contains(find) && count < 10)
        //            {
        //                XYfound = true;
        //            }

        //            if (XYfound && count < 10)
        //            {

        //                //get and item for each x and y (then RA and DEc...later)
        //                // 4 x 2 array with x,y of first 4 stars 
        //                starsXY[count, 0] = Convert.ToDouble(item.Substring(12, 15));
        //                starsXY[count, 1] = Convert.ToDouble(item.Substring(36, 15));
        //                // int nextSpace = IndexOfSecond(item, "        ");
        //                //  FieldLine[count, 1] = item.Substring(36, 15);
        //                FileLog2(starsXY[count, 0] + "    " + starsXY[count, 1]);
        //                count++;

        //            }

        //        }
        //        if (!XYfound)
        //        {
        //            Log("ParseXY Error - Aborted");
        //            return;
        //        }
        //        reader.Close();
        //    }
        //    //  Thread.Sleep(2000);
        //    //   string command2 = "./tablist " + "solve.xyls";

        //    //do the same for first 4 stars RA and Dec

        //    //then find curson position of desired center here position

        //    //then enter into calculation

        //    //slew, recheck, repeat until...
        //    catch (Exception e)
        //    {
        //        Log("failed to parse X,Y data");
        //        FileLog("failed to parse X,Y data:  " + e.ToString());
        //    }




        //}

        //private void ParseRD()//*****combine w/ parse XY  ***** only diff is command
        //{
        //    try
        //    {
        //        string find = "   1 ";
        //        int count;

        //        string commandRD = "./tablist " + "solve.rdls";
        //        tabList(commandRD);
        //        //not reading a new text2 file w/ RD data...its using the prev XY one....

        //        StreamReader reader2 = new StreamReader(@"c:\cygwin\home\astro\text2.txt"); //read the cygwin log file
        //        //  reader = FileInfo.OpenText("filename.txt");
        //        string line2;
        //        // while ((line = reader.ReadToEnd()) != null) {
        //        line2 = reader2.ReadToEnd();
        //        string[] items2 = line2.Split('\n');

        //        //   string[,] FieldLine2 = new String[4, 2];
        //        // string find = "   1 ";
        //        RDfound = false;
        //        count = 0;
        //        FileLog2("     RA     " + "    Dec     ");
        //        foreach (string item2 in items2)
        //        {

        //            if (item2.Contains(find) && count < 10)
        //            {
        //                RDfound = true;
        //            }

        //            if (RDfound && count < 10)
        //            {

        //                //get and item for each x and y (then RA and DEc...later)
        //                // 4 x 2 array with x,y of first 4 stars 
        //                starsRD[count, 0] = Convert.ToDouble(item2.Substring(12, 15)) / 15;
        //                starsRD[count, 1] = Convert.ToDouble(item2.Substring(34, 15));
        //                // int nextSpace = IndexOfSecond(item, "        ");
        //                //FieldLine2[count, 1] = item2.Substring(34, 15);
        //                FileLog2(starsRD[count, 0] + "    " + starsRD[count, 1]);
        //                count++;

        //            }
        //        }
        //        if (!RDfound)
        //        {
        //            Log("ParseRD Error - Aborted");
        //            return;
        //        }

        //    }
        //    catch (Exception e)
        //    {
        //        Log("failed to parse RA/Dec data");
        //        FileLog("failed to parse RA/Dec data:  " + e.ToString());
        //    }
        //}

        // *****write parse for online that does both from path2/text3.txt

        public double[,] starsXY = new double[10, 2];
        bool XYfound = false;
        bool RDfound = false;
        public double[,] starsRD = new double[10, 2];
        bool parseError = false;
        public void ParseXYRD()
        {
            try
            {


                //  string dataFile = "";
                //  FitsCorrReader();  // done in astrometry Class
                if (GlobalVariables.LocalPlateSolve)
                {
                    AstrometryNet ast = new AstrometryNet(); // rem'd for onlyinstance
                                                             // ast = AstrometryNet.OnlyInstance;
                                                             //   AstrometryNet.GetForm.Show();
                    ast.MakeDataFile();
                    Thread.Sleep(200);
                    ast.Close();
                }

                StreamReader reader2 = new StreamReader(GlobalVariables.Path2 + "\\corr2text.txt"); //read text3 which is corr.fits converted. X, Y, RA, Dec
                                                                                                    // StreamReader reader2 = new StreamReader(dataFile);
                string content;
                // while ((line = reader.ReadToEnd()) != null) {
                content = reader2.ReadToEnd();
                string[] lines = content.Split('\n');

                //  starsXY[,] xy = new string[lines.Length, lines.Length];
                // string[] y = new string[lines.Length];
                // string[] ra = new string[lines.Length];
                // string[] dec = new string[lines.Length];
                FileLog2("        X      " + "     Y     " + "      RA     " + "       DEC");
                // string find = "   1 ";
                //   int count = 0;
                int k = 0;
                foreach (var line in lines)
                {

                    //if (GlobalVariables.LocalPlateSolve)
                    //{
                    //    if (line.Contains(find) && count < 10) // skip until "    1" if found , count stays 0
                    //    {
                    //        RDfound = true;
                    //    }
                    //}
                    //else
                    //    RDfound = true;

                    //    if (RDfound && count < 10)
                    //    {
                    if (k < 10)
                    {
                        var values = line.Split('\t');

                        // X = values [0] = 2407.794   
                        // Y =         [1] = 1167.264
                        // RA=         [2] = 68.48058
                        // DEC=         [3[ = 2407.93/r  
                        if (IsDouble(values[0]))
                            starsXY[k, 0] = Convert.ToDouble(values[0]);
                        else
                            parseError = true;
                        if (IsDouble(values[1]))
                            starsXY[k, 1] = Convert.ToDouble(values[1]);
                        else
                            parseError = true;
                        if (IsDouble(values[2]))
                            starsRD[k, 0] = Convert.ToDouble(values[2]);
                        else
                            parseError = true;
                        if (IsDouble(values[3]))
                            starsRD[k, 1] = Convert.ToDouble(values[3]);
                        else
                            parseError = true;

                        FileLog2("    " + starsXY[k, 0] + "    " + starsXY[k, 1] + "     " + starsRD[k, 0] + "    " + starsRD[k, 1]);
                        // FileLog2(starsRD[count, 0] + "    " + starsRD[count, 1]);
                        k++;
                        if (parseError)
                        {
                            Log("Erro parsing X/Y and RA/DEC data");
                            FileLog("failed to parse X/Y and RA/Dec data");
                            return;
                        }

                    }
                    else
                    {
                        XYfound = true;
                        RDfound = true;
                    }


                }
            }
            catch (Exception e)
            {
                Log("failed to parse X/Y and RA/Dec data");
                FileLog("failed to parse X?Y and RA/Dec data:  " + e.ToString());
            }
        }


        private bool IsDouble(string test)
        {
            double d;
            return double.TryParse(test, out d);
        }

        //private void FitsCorrReader()
        //{
        //    Fits f = new Fits(GlobalVariables.Path2 +  "\\corr.fits");
        //    BinaryTableHDU h = (BinaryTableHDU)f.GetHDU(1);

        //    //   Object[] row23 = h.GetRow(23);
        //    Object col_x = h.GetColumn(0);
        //    Object col_y = h.GetColumn(1);
        //    Object col_ra = h.GetColumn(3);
        //    Object col_dec = h.GetColumn(4);
        //    float[] x = new float[h.NRows];
        //    float[] y = new float[h.NRows];
        //    float[] ra = new float[h.NRows];
        //    float[] dec = new float[h.NRows];
        //    int i = 0;
        //    foreach (float item in (dynamic)col_x)
        //    {
        //        x[i] = (dynamic)(item);
        //        i++;
        //    }
        //    i = 0;
        //    foreach (float item2 in (dynamic)col_y)
        //    {
        //        y[i] = (dynamic)(item2);
        //        i++;
        //    }
        //    i = 0;
        //    foreach (float item in (dynamic)col_ra)
        //    {
        //        ra[i] = (dynamic)(item);
        //        i++;
        //    }
        //    i = 0;
        //    foreach (float item2 in (dynamic)col_dec)
        //    {
        //        dec[i] = (dynamic)(item2);
        //        i++;
        //    }
        //    using (StreamWriter sr = new StreamWriter(GlobalVariables.Path2 + "\\corr2text.txt"))
        //    {
        //        // sr.WriteLine("   x   " + "         y   ");
        //        for (i = 0; i < x.Length; i++)
        //        {
        //            sr.WriteLine(x[i] + "    \t" + y[i] + "    \t" + ra[i] + "    \t" + dec[i]);

        //        }
        //        sr.Dispose();
        //    }
        //}


        public double centerHereRA;
        public double centerHereDec;
        private void cramersRule()
        {
            try
            {
                double centerDec = DEC; //arbitrary dec postion for now
                double centerRA = RA;  //arbitrary. 
                if (DEC == 0 || RA == 0)
                {
                    MessageBox.Show("must slew to plate solve location first");
                    return;
                }
                var FL = 1;//arbitrary
                var dr = Math.PI / 180;
                double a0 = centerRA * dr;  // was * 15 as well
                double a0deg = centerRA;// was * 15
                double d0 = centerDec * dr;
                double d0deg = centerDec;
                var n = 10; //number of stars
                var sd = Math.Sin(d0);
                var cd = Math.Cos(d0);
                double[] ra = new double[10];
                double[] rd = new double[10];
                double r1 = 0;
                double r2 = 0;
                double r3 = 0;
                double r7 = 0;
                double r8 = 0;
                double r9 = 0;
                double xs = 0;
                double ys = 0;
                double[,] r = new double[10, 9];
                double[] starsY1 = new double[10];
                double[] starsX1 = new double[10];
                for (int i = 0; i < n; i++)
                {
                    starsRD[i, 0] = starsRD[i, 0] * dr;// was * 15 as well
                    starsRD[i, 1] = starsRD[i, 1] * dr;
                }

                for (var j = 0; j < n; j++)
                {

                    var sj = Math.Sin(starsRD[j, 1]);
                    var cj = Math.Cos(starsRD[j, 1]);
                    var hh = sj * sd + cj * cd * Math.Cos(starsRD[j, 0] - a0);
                    starsX1[j] = cj * Math.Sin(starsRD[j, 0] - a0) / hh;
                    starsY1[j] = (sj * cd - cj * sd * Math.Cos(starsRD[j, 0] - a0)) / hh;
                }
                for (var j = 0; j < n; j++)
                {
                    xs = (xs) + (starsXY[j, 0]);
                    ys = (ys) + (starsXY[j, 1]);
                    r[j, 0] = starsXY[j, 0] * starsXY[j, 0];
                    r1 = (r1) + (r[j, 0]);
                    r[j, 1] = starsXY[j, 1] * starsXY[j, 1];
                    r2 = (r2) + (r[j, 1]);
                    r[j, 2] = starsXY[j, 0] * starsXY[j, 1];
                    r3 = (r3) + (r[j, 2]);
                    r[j, 6] = starsY1[j] - starsXY[j, 1] / FL;
                    r7 = (r7) + (r[j, 6]);
                    r[j, 7] = r[j, 6] * starsXY[j, 0];
                    r8 = (r8) + (r[j, 7]);
                    r[j, 8] = r[j, 6] * starsXY[j, 1];
                    r9 = (r9) + (r[j, 8]);
                }
                // Now solve for plate constants d, e, f, by Cramer's Rule
                var dd = r1 * (r2 * n - ys * ys) - r3 * (r3 * n - xs * ys) + xs * (r3 * ys - xs * r2);
                var ddd = r8 * (r2 * n - ys * ys) - r3 * (r9 * n - r7 * ys) + xs * (r9 * ys - r7 * r2);
                var eee = r1 * (r9 * n - r7 * ys) - r8 * (r3 * n - xs * ys) + xs * (r3 * r7 - xs * r9);
                var fff = r1 * (r2 * r7 - ys * r9) - r3 * (r3 * r7 - xs * r9) + r8 * (r3 * ys - xs * r2);
                ddd = ddd / dd;
                eee = eee / dd;
                fff = fff / dd;
                double r4 = 0;
                double r5 = 0;
                double r6 = 0;
                for (int j = 0; j < n; j++)
                {
                    r[j, 3] = starsX1[j] - starsXY[j, 0] / FL;
                    r4 = (r4) + (r[j, 3]);
                    r[j, 4] = r[j, 3] * starsXY[j, 0];
                    r5 = (r5) + (r[j, 4]);
                    r[j, 5] = r[j, 3] * starsXY[j, 1];
                    r6 = (r6) + (r[j, 5]);
                }
                // Now solve for plate constants a, b, c, by Cramer's Rule
                var aaa = r5 * (r2 * n - ys * ys) - r3 * (r6 * n - r4 * ys) + xs * (r6 * ys - r4 * r2);
                var bbb = r1 * (r6 * n - r4 * ys) - r5 * (r3 * n - xs * ys) + xs * (r3 * r4 - xs * r6);
                var ccc = r1 * (r2 * r4 - ys * r6) - r3 * (r3 * r4 - xs * r6) + r5 * (r3 * ys - xs * r2);
                aaa = aaa / dd;
                bbb = bbb / dd;
                ccc = ccc / dd;
                // Now find the rms residuals
                double s = 0;//orig is as
                double ds = 0;
                double res2 = 0;
                for (var j = 0; j < n; j++)
                {
                    ra[j] = starsXY[j, 0] - FL * (starsX1[j] - (aaa * starsXY[j, 0] + bbb * starsXY[j, 1] + ccc));
                    rd[j] = starsXY[j, 1] - FL * (starsY1[j] - (ddd * starsXY[j, 0] + eee * starsXY[j, 1] + fff));
                    s = (s) + Math.Pow(ra[j], 2);
                    ds = (ds) + Math.Pow(rd[j], 2);
                    res2 = (res2) + Math.Pow(ra[j], 2) + Math.Pow(rd[j], 2);
                }
                double xt;
                double yt;


                xt = Convert.ToDouble(textBox31.Text);//arbitrary for now //input x pixel for target 
                yt = Convert.ToDouble(textBox32.Text);

                var sigma1 = Math.Sqrt(s / (n - 3));
                var sigma2 = Math.Sqrt(ds / (n - 3));
                var sigma3 = (sigma1 / FL) * 3600 / (dr * 15 * Math.Cos(d0));
                var sigma4 = (sigma2 / FL) * 3600 / dr;
                var sigma5 = Math.Sqrt(res2 / (n - 3));
                var sigma6 = (sigma5 / FL) * 3600 / dr;
                Log("RMS residual RA = " + sigma3.ToString() + " arcsec");
                FileLog2("RMS residual RA = " + sigma3.ToString() + " arcsec");
                Log("RMS residual Dec = " + sigma4.ToString() + " arcsec");
                FileLog2("RMS residual Dec = " + sigma4.ToString() + " arcsec");
                Log("RMS residual = " + sigma6.ToString() + " arcsec");
                FileLog2("RMS residual = " + sigma6.ToString() + " arcsec");
                // form.sigma3.value = sigma3;//****not sure what these are...
                //  form.sigma4.value = sigma4;
                //  form.sigma6.value = sigma6;
                // Compute the standard coordinates of target
                //  result = xy(form.xyt.value);
                //   var xt = result.x;  var yt = result.y;
                // var a0 = nRA * 15 * dr;
                //  var a0 = centerRA * 15 * dr;
                var xx = aaa * xt + bbb * yt + ccc + xt / FL;
                var yy = ddd * xt + eee * yt + fff + yt / FL;
                bbb = cd - yy * sd;
                var g = Math.Sqrt((xx * xx) + (bbb * bbb));
                // Find right ascension of target
                var a5 = Math.Atan(xx / bbb);
                if (bbb < 0) a5 = a5 + Math.PI;
                var a6 = (a5) + (a0);
                if (a6 > 2 * Math.PI) a6 = a6 - 2 * Math.PI;
                if (a6 < 0) a6 = (a6) + 2 * Math.PI;
                centerHereRA = (a6 / (dr * 15)) * 15;//was * 15

                Log("Center here RA = " + centerHereRA.ToString());
                // result = sexagesimal(v);  //use this to convert to hh:mm:ss
                //  var ia1=result.hours;
                //  var ia2=result.minute


                //  var a3=result.seconds;
                // Find declination of target
                double d6 = Math.Atan((sd + yy * cd) / g);
                centerHereDec = d6 / dr;

                Log("Center here Dec = " + centerHereDec.ToString());
                FileLog2("Calculated Offset:  Dec = " + centerHereDec.ToString() + "     RA = " + centerHereRA.ToString());
                // result = sexagesimal(v);
                //  var id1=result.hours;
                //   var id2=result.minutes;
                //  var d3=result.seconds;
                //  form.radect.value = ia1+" "+ia2+" "+a3+" "+id1+" "+id2+" "+d3;
            }
            catch (Exception e)
            {
                Log("Calculation failed");
                FileLog("Calculation Failed :  " + e.ToString());
            }



        }

        //partialsolvehere
        bool atfocus = false;
        bool attarget = false;
        private void PartialSolve()
        {
            try
            {

                toolStripStatusLabel1.Text = "Plate Solving";
                toolStripStatusLabel1.BackColor = Color.Lime;
                if (GlobalVariables.LocalPlateSolve)
                {
                    //scope = new ASCOM.DriverAccess.Telescope(devId);

                    if (usingASCOM == false)
                    {
                        MessageBox.Show("must be connected to ASCOM telescope to use, aborting");
                        return;
                    }


                    CurrentRA = scope.RightAscension;
                    CurrentDEC = scope.Declination;
                    //    var destDir = @"c:\cygwin\home\astro";
                    //    var destDir = @"C:\Users\ksipp_000\AppData\Local\cygwin_ansvr\bin"; // 10-22-16
                    //   var destDir = @"%localappdata%\cygwin_ansvr\bin"; // 10-22-16

                    var destDir = (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\cygwin_ansvr\bin"); // 10-22-16



                    // var pattern = "*.csv";
                    // var file = solveImage;
                    //   var sourceDir = @"c:\cygwin\home\astro";



                    // begin\r rem 4-26-16

                    //string extension;
                    //var destfile = "";
                    //if (Path.GetExtension(GlobalVariables.SolveImage) == "fit")
                    //    destfile = "solve.fit";
                    //else
                    //    destfile = "solve.jpg";

                    // end rem 


                    // begin addition
                    var destfile = Path.GetFileName(GlobalVariables.SolveImage);


                    // rem below 10-22-16  // changed solve-filed option -O to allow overwrite, no need to delet anything.  

                    //if (GlobalVariables.SolveImage != Path.Combine(destDir, Path.GetFileName(GlobalVariables.SolveImage)))
                    //{
                    //    foreach (var files in new DirectoryInfo(destDir).GetFiles("*.*"))
                    //    {
                    //        if (files.Name != "tablist.exe")
                    //            files.Delete();//empty the directory
                    //    }
                    //    // File.Copy(solveImage, Path.Combine(destDir, Path.GetFileName(solveImage)));
                    //    File.Copy(GlobalVariables.SolveImage, Path.Combine(destDir, destfile));//copy to cygwin  this one works!


                    //}

                    ExecuteCommand();



                    // StreamReader reader = new StreamReader(@"c:\cygwin\text.txt"); //read the cygwin log file
                    //  reader = FileInfo.OpenText("filename.txt");
                    //    StreamReader reader = new StreamReader(@"C:\Users\ksipp_000\AppData\Local\cygwin_ansvr\text.txt"); // 10-22-16
                    StreamReader reader = new StreamReader((Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)) + @"\cygwin_ansvr\text.txt"); // 10-22-16


                    string line;
                    // while ((line = reader.ReadToEnd()) != null) {
                    line = reader.ReadToEnd();
                    string[] items = line.Split('\n');
                    string FieldLine = null;
                    foreach (string item in items)
                    {
                        //try parsing a different line in .txt file

                        if (item.EndsWith("pix.\r"))
                        {

                            //e.g.    RA,Dec = (205.003,49.9932), pixel scale 1.23661 arcsec/pix. 

                            FieldLine = item;
                            Log(FieldLine);
                            //     int FirstComma = item.IndexOf(","); 
                            int SecondComma = IndexOfSecond(item, ",");
                            int Start = item.IndexOf("(");
                            int End = item.IndexOf(")");
                            //   int EqualsIndex = FieldLine.IndexOf(@"=");
                            string ParsedRA = "";
                            string ParsedDEC = "";
                            int RAend = SecondComma - Start;
                            int DECend = End - SecondComma;
                            ParsedRA = FieldLine.Substring(Start + 1, RAend);
                            ParsedRA = Regex.Replace(ParsedRA, "[,]", "");

                            ParsedDEC = FieldLine.Substring(SecondComma + 1, DECend);
                            ParsedDEC = Regex.Replace(ParsedDEC, "[(),d]", "");

                            Log("Parsed DEC = " + ParsedDEC);
                            DEC = Convert.ToDouble(ParsedDEC);//coords from plate solve
                            RA = Convert.ToDouble(ParsedRA) / 15; //convert to hours
                            Log("Parsed RA = " + RA.ToString());

                        }
                    }
                    reader.Close();

                }
                if (SettingFocusSolve)
                {
                    // ****** need error handling for failed solve

                    FocusDEC = DEC;
                    FocusRA = RA;
                    Log("Solved Focus Location:  RA = " + FocusRA.ToString() + "   DEC = " + FocusDEC.ToString());
                    FileLog2("Solved Focus Location:  RA = " + FocusRA.ToString() + "   DEC = " + FocusDEC.ToString());
                    button32.Text = "Solved"; // change the set button only...since not necessarily there (could be from prev session) 
                    button32.BackColor = System.Drawing.Color.Yellow;
                    fileSystemWatcher7.EnableRaisingEvents = false;
                    //   SettingFocusSolve = false; // remd 11-9-16 and one below

                }
                if (SettingTargetSolve)
                {
                    // ****** need error handling for failed solve

                    TargetDEC = DEC;
                    TargetRA = RA;
                    Log("Solved Target Location:  RA = " + TargetRA.ToString() + "   DEC = " + TargetDEC.ToString());
                    FileLog2("Solved Target Location:  RA = " + TargetRA.ToString() + "   DEC = " + TargetDEC.ToString());
                    button34.Text = "Solved"; // change the set button only...since not necessarily there (could be from prev session) 
                    button34.BackColor = System.Drawing.Color.Yellow;
                    fileSystemWatcher7.EnableRaisingEvents = false;
                    //   SettingTargetSolve = false;
                }

                if (ConfirmingLocation)
                {
                    CurrentRA = scope.RightAscension;
                    CurrentDEC = scope.Declination;

                    // first attempts at comparing parse solve coords w/ scope coords.
                    //need to FIX...seems to maintain RA and DEC from first plate solve after the second one. 
                    //if (usingASCOM == true)
                    //{
                    //    //  scope = new ASCOM.DriverAccess.Telescope(devId);
                    //    scope.SlewToCoordinates(CurrentRA, CurrentDEC);//go back to where originally though it was supposed to be
                    //                                                   //should be closer after the sync
                    //}
                    //else
                    //    MessageBox.Show("Must use ASCOM mount connection", "scopefocus");

                    if ((Math.Abs(CurrentRA - RA) * 60 > Convert.ToDouble(textBox59.Text)) || (Math.Abs(CurrentDEC - DEC) * 60 > Convert.ToDouble(textBox59.Text)))//*************untested!!****
                    {
                        scope.SyncToCoordinates(RA, DEC);//sync to parsed(solve) location 
                        Log("repeating" + "DeltaRA = " + ((Math.Abs(CurrentRA - RA) * 60).ToString()) + "     DeltaDEC = " + ((Math.Abs(CurrentDEC - DEC) * 60).ToString()));
                        if (solvetry < 3)
                        {
                            solvetry++;
                            SolveCapture();
                        }
                        if (solvetry == 3)
                        {
                            Log("Solve Confirmation failed, using mount coordinates");
                            fileSystemWatcher7.EnableRaisingEvents = false;

                            solvetry = 0;
                            ConfirmingLocation = false;
                            fileSystemWatcher7.EnableRaisingEvents = false;
                            if (button33.Text == "Pending")
                            {
                                button33.Text = "At Focus";
                                button33.BackColor = System.Drawing.Color.Lime;
                                button35.Text = "Goto";
                                button35.UseVisualStyleBackColor = true;
                            }
                            if (button35.Text == "Pending")
                            {
                                button35.Text = "At Target";
                                button35.BackColor = System.Drawing.Color.Lime;
                                button33.Text = "Goto";
                                button33.UseVisualStyleBackColor = true;
                            }

                            //  *****   not closing neb socket properly ******   4-12-16

                            if (!UseClipBoard.Checked)
                            {
                                Thread.Sleep(500);
                                serverStream.Flush();
                                Thread.Sleep(1000);
                                //    serverStream.Close();
                                //     Thread.Sleep(1000);
                                SetForegroundWindow(Handles.NebhWnd);
                                Thread.Sleep(1000);
                                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                                Thread.Sleep(2000);
                                //      NebListenOn = false;
                                // clientSocket.GetStream().Close();//added 5-17-12
                                //  clientSocket.Client.Disconnect(true);//added 5-17-12
                                clientSocket.Close();
                                //   Log("at focus star");
                            }



                        }
                        //     button55.PerformClick();//prob dont need since fsw7 still on
                        //if (fileSystemWatcher7.EnableRaisingEvents == false)
                        //    fileSystemWatcher7.EnableRaisingEvents = true;
                        //SetForegroundWindow(Handles.NebhWnd);  //rem'd to testing 
                        //PostMessage(Handles.CaptureMainhWnd, BN_CLICKED, 0, 0);//rem'd for testing
                    }
                    else
                    {

                        Log("Confimred: At Plate Solve Location");
                        solvetry = 0;
                        ConfirmingLocation = false;
                        fileSystemWatcher7.EnableRaisingEvents = false;
                        if (button33.Text == "Pending")
                        {
                            button33.Text = "At Focus";
                            button33.BackColor = System.Drawing.Color.Lime;
                            button35.Text = "Goto";
                            button35.UseVisualStyleBackColor = true;
                        }
                        if (button35.Text == "Pending")
                        {
                            button35.Text = "At Target";
                            button35.BackColor = System.Drawing.Color.Lime;
                            button33.Text = "Goto";
                            button33.UseVisualStyleBackColor = true;
                        }

                        //  *****   not closing neb socket properly ******   4-12-16

                        if (!UseClipBoard.Checked) // 11-8-16
                        {
                            Thread.Sleep(500);
                            serverStream.Flush();
                            Thread.Sleep(1000);
                            //    serverStream.Close();
                            //     Thread.Sleep(1000);
                            SetForegroundWindow(Handles.NebhWnd);
                            Thread.Sleep(1000);
                            PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                            Thread.Sleep(2000);
                            //     NebListenOn = false;
                            // clientSocket.GetStream().Close();//added 5-17-12
                            //  clientSocket.Client.Disconnect(true);//added 5-17-12
                            clientSocket.Close();
                            //   Log("at focus star");
                        }
                    }


                }

                toolStripStatusLabel1.Text = "Ready";
                toolStripStatusLabel1.BackColor = Color.WhiteSmoke;
                this.Refresh();
            }
            catch (Exception e)
            {
                Log("Partial Solve failed");
                FileLog("Partial Solve Failed" + e.ToString());
            }

        }


        private void Solve()
        {
            try
            {
                PartialSolve();
                //toolStripStatusLabel1.Text = "Plate Solving";
                //toolStripStatusLabel1.BackColor = Color.Lime;
                //if (GlobalVariables.LocalPlateSolve)
                //{
                //    //scope = new ASCOM.DriverAccess.Telescope(devId);
                //    if (usingASCOM == false)
                //    {
                //        MessageBox.Show("must be connected to ASCOM telescope to use, aborting");
                //        return;
                //    }
                //    CurrentRA = scope.RightAscension;
                //    CurrentDEC = scope.Declination;
                //    var destDir = @"c:\cygwin\home\astro";
                //    // var pattern = "*.csv";
                //    // var file = solveImage;
                //    //   var sourceDir = @"c:\cygwin\home\astro";
                //    var destfile = "solve.fit";
                //    if (GlobalVariables.SolveImage != Path.Combine(destDir, Path.GetFileName(GlobalVariables.SolveImage)))
                //    {
                //        foreach (var files in new DirectoryInfo(destDir).GetFiles("*.*"))
                //        {
                //            if (files.Name != "tablist.exe")
                //                files.Delete();//empty the directory
                //        }
                //        // File.Copy(solveImage, Path.Combine(destDir, Path.GetFileName(solveImage)));
                //        File.Copy(GlobalVariables.SolveImage, Path.Combine(destDir, destfile));//copy to cygwin
                //    }
                //    ExecuteCommand();



                //    StreamReader reader = new StreamReader(@"c:\cygwin\text.txt"); //read the cygwin log file
                //                                                                   //  reader = FileInfo.OpenText("filename.txt");
                //    string line;
                //    // while ((line = reader.ReadToEnd()) != null) {
                //    line = reader.ReadToEnd();
                //    string[] items = line.Split('\n');
                //    string FieldLine = null;
                //    foreach (string item in items)
                //    {
                //        //try parsing a different line in .txt file

                //        if (item.EndsWith("pix.\r"))
                //        {

                //            //e.g.    RA,Dec = (205.003,49.9932), pixel scale 1.23661 arcsec/pix. 

                //            FieldLine = item;
                //            Log(FieldLine);
                //            //     int FirstComma = item.IndexOf(","); 
                //            int SecondComma = IndexOfSecond(item, ",");
                //            int Start = item.IndexOf("(");
                //            int End = item.IndexOf(")");
                //            //   int EqualsIndex = FieldLine.IndexOf(@"=");
                //            string ParsedRA = "";
                //            string ParsedDEC = "";
                //            int RAend = SecondComma - Start;
                //            int DECend = End - SecondComma;
                //            ParsedRA = FieldLine.Substring(Start + 1, RAend);
                //            ParsedRA = Regex.Replace(ParsedRA, "[,]", "");

                //            ParsedDEC = FieldLine.Substring(SecondComma + 1, DECend);
                //            ParsedDEC = Regex.Replace(ParsedDEC, "[(),d]", "");

                //            Log("Parsed DEC = " + ParsedDEC);
                //            DEC = Convert.ToDouble(ParsedDEC);//coords from plate solve
                //            RA = Convert.ToDouble(ParsedRA) / 15; //convert to hours
                //            Log("Parsed RA = " + RA.ToString());

                //        }
                //    }
                //}

                if ((!SettingFocusSolve) && (!SettingTargetSolve))
                {
                    string text1 = "Location: \n\r" + "RA: " + RA.ToString() + "\n\r" + "DEC: " + DEC.ToString();
                    DialogResult result = CustomMsgBox.Show(text1, "Plate Solve Success", "Slew/Set", "Snyc", "Ignore");
                    //   if ((result == DialogResult.Yes) & (checkBox25.Checked == false)) //slew
                    if (result == DialogResult.Yes)  //slew
                    {
                        // scope = new ASCOM.DriverAccess.Telescope(devId);
                        toolStripStatusLabel1.Text = "Slewing to Target";
                        toolStripStatusLabel1.BackColor = Color.Red;
                        this.Refresh();
                        scope.SlewToCoordinates(RA, DEC);
                        toolStripStatusLabel1.Text = "Ready";
                        toolStripStatusLabel1.BackColor = Color.WhiteSmoke;
                        Log("Slewed to RA = " + RA.ToString() + "   Dec = " + DEC.ToString());
                        FileLog2("Slewed to RA = " + RA.ToString() + "   Dec = " + DEC.ToString());




                        //remd 11-9-16
                        //if (SettingFocusSolve)
                        //{
                        //    FocusDEC = DEC;
                        //    FocusRA = RA;
                        //    button32.Text = "Solved";
                        //    button32.BackColor = System.Drawing.Color.Lime;
                        //    button33.Text = "At Focus";  // becuase with online solve the slew/set button was hit.  
                        //    button33.BackColor = Color.Lime;
                        //    button35.Text = "Goto";
                        //    button35.UseVisualStyleBackColor = true;
                        //    Log("Focus Location set to RA = " + RA.ToString() + "   Dec = " + DEC.ToString());
                        //    FileLog2("Focus Location set to RA = " + RA.ToString() + "   Dec = " + DEC.ToString());
                        //    SettingFocusSolve = false;
                        //}

                        //if (SettingTargetSolve)
                        //{
                        //    TargetDEC = DEC;
                        //    TargetRA = RA;
                        //    button34.Text = "Solved"; 
                        //    button34.BackColor = System.Drawing.Color.Lime;
                        //    button35.Text = "At Target";
                        //    button35.BackColor = Color.Lime;
                        //    button33.Text = "Goto";
                        //    button33.UseVisualStyleBackColor = true;

                        //    Log("Target Location set to RA = " + RA.ToString() + "   Dec = " + DEC.ToString());
                        //    FileLog2("Target Location set to RA = " + RA.ToString() + "   Dec = " + DEC.ToString());
                        //    SettingTargetSolve = false;
                        //}

                        if (checkBox25.Checked == true)//repeat until tolerance met
                        {
                            //group rem for new method below   4-12 16

                            //        //  Thread.Sleep(5000);
                            //        Log("calibrating");
                            //        //scope.SlewToCoordinates(RA, DEC);
                            //        CurrentRA = scope.RightAscension;
                            //        CurrentDEC = scope.Declination;
                            //        // first attempts at comparing parse solve coords w/ scope coords.
                            //        //need to FIX...seems to maintain RA and DEC from first plate solve after the second one. 
                            //        if (usingASCOM == true)
                            //        {
                            //            //  scope = new ASCOM.DriverAccess.Telescope(devId);
                            //            scope.SlewToCoordinates(CurrentRA, CurrentDEC);//go back to where originally though it was supposed to be
                            //                                                           //should be closer after the sync
                            //        }
                            //        else
                            //            MessageBox.Show("Must use ASCOM mount connection", "scopefocus");

                            //        if ((Math.Abs(CurrentRA - RA) * 60 > Convert.ToDouble(textBox59.Text)) || (Math.Abs(CurrentDEC - DEC) * 60 > Convert.ToDouble(textBox59.Text)))//*************untested!!****
                            //        {
                            //            scope.SyncToCoordinates(RA, DEC);//sync to parsed(solve) location 
                            //            Log("repeating" + "DeltaRA = " + ((Math.Abs(CurrentRA - RA) * 60).ToString()) + "     DeltaDEC = " + ((Math.Abs(CurrentDEC - DEC) * 60).ToString()));
                            ////     button55.PerformClick();//prob dont need since fsw7 still on
                            //          if (fileSystemWatcher7.EnableRaisingEvents == false)
                            //              fileSystemWatcher7.EnableRaisingEvents = true;
                            //            SetForegroundWindow(Handles.NebhWnd);  //rem'd to testing 
                            //            PostMessage(Handles.CaptureMainhWnd, BN_CLICKED, 0, 0);//rem'd for testing
                            //        }
                            //        else
                            //        {
                            //            scope.SyncToCoordinates(RA, DEC);
                            //            Log("sync tolerance met");
                            //            fileSystemWatcher7.EnableRaisingEvents = false;
                            //            Log("synced to:  RA = " + scope.RightAscension.ToString() + "     Dec = " + scope.Declination.ToString());
                            //        }

                            SolveToleranceCheck();

                        }

                    }




                    if (result == DialogResult.No) // sync
                    {
                        DialogResult result2 = MessageBox.Show("Are you sure you want to sync to this position?", "scopefocus", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                        if (result2 == DialogResult.OK)
                        {
                            scope.SyncToCoordinates(RA, DEC);
                            Log("synced to:  RA = " + scope.RightAscension.ToString() + "     Dec = " + scope.Declination.ToString());
                            fileSystemWatcher7.EnableRaisingEvents = false;
                        }
                        else
                            return;
                    }
                    if (result == DialogResult.Ignore) // ignore
                    {
                        fileSystemWatcher7.EnableRaisingEvents = false;
                        button55.BackColor = Color.WhiteSmoke;
                        toolStripStatusLabel1.Text = "Ready";
                        toolStripStatusLabel1.BackColor = Color.WhiteSmoke;
                        ast.Close();  // this may not need to be here
                        return;
                    }

                }

                button55.BackColor = Color.WhiteSmoke;
                toolStripStatusLabel1.Text = "Ready";
                toolStripStatusLabel1.BackColor = Color.WhiteSmoke;
                ast.Close();   // ??? remove       
                //else
                //{
                //    DEC = AstrometryNet.DecCenter;
                //    RA = AstrometryNet.RaCenter / 15;
                //}

                //**** added if statement 7-3-13 *******
                //if ((usingASCOM == true) & (checkBox25.Checked == false) & (checkBox24.Checked == false))//only slew if not calibrating and Not just syncing (CB24)
                //{
                //    // scope = new ASCOM.DriverAccess.Telescope(devId);
                //    toolStripStatusLabel1.Text = "Slewing to Target";
                //    this.Refresh();
                //    scope.SlewToCoordinates(RA, DEC);

                //    Log("Slewed to RA = " + RA.ToString() + "   Dec = " + DEC.ToString());
                //    FileLog2("Slewed to RA = " + RA.ToString() + "   Dec = " + DEC.ToString());
                //}


                //if (checkBox25.Checked == true)//repeat until tolerance met
                //{
                //    //  Thread.Sleep(5000);
                //    Log("calibrating");
                //    //scope.SlewToCoordinates(RA, DEC);
                //    CurrentRA = scope.RightAscension;
                //    CurrentDEC = scope.Declination;
                //    // first attempts at comparing parse solve coords w/ scope coords.
                //    //need to FIX...seems to maintain RA and DEC from first plate solve after the second one. 
                //    if (usingASCOM == true)
                //    {
                //        //  scope = new ASCOM.DriverAccess.Telescope(devId);
                //        scope.SlewToCoordinates(CurrentRA, CurrentDEC);//go back to where originally though it was supposed to be
                //                                                       //should be closer after the sync
                //    }
                //    else
                //        MessageBox.Show("Must use ASCOM mount connection", "scopefocus");

                //    if ((Math.Abs(CurrentRA - RA) * 60 > Convert.ToDouble(textBox59.Text)) || (Math.Abs(CurrentDEC - DEC) * 60 > Convert.ToDouble(textBox59.Text)))//*************untested!!****
                //    {
                //        scope.SyncToCoordinates(RA, DEC);//sync to parsed(solve) location 
                //        Log("repeating" + "DeltaRA = " + ((Math.Abs(CurrentRA - RA) * 60).ToString()) + "     DeltaDEC = " + ((Math.Abs(CurrentDEC - DEC) * 60).ToString()));
                //        //     button55.PerformClick();//prob dont need since fsw7 still on
                //        SetForegroundWindow(Handles.NebhWnd);  //rem'd to testing 
                //        PostMessage(Handles.CaptureMainhWnd, BN_CLICKED, 0, 0);//rem'd for testing
                //    }
                //    else
                //    {
                //        scope.SyncToCoordinates(RA, DEC);
                //        Log("sync tolerance met");
                //        fileSystemWatcher7.EnableRaisingEvents = false;
                //        Log("synced to:  RA = " + scope.RightAscension.ToString() + "     Dec = " + scope.Declination.ToString());
                //    }
                //}
                //if (checkBox24.Checked == true)//this will just sync to the sloved location
                //{
                //    scope.SyncToCoordinates(RA, DEC);
                //    Log("synced to:  RA = " + scope.RightAscension.ToString() + "     Dec = " + scope.Declination.ToString());
                //}
            }
            catch (Exception e)
            {
                Log("Solve failed");
                FileLog("Solve Failed" + e.ToString());
            }


        }

        private void SolveToleranceCheck()
        {
            Log("Plate Solve Position Check");
            //scope.SlewToCoordinates(RA, DEC);
            CurrentRA = scope.RightAscension;
            CurrentDEC = scope.Declination;
            // first attempts at comparing parse solve coords w/ scope coords.
            //need to FIX...seems to maintain RA and DEC from first plate solve after the second one. 
            //if (usingASCOM == true)
            //{
            //    //  scope = new ASCOM.DriverAccess.Telescope(devId);
            //    scope.SlewToCoordinates(CurrentRA, CurrentDEC);//go back to where originally though it was supposed to be
            //                                                   //should be closer after the sync
            //}
            //else
            //    MessageBox.Show("Must use ASCOM mount connection", "scopefocus");

            if ((Math.Abs(CurrentRA - RA) * 60 > Convert.ToDouble(textBox59.Text)) || (Math.Abs(CurrentDEC - DEC) * 60 > Convert.ToDouble(textBox59.Text)))//*************untested!!****
            {
                scope.SyncToCoordinates(RA, DEC);//sync to parsed(solve) location 
                Log("repeating" + "DeltaRA = " + ((Math.Abs(CurrentRA - RA) * 60).ToString()) + "     DeltaDEC = " + ((Math.Abs(CurrentDEC - DEC) * 60).ToString()));
                //     button55.PerformClick();//prob dont need since fsw7 still on
                fileSystemWatcher7.SynchronizingObject = this;
                if (fileSystemWatcher7.EnableRaisingEvents == false)
                    fileSystemWatcher7.EnableRaisingEvents = true;
                SetForegroundWindow(Handles.NebhWnd);  //rem'd to testing 
                PostMessage(Handles.CaptureMainhWnd, BN_CLICKED, 0, 0);//rem'd for testing
            }
            else
            {
                scope.SyncToCoordinates(RA, DEC);
                Log("sync tolerance met");
                fileSystemWatcher7.EnableRaisingEvents = false;
                Log("synced to:  RA = " + scope.RightAscension.ToString() + "     Dec = " + scope.Declination.ToString());
                Thread.Sleep(500);
                if (!UseClipBoard.Checked) // 11-8-16  TODO not sure about why this is here
                    serverStream.Flush();
                Thread.Sleep(1000);
                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(2000);
                //    NebListenOn = false;
                // clientSocket.GetStream().Close();//added 5-17-12
                //  clientSocket.Client.Disconnect(true);//added 5-17-12
                if (!UseClipBoard.Checked)
                    clientSocket.Close();
            }
        }


        private async void button55_Click(object sender, EventArgs e)
        {
            //if (GlobalVariables.LocalPlateSolve)
            //{
            //    Astrometry ast = new Astrometry();
            //    ast.OnlineSolve(solveImage);
            //}

            //    if (lo
            button55.BackColor = Color.Lime;
            if (checkBox26.Checked == true)
            {
                fileSystemWatcher7.SynchronizingObject = this;
                fileSystemWatcher7.EnableRaisingEvents = true;
                fileSystemWatcher7.Path = GlobalVariables.Path2;
                Log("Plate Solve, wating for Nebulosity capture");
                FileLog2("Plate Solve - Nebulosity Caputre");
            }
            else
            {
                if (GlobalVariables.LocalPlateSolve)
                    Solve();
                else
                {
                    //if (backgroundWorker1.IsBusy != true)
                    //{
                    //    // Start the asynchronous operation.
                    //    backgroundWorker1.RunWorkerAsync();
                    //}
                    Log("Attempt Plate Solve: " + GlobalVariables.SolveImage);
                    FileLog2("Plate Solve: " + GlobalVariables.SolveImage);
                    solveRequested = true;
                    AstrometryRunning = true;
                    // try
                    // {
                    AstrometryNet ast = new AstrometryNet();  // remd for onlyinstance
                                                              // ast = AstrometryNet.OnlyInstance;
                                                              //  AstrometryNet.GetForm.Show();
                    await ast.OnlineSolve(GlobalVariables.SolveImage);
                    // }
                    //catch
                    // {

                    //     AstrometryNet ast = new AstrometryNet();
                    //     await ast.OnlineSolve(GlobalVariables.SolveImage);

                    // }

                    //while (astrometryRunning)
                    //    Thread.Sleep(100);
                    //Solve();
                }


            }
        } // tablisthere
        public void tabList(string file)//start cygwin term, save log, execute tablist command
        {

            try
            {
                Process proc = new Process();
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true;
                //   proc.StartInfo.FileName = @"C:\cygwin\bin\mintty.exe"; // remd 10-22-16
                //  proc.StartInfo.FileName = @"\cygwin_ansvr\bin\mintty.exe";
                // @"c:/cygwin/bin/mintty.exe";
                string filename = @"\cygwin_ansvr\bin\mintty.exe";     //--login -c ""/usr/bin/solve-field -p -O -U none -R none -M none -N none -C cancel--crpix -center -z 2--objs 100 -L .5 -H 2.3 /tmp/stars.fit"" > text.txt";
                proc.StartInfo.FileName = (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + filename);  //orig working
                proc.StartInfo.Arguments = "--log /table.txt -i /Cygwin-Terminal.ico -";
                //creates text file of the cygwin terminal.  
                //parse the needed info from the txt file
                proc.Start();

                StreamWriter sw = proc.StandardInput;
                StreamReader reader = proc.StandardOutput;
                //  StreamReader sr = proc.StandardOutput;
                StreamReader se = proc.StandardError;
                //C:\cygwin\lib\astrometry\bin
                sw.AutoFlush = true;
                Thread.Sleep(2000);
                //    string command = "./tablist " +  "solve.xyls";
                //     SendKeys.Send("cd" + " " + "/home/astro"); // remd 10-22-16
                SendKeys.Send("cd" + " " + "/tmp");
                Thread.Sleep(200);
                SendKeys.Send("~");
                Thread.Sleep(200);
                // SendKeys.Send("solve-field" + " " + "--sigma" + " " + "100" + " " + "-L" + " " + "0.5" + " " + "-H" + " " + "2" + " " + Path.GetFileName(solveImage));
                SendKeys.Send(file);
                Thread.Sleep(200);
                SendKeys.SendWait("~");


                SendKeys.Send("exit");
                SendKeys.Send("~");


                sw.Close();

                reader.Close();

                proc.WaitForExit();



                proc.Close();
            }
            catch (Exception e)
            {
                Log("tablist.exe failed");
                FileLog("tablist.exe failed:  " + e.ToString());
            }

        }
        public static string text;
        private int cygwinId;
        Process proc = new Process();
        // executecommandhere

        public void ExecuteCommand()//starts cygwin term, saves log.txt that captures terminal screen, executes solve command
        {
            try
            {
                sigma = Convert.ToInt32(textBox60.Text.ToString());
                Low = Convert.ToDouble(textBox61.Text.ToString());
                High = Convert.ToDouble(textBox62.Text.ToString());
                DwnSz = Convert.ToInt16(textBox49.Text.ToString());

                //   string stOut = "";
                if (textBox60.Text == "" || textBox61.Text == "" || textBox62.Text == "")
                    MessageBox.Show("Plate solve parameters cannot be blank", "scopefocus");
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true; // 10-22-16 was true
                                                      //   proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal; // new 10-22-16


                string filename = @"\cygwin_ansvr\bin\mintty.exe";     //--login -c ""/usr/bin/solve-field -p -O -U none -R none -M none -N none -C cancel--crpix -center -z 2--objs 100 -L .5 -H 2.3 /tmp/stars.fit"" > text.txt";
                                                                       //  proc.StartInfo.FileName = @"C:\cygwin\bin\mintty.exe";  //orig working
                proc.StartInfo.FileName = (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + filename);  //orig working
                proc.StartInfo.Arguments = "--log /text.txt -i /Cygwin-Terminal.ico -"; // orig working



                proc.Start();
                Thread.Sleep(1000);

                //  cygwinId = proc.Id; // remd 10-22-16


                // rem below 10-22-16




                StreamWriter sw = proc.StandardInput;
                StreamReader reader = proc.StandardOutput;
                //  StreamReader sr = proc.StandardOutput;
                StreamReader se = proc.StandardError;

                sw.AutoFlush = true;
                Thread.Sleep(2000);


                //     string command = "solve-field --sigma " + sigma.ToString() + " -z  "+ DwnSz+ " -N none --no-plots -L " + Low.ToString() + " -H " + High.ToString() + " solve.fit";  // Original working command line
                string command = "/usr/bin/solve-field --sigma " + sigma.ToString() + " -z  " + DwnSz + " -O --objs 100 -N none --no-plots -L " + Low.ToString() + " -H " + High.ToString() + " " + Path.GetFileName(GlobalVariables.SolveImage);
                //   string command = "/usr/bin/solve-field --sigma " + sigma.ToString() + " -z  " + DwnSz + " -O -N none --no-plots -L " + Low.ToString() + " -H " + High.ToString() + " " + GlobalVariables.SolveImage;

                //  string command = "solve-field --sigma " + sigma.ToString() + " -L " + Low.ToString() + " -H " + High.ToString() + " solve.fit";
                //  SendKeys.Send("cd" + " " + "/home/astro");  changed 10-22-16
                SendKeys.Send("cd" + " " + "/tmp");
                Thread.Sleep(200);
                SendKeys.Send("~");
                Thread.Sleep(200);
                // SendKeys.Send("solve-field" + " " + "--sigma" + " " + "100" + " " + "-L" + " " + "0.5" + " " + "-H" + " " + "2" + " " + Path.GetFileName(solveImage));
                SendKeys.Send(command);
                Thread.Sleep(500); // was 200
                SendKeys.SendWait("~");
                //     Thread.Sleep(2000); // was rem'd
                //  text = reader.ReadToEnd();
                //  Thread.Sleep(5000);
                SendKeys.Send("exit");
                SendKeys.Send("~");




                /*
               while(true)
                {

                if(sr.Peek() >= 0)
                                {
                                Console.WriteLine("sr.Peek = " + sr.Peek());
                                Console.WriteLine("sr = " + sr.ReadLine());
                                }

                               if(se.Peek() >= 0)
                                {
                                Console.WriteLine("se.Peek = " + se.Peek());
                                Console.WriteLine("se = " + se.ReadLine());
                                }
                                }

                               */

                while (!IsFileReady((Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\cygwin_ansvr\text.txt")))
                    Thread.Sleep(20);



                sw.Close();
                // sr.Close();

                reader.Close();

                proc.WaitForExit();


                proc.Close();

            }
            catch (Exception e)
            {
                Log("Execute Command Failed");
                FileLog("Execute Command Failed " + e.ToString());
            }
        }


        public static bool IsFileReady(String sFilename)
        {
            // If the file can be opened for exclusive access it means that the file
            // is no longer locked by another process.
            try
            {
                using (FileStream inputStream = File.Open(sFilename, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    if (inputStream.Length > 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                }
            }
            catch (Exception)
            {
                return false;
            }
        }





        double TimeToFlip;
        double CurrentRA;
        double CurrentDEC;
        bool FlipNeeded = false;
        int idleCount = 0;
        private void timer2_Tick(object sender, EventArgs e)
        {
            // inteval set at 1 second
            try
            {
                //  Log("Idle Count " + idleCount.ToString());
                //11-20-16  try to add idle monitor
                if (checkBox34.Checked)
                {
                    if (toolStripStatusLabel1.Text.Length > 9)
                    {
                        if ((toolStripStatusLabel1.Text.Substring(0, 10) == "Capturing ") && (sequenceRunning)) // once neb starts an exposure the numbers show up after capturing
                        { //need to wait a bit before cunting since starting sequence takes a few seconds.  
                            idleCount++;

                            if (idleCount == CaptureTime3 * 1.2 + 10)  //give it about 1.5 time capttime if nothing happening send message
                            {
                                Log("Idle Timnout");
                                FileLog("Idle Timeout");
                                Send("****Idle Timeout***");
                            }
                        }
                    }
                }

                if (checkBox30.Checked == true)
                    LostStarMonitor();
                //*********** this results in error when closing ***********
                //may need 'if scope.connected'
                //    scope = new ASCOM.DriverAccess.Telescope(devId);

                //*****this times out...?? less frequent sampling*********
                textBox53.Text = Math.Round(scope.SiderealTime, 4).ToString();
                textBox54.Text = Math.Round(scope.RightAscension, 4).ToString();
                textBox55.Text = Math.Round(scope.Declination, 4).ToString();
                TimeToFlip = Math.Round(Math.Abs(scope.SiderealTime - scope.RightAscension), 2);
                textBox57.Text = TimeToFlip.ToString();
                //if (TimeToFlip < 0)
                //    FlipNeeded = true;
                //else
                //    FlipNeeded = false;

                if (scope.SideOfPier == scope.DestinationSideOfPier(scope.RightAscension, scope.Declination))
                    FlipNeeded = false;
                else
                {
                    FlipNeeded = true;
                    if (textBox48.Text == "false")
                    {
                        //this way on ly logs/send once because the it hasn't been changed yet 
                        Log("Flip pending in " + TimeToFlip.ToString());
                        FileLog2("Flip pending in " + TimeToFlip.ToString());
                        Send("Flip pending in " + TimeToFlip.ToString());
                    }
                }

                    textBox5.Text = scope.SideOfPier.ToString();
                textBox45.Text = scope.DestinationSideOfPier(scope.RightAscension, scope.Declination).ToString();//see if the slewing to the - 
                                                                                                                 //current location would warrant a flip
                                                                                                                 //  textBoxPHDstatus.Text = ph.getAppState();
            //    ph.PHDgetAppState();
                textBoxPHDstatus.Text = PHD2comm.PHDAppState;

                // TimeSpan ts = TimeSpan.FromHours(Decimal.ToDouble(scope.SiderealTime));
                TimeSpan ts = TimeSpan.FromHours(TimeToFlip);
                string TF;
                int D = ts.Days;
                int H = Math.Abs(ts.Hours);
                int M = Math.Abs(ts.Minutes);
                int S = Math.Abs(ts.Seconds);
                if (ts.Hours < 0 || ts.Minutes < 0 || ts.Seconds < 0)
                    TF = "-" + H.ToString() + ":" + M.ToString() + ":" + S.ToString();
                else
                    TF = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
                textBox56.Text = (TF);
                textBox48.Text = FlipNeeded.ToString();
                if (solveRequested)
                {
                    if (AstrometryNet.OperationCancelled == true)
                    {
                        // ast.Dispose();
                        AstrometryNet.OperationCancelled = false;
                        ast.Dispose();
                        AstrometryRunning = false;
                        solveRequested = false;
                        button55.BackColor = Color.WhiteSmoke;
                        Log("Plate Solve Cancelled by user");
                        return;
                    }

                    if ((!AstrometryRunning) & (!AstrometryNet.OperationCancelled))
                    {
                        // Log("Plate Solve Complete");
                        ast.Dispose();
                        // ast.Close();

                        AstrometryRunning = false;
                        solveRequested = false;
                        Solve();
                    }

                }



            }
            catch (Exception ex)
            {
                // remd 11-8-16
                //   Log("Timer 2 Tick error");
                //  FileLog("Timer 2 Tick error" + ex.ToString());
            }


            //***************** try phd monitor  ***"p******* 4-11-14
            //   if (Capturing)
            //       LostStarMonitor();
            //this seems like it needs to be done in background...causes hang.  
        }
        private int lostStarCount = 0;
        private int recover = 0;
        bool starLost = false;
        private void LostStarMonitor()
        {



            if (starLost == true)   //  *************** added 4-15-15  *************
            {
                //    if (PHDcommand(14) == "1") // try to autoselect new star, if successful, start guiding
                // if (ph.getAppState() == "LockLost")
                if (PHD2comm.PHDAppState == "Stopped")
                {
                    ph.Guide();
                    Thread.Sleep(3000);
                }
                if (PHD2comm.PHDAppState == "LostLock")
               // if (textBoxPHDstatus.Text == "LostLock")
                {
                    ph.StopCapture();
                    Thread.Sleep(1000);
                    ph.Guide();
                    Thread.Sleep(2000);
                    //  PHDcommand(20);
                    //while (ph.getAppState() != "Guiding")
                    //{
                    //    Thread.Sleep(1000);
                    //    ph.Guide();
                    //}
                    //starLost = false;
                    //lostStarCount = 0;
                    //recover = 0;
                    //Log("Guide star AutoSelected - guiding resumed");
                    //Send("AutoSelect Success - Guiding resumed");
                }
            }
            // 12-7-16
            //  textBoxPHDstatus.Text = ph.getAppState();  already done in timer2_tick
            //  textBoxPHDstatus.Text = PHDcommand(PHD_GETSTATUS).ToString();

            //if (textBoxPHDstatus.Text == "4")
            // if (ph.getAppState() == "LockLost")
    //            if ((textBoxPHDstatus.Text == "LostLock") && (!starLost)) // don't need to count anymore once triggered
                if ((PHD2comm.PHDAppState == "LostLock") && (!starLost)) // don't need to count anymore once triggered
                {
                lostStarCount++;
                if (lostStarCount == (int)numericUpDown41.Value)
                {

                    Send("Lost star alert");
                    Log("Lost star notification sent");
                    //   lostStarCount = 0;
                    starLost = true;
                    //********* added 4-15-15 
                }
            }
           // if (textBoxPHDstatus.Text == "3" && lostStarCount > 0)  //if recovers on its own.  
             //   if ((ph.getAppState() == "Guiding") && (lostStarCount > 0))  //if recovers on its own.  
                if ((PHD2comm.PHDAppState == "Guiding") && (lostStarCount > 0))  //if recovers on its own.  
                {
                recover++;
                if ((recover == (int)numericUpDown41.Value) && (starLost))
                {
                    starLost = false;
                    lostStarCount = 0;
                    recover = 0;
                    Log("Guide star locked - guiding resumed");
                    Send("Guiding resumed");
                }
            }

        }

        //  public static string solveImage = "";
        private void textBox58_Click(object sender, EventArgs e)//select image from file browser, clean the solve folder and past the image as "solve.fit"
        {
            DialogResult result = openFileDialog2.ShowDialog();
            GlobalVariables.SolveImage = openFileDialog2.FileName.ToString();
            textBox58.Text = GlobalVariables.SolveImage;
            //  var sourceDir = @"c:\sourcedir";

            // remd 10-22-16
            if (GlobalVariables.LocalPlateSolve)
            {
                // var destDir = @"c:\cygwin\home\astro";
                var destDir = (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)) + @"\cygwin_ansvr\tmp";
                // var pattern = "*.csv";
                // var file = solveImage;
                //   var sourceDir = @"c:\cygwin\home\astro";
                //    var destfile = "solve.fit"; // changed 5-1-16
                var destfile = Path.GetFileName(GlobalVariables.SolveImage);

                if (GlobalVariables.SolveImage != Path.Combine(destDir, Path.GetFileName(GlobalVariables.SolveImage)))
                {
                    // remd 10-22-16

                    foreach (var files in new DirectoryInfo(destDir).GetFiles("*.*"))
                    {
                        if (files.Name != "tablist.exe")
                            files.Delete();//empty the directory
                    }
                    //  File.Copy(solveImage, Path.Combine(destDir, Path.GetFileName(solveImage)));
                    File.Copy(GlobalVariables.SolveImage, Path.Combine(destDir, destfile));
                }
            }
        }

        private void textBox59_TextChanged(object sender, EventArgs e)
        {
            //  CalibrationTol = Convert.ToDouble(textBox59.Text);
        }
        public static IEnumerable<FileInfo> GetLatestFiles(string path, string baseName)
        {
            return new DirectoryInfo(path)
                .GetFiles(baseName + "*.fit")
                .GroupBy(f => f.Extension)
                .Select(g => g.OrderByDescending(f => f.LastWriteTime).First());
        }
        //automatically select last obtained capture from Neb
        private async void fileSystemWatcher7_Created(object sender, FileSystemEventArgs e)
        {
            FileInfo file = new FileInfo(e.FullPath);//returns name of file that triggered FSW7
            Log(file.Name.ToString());
            if (SettingFocusSolve)
            {
                GlobalVariables.FocusImage = GlobalVariables.Path2 + @"\" + file.Name.ToString();
                GlobalVariables.SolveImage = GlobalVariables.FocusImage;
                textBox50.Text = GlobalVariables.SolveImage;
                Thread.Sleep(1000);
                PartialSolve();
                //*************  need error handling for failed solve 4-12-16
                //FocusDEC = DEC;
                //FocusRA = RA;
                //Log("Solved Focus Location:  RA = " + FocusRA.ToString() + "   DEC = " + FocusDEC.ToString());
                //FileLog2("Solved Focus Location:  RA = " + FocusRA.ToString() + "   DEC = " + FocusDEC.ToString());
                //button32.Text = "Solved"; // change the set button only...since not necessarily there (could be from prev session) 
                //button32.BackColor = System.Drawing.Color.Yellow;
                //fileSystemWatcher7.EnableRaisingEvents = false;
                //SettingFocusSolve = false;
                return;
            }
            if (SettingTargetSolve)
            {
                GlobalVariables.TargetImage = GlobalVariables.Path2 + @"\" + file.Name.ToString();
                GlobalVariables.SolveImage = GlobalVariables.TargetImage;
                textBox51.Text = GlobalVariables.SolveImage;
                Thread.Sleep(1000);
                PartialSolve();
                // ****** need error handling for failed solve

                //TargetDEC = DEC;
                //TargetRA = RA;
                //Log("Solved Target Location:  RA = " + FocusRA.ToString() + "   DEC = " + FocusDEC.ToString());
                //FileLog2("Solved Target Location:  RA = " + FocusRA.ToString() + "   DEC = " + FocusDEC.ToString());
                //button34.Text = "Solved"; // change the set button only...since not necessarily there (could be from prev session) 
                //button34.BackColor = System.Drawing.Color.Yellow;
                //fileSystemWatcher7.EnableRaisingEvents = false;
                //SettingTargetSolve = false;
                return;
            }

            if (ConfirmingLocation)
            {

                Thread.Sleep(1000);

                GlobalVariables.SolveImage = file.ToString();
                PartialSolve();
                return;
            }

            GlobalVariables.SolveImage = file.ToString();
            Thread.Sleep(1000);//? remove

            if (GlobalVariables.LocalPlateSolve)
                Solve();
            else
            {
                //if (backgroundWorker1.IsBusy != true)
                //{
                //    backgroundWorker1.RunWorkerAsync();
                //}
                AstrometryNet ast = new AstrometryNet();  // remd for onlyinstance
                                                          //  ast = AstrometryNet.OnlyInstance;
                                                          //    AstrometryNet.GetForm.Show();
                await ast.OnlineSolve(GlobalVariables.SolveImage);
                solveRequested = true;
                //while (astrometryRunning)
                //    Thread.Sleep(100);
                //Solve();

            }
        }
        double RefStarDec;   //postion of star used to sync for polar aligment in degrees
        double RefStarRA;    // RA of PA sync star in hours

        private void button57_Click(object sender, EventArgs e)
        {
            double AzmError = Convert.ToDouble(textBox46.Text);//enter from astrotortilla PA routine in arcminutes
            double AltError = Convert.ToDouble(textBox47.Text);//from AT arcmin
            double M = AzmError * (Math.Cos(scope.Altitude)) / (Math.Cos(RefStarDec));//angular distance corrected altitude and dec or refstar
            double CorrectedRA = RefStarRA - (M / 900);//RA of where star should be if good PA
            double CorrectedDec = RefStarDec - (AltError / 60);///dec same as above
            scope.SlewToCoordinatesAsync(CorrectedRA, CorrectedDec);  //move scope to this position
            Log("slewing to correction position: RA = " + CorrectedRA.ToString() + "     Dec = " + CorrectedDec.ToString());
            MessageBox.Show("Use Mount Alt/Azm adjustments to center refernce star", "scopefocus");
            scope.SyncToCoordinates(CorrectedRA, CorrectedDec);


        }

        private void button58_Click(object sender, EventArgs e)
        {
            RefStarDec = scope.Declination;
            RefStarRA = scope.RightAscension;
            Log("synced:  RA = " + scope.RightAscension.ToString() + "     Dec = " + scope.Declination.ToString());
        }

        //this should be re-written so the arduino is polled as to the state of ContinuousHoldOn to avoid any ambiguity
        //currently it is possible to have them in the opposite state.  reset both will clear if that happens
        //happens if toggled too fast
        bool ContinuousHoldOn = false;
        private void button59_Click(object sender, EventArgs e)
        {


            /*
            if (connect != 1)
            {
                DialogResult result2 = MessageBox.Show("Arduino Not Connected", "Arduino scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
             */
            if (ContinuousHoldOn == false)
            {
                ContinuousHoldOn = true;
                button59.Text = "Hold is ON";
                button59.BackColor = System.Drawing.Color.Lime;
                Log("Continuous Hold set to ON");
            }
            else
            {
                button59.Text = "Cont Hold";
                ContinuousHoldOn = false;
                button59.BackColor = System.Drawing.Color.WhiteSmoke;
                Log("Continuous Hold disabled");
            }
            /*
            if (!usingASCOMFocus)
            {
              //  string str = "";
                //  Thread.Sleep(10);
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
                Thread.Sleep(10);
                port.Write("H");
                Thread.Sleep(50);
                // str = port.ReadExisting();
                Thread.Sleep(50);
                //   port.DiscardInBuffer();
                port.DiscardOutBuffer();
            }
             */
        }
        //bool cygwinAbort = false;
        private void button60_Click(object sender, EventArgs e)//this doesn't work
        {
            this.Show();
            foreach (System.Diagnostics.Process myProc in System.Diagnostics.Process.GetProcesses())
            {
                if (myProc.ProcessName == "mintty")
                {
                    myProc.Kill();
                }
            }


            // cygwinAbort = true;
            //SendKeys.Send("^(C)");
            ////int procID = Convert.ToInt16(002E100E); ****need to get the process number
            ////  (much like finding a handle....then close it.
            //// Process proc = Process.GetProcessById(procID);
            //proc.CloseMainWindow();
            //proc.WaitForExit();

        }

        //start try to monitor metricHFR
        //need to set baseline,then compare and if deviation of > textbox63 % refocus 
        //need to find a place to insert monitorHFR, likely Filesystemwater 4
        //also add PHD monitoring see above
        private int baselineMetricHFR;
        private void checkBox27_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox27.Checked == true)
                if (textBox25.Text == "")
                {
                    DialogResult result = MessageBox.Show("need basline HFR, obtain now?", "scopefocus", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (result == DialogResult.Cancel)
                        return;
                    else
                    {
                        SampleMetric();
                        //  Thread.Sleep(Convert.ToInt16(textBox43.Text) * 2 *(int)numericUpDown21.Value); //wait twice the exposure length x N
                        //    baselineMetricHFR = Convert.ToInt32(textBox25.Text);
                    }


                }

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void button61_Click(object sender, EventArgs e)
        {
            focuser.Halt();

        }

        private void checkBox28_CheckedChanged(object sender, EventArgs e)
        {
            //the only use for action in the driver is reverse so "reverse" is essentially ignored. 
            //sends the reverse command w/ either True or False as arguments, converted to bool in the driver
            //since Action must be string1, string2
            if (checkBox28.Checked == true)
            {
                focuser.Action("Reverse", "True");
                Log("Reversed");

            }
            if (checkBox28.Checked == false)
            {
                focuser.Action("Reverse", "False");

            }

        }

      

        private void tabControl1_Click(object sender, EventArgs e)
        {
            //*****  remd 4-24-14 
            //if (focuser == null)
            //{
            //    MessageBox.Show("Focuser not connected!", "scopefocus");
            //    return;
            //}

        }



        private void numericUpDown6_Enter(object sender, EventArgs e)
        {
            button7.PerformClick();
        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            //   this.numericUpDown6.Maximum = travel;

        }



        private void numericUpDown6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;//should suppress 'ding' w/ enter key but doesn't work!!!

                button7.PerformClick();                   // 

            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }



        //  int backlashComp = 0;

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

            //  backlashComp = Convert.ToInt16(textBox8.Text);
            //    backlash = Convert.ToInt16(textBox8.Text);


            //need to define an arduino command to do this at firmware level
        }



        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            string pos = textBox1.Text;

            if (e.KeyData == Keys.Enter)
            {
                if ((pos == "") || (Convert.ToInt32(pos) < 0) || (Convert.ToInt32(pos) > travel))
                {
                    MessageBox.Show("Invalid position", "scopefocus");
                    textBox1.Text = focuser.Position.ToString();
                    return;
                }
                DialogResult result1 = MessageBox.Show("Manually set focuser motor position to " + pos + "?", "scopefocus", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result1 == DialogResult.OK)
                {
                    focuser.Action("setPosition", pos);
                }
                else
                {
                    textBox1.Text = focuser.Position.ToString();
                    return;
                }
                e.SuppressKeyPress = true;
            }


        }

        //private static string Filter.DevId3;

        //public void Filter.filterWheelChooser() // 1-7-17 moved to separate class
        //{

        //    ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
        //    chooser.DeviceType = "Filter.filterWheel";
        //    Filter.DevId3 = chooser.Choose();
        //    //  ASCOM.DriverAccess.Focuser focuser = new ASCOM.DriverAccess.Focuser(devId2);
        //    if (Filter.DevId3 != "")
        //        Filter.filterWheel = new Filter.filterWheel(Filter.DevId3);
        //    else
        //        return;
        //    Filter.filterWheel.Connected = true;
        //    Log("Filter.filterWheel connected " + Filter.DevId3);
        //    FileLog2("Filter.filterWheel connected " + Filter.DevId3);
        //    buttonFilterConnect.BackColor = System.Drawing.Color.Lime;
        //    if (!checkBox31.Checked)
        //        Filter.filterWheel.Position = 0;
        //    DisplayCurrentFilter();
        //    if (!checkBox31.Checked)
        //    {
        //        foreach (string filter in Filter.filterWheel.Names)
        //            comboBox6.Items.Add(filter);
        //        comboBox6.SelectedItem = Filter.filterWheel.Position;
        //        ComboBoxFill();
        //    }
        //}


        private void ComboBoxFill()
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            comboBox4.Items.Clear();
            comboBox5.Items.Clear();
            comboBox8.Items.Clear();
           
            foreach (string filter in Filter.filterWheel.Names)
            {
                comboBox2.Items.Add(filter);
                comboBox3.Items.Add(filter);
                comboBox4.Items.Add(filter);
                comboBox5.Items.Add(filter);
                comboBox1.Items.Add(filter);
                comboBox8.Items.Add(filter);

            }
        
            comboBox8.Items.Add("Dark 1");
            comboBox8.Items.Add("Dark 2");
            comboBox1.Items.Add("Dark 1");
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 1;
            comboBox4.SelectedIndex = 2;
            comboBox5.SelectedIndex = 3;
            comboBox1.SelectedIndex = 5;
            comboBox8.SelectedIndex = 6;

        }

        private void buttonFilterConnect_Click(object sender, EventArgs e)
        {
            //  WindowsFormsApplication1.EquipConnect c = new WindowsFormsApplication1.EquipConnect();
            //  c.Show();
            //   Filter f = new Filter();
            if (buttonFilterConnect.BackColor != Color.Lime)
            {
                Filter.FilterWheelChooser();
                if (Filter.filterWheel.Connected == true)
                {
                    Log("Filter.filterWheel connected " + Filter.DevId3);
                    FileLog2("Filter.filterWheel connected " + Filter.DevId3);
                    buttonFilterConnect.BackColor = System.Drawing.Color.Lime;
                    if (!checkBox31.Checked)
                        Filter.filterWheel.Position = 0;
                    DisplayCurrentFilter();
                }
                if (!checkBox31.Checked)
                {
                    foreach (string filter in Filter.filterWheel.Names)
                        comboBox6.Items.Add(filter);
                    comboBox6.SelectedItem = Filter.filterWheel.Position;
                    ComboBoxFill();
                }

            }

            else
            {

                Filter.filterWheel.Connected = false;
                buttonFilterConnect.BackColor = Color.WhiteSmoke;
                Filter.filterWheel.Dispose();
                Log(Filter.DevId3 + " disconnected");
                Filter.DevId3 = "";
            }
        }

        private void comboBox6_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (!checkBox31.Checked)
                Filter.filterWheel.Position = (short)comboBox6.SelectedIndex;
            DisplayCurrentFilter();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            //{
            //    button22.PerformClick();
            //    return;
            //}
            SequenceGo();
        }

        private void button19_Click_2(object sender, EventArgs e)
        {


            DialogResult result;
            result = MessageBox.Show("Abort the current sequence?", "scopefocus",
                            MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
            if (result == DialogResult.OK)
            {
                backgroundWorker2.CancelAsync();// added 1-23-13

                ShowWindowAsync(Handles.NebhWnd, SW_SHOW);
                Thread.Sleep(200);
                SetForegroundWindow(Handles.Aborthwnd);
                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);

                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                Thread.Sleep(250);


                fileSystemWatcher4.EnableRaisingEvents = false;
                filterCountCurrent = 0;
                subCountCurrent = 0;
                standby();



                if (!UseClipBoard.Checked)
                {
                    clientSocket2.Close();
                    Thread.Sleep(250);
                    SendKeys.SendWait("~");
                    SendKeys.Flush();
                }
                // ************ end 11-23-13 addt

            }
            if (result == DialogResult.Cancel)
                return;
            //if (paused == false)
            //{
            //    fileSystemWatcher4.EnableRaisingEvents = false;
            //    button22.BackColor = System.Drawing.Color.Lime;
            //    paused = true;
            //    button22.Text = "Resume";
            //    PHDcommand(PHD_PAUSE);
            //    Log("guiding puased");
            //}
            //else
            //{
            //    fileSystemWatcher4.EnableRaisingEvents = true;
            //    button22.UseVisualStyleBackColor = true;
            //    paused = false;
            //    button22.Text = "Pause";
            //    if (Handles.PHDVNumber == 2)
            //        resumePHD2();
            //    else
            //        PHDcommand(PHD_RESUME);
            //}
        }

        // restet sequence
        private void button15_Click(object sender, EventArgs e)
        {
            if (Filter.DevId3 != null)
                ComboBoxFill();
            checkBox1.Checked = true;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox9.Checked = false;
            checkBox13.Checked = false;
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox8.SelectedIndex = -1;
           
            numericUpDown12.Value = 0;
            numericUpDown13.Value = 0;
            numericUpDown14.Value = 0;
            numericUpDown15.Value = 0;
            numericUpDown20.Value = 0;
            numericUpDown29.Value = 0;
            numericUpDown32.Value = 0;
            numericUpDown11.Value = 0;
            numericUpDown16.Value = 0;
            numericUpDown17.Value = 0;
            numericUpDown18.Value = 0;
            numericUpDown19.Value = 0;
            numericUpDown30.Value = 0;
            numericUpDown34.Value = 0;
            numericUpDown24.Value = 0;
            numericUpDown25.Value = 0;
            numericUpDown26.Value = 0;
            numericUpDown27.Value = 0;
            numericUpDown28.Value = 0;
            numericUpDown31.Value = 0;
            numericUpDown36.Value = 0;
            checkBox19.Checked = false;
            checkBox18.Checked = false;
            checkBox15.Checked = false;
            checkBox17.Checked = false;
            checkBox8.Checked = false;


        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            WindowsFormsApplication1.Properties.Settings.Default.None = radioButton8.Checked;
        }

        private void radioButton8_MouseDown(object sender, MouseEventArgs e)
        {
            if (radioButton4.Checked == true)
            {
                radioButton10.Checked = false;
                radioButton9.Checked = false;
                radioButton4.Checked = false;
                _equip = GlobalVariables.EquipPrefix;
            }
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            if ((radioButton10.Checked == true) || (radioButton9.Checked == true))
                radioButton8.Checked = false;
            WindowsFormsApplication1.Properties.Settings.Default.Reducer = radioButton9.Checked;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            WindowsFormsApplication1.Properties.Settings.Default.AO = radioButton4.Checked;
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            if ((radioButton10.Checked == true) || (radioButton9.Checked == true))
                radioButton8.Checked = false;
            WindowsFormsApplication1.Properties.Settings.Default.Reducer = radioButton10.Checked;
        }
        string devId4;
        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                if (SwitchIsConnected)
                {
                    FlatFlap.Connected = false;
                    Log(devId4 + " disconnected");
                    button12.BackColor = System.Drawing.Color.WhiteSmoke;
                    devId4 = "";
                    return;
                }
                ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
                chooser.DeviceType = "Switch";
                devId4 = chooser.Choose();
                if (!string.IsNullOrEmpty(devId4))
                {
                    FlatFlap = new ASCOM.DriverAccess.Switch(devId4);
                    FlatFlap.Connected = true;
                    Thread.Sleep(200);
                    FlatFlap.SetSwitch(0, false);
                }

                else
                {
                    return;
                }
                if (SwitchIsConnected)
                    button12.BackColor = System.Drawing.Color.Lime;
                Log("connected to " + devId4);
                FileLog2("connected to " + devId4);
            }

            catch (Exception ex)
            {
                Log("switch connect error" + ex.ToString());
                Send("switch connect error" + ex.ToString());
                FileLog("switch connect error" + ex.ToString());

            }
        }


            private bool SwitchIsConnected
            {
                get
                {
                    return((this.FlatFlap != null) && (FlatFlap.Connected == true));
                }

            }

            private void textBox8_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.KeyData == Keys.Enter)
                {
                    if ((textBox8.Text == "") || ((Convert.ToInt16(textBox8.Text) < 0)) || ((Convert.ToInt16(textBox8.Text) > travel)))
                    {
                        MessageBox.Show("Invalid backlash setting", "scopefocus");
                     //   textBox1.Text = focuser.Position.ToString();
                        return;
                    }
                   // DialogResult result1 = MessageBox.Show("Manually set focuser motor position to " + pos + "?", "scopefocus", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                  //  if (result1 == DialogResult.OK)
                 //   {
                        
                 //   }
                    else

                    {
                        backlash = Convert.ToInt16(textBox8.Text);
                        return;
                    }
                    //e.SuppressKeyPress = true;
                }

            }

            private void button28_Click(object sender, EventArgs e)
            {
            ph.StopCapture();
              //  PHDSocketPause(true);
                //string pause = PHDcommand(1);
                //Log(pause.ToString());

            }

            private void button31_Click(object sender, EventArgs e)
            {
            ph.Guide();
                //resumePHD2();              
            }
            int resumeAttempts = 0;
            private void resumePHD2()//resumephd2here
            {
            // remd all below 12-10-16
            ph.Guide();
                //if (PHDcommand(2) == "0")//immediate resume...if problems goes further
                //{
                //    if (backgroundWorker2.IsBusy == false)
                //        Log("guiding resumed");
                //    return;
                //}
                //delay(2);
                //bool success = false;
                //while (!success && resumeAttempts < 4)
                //{
                //    string status = PHDcommand(17);
                //    if (backgroundWorker2.IsBusy == false)
                //    Log("waiting to resume guiding");
                //    switch (status)
                //    {
                //        case "3":
                //            success = true;
                //            break;
                //        case "1":
                //            delay(2);
                //            PHDcommand(20);
                //            //delay(2);
                //            //if (PHDcommand(17) == "3")
                //            //{
                //            //    Log("guiding resumed");
                //            //    resumeAttempts = 0;
                //            //}
                //            //else
                //            //{
                //            //    if (resumeAttempts < 3)
                //            //    {
                //            //        resumeAttempts++;
                //            //        resumePHD2();
                //            //    }
                //            //    else
                //            //        Log("resume guiding failed");
                //            //}
                //            break;
                //        case "100":
                //            delay(2);
                //            PHDcommand(2);
                //            break;
                //        case "101":
                //            delay(2);
                //            if (PHDcommand(14) == "1")
                //            {
                //                if (backgroundWorker2.IsBusy == false)
                //                Log("PHD autostar selected");
                //                delay(2);
                //                PHDcommand(20);
                //                // delay(2);
                //                //if (PHDcommand(17) == "3")//checks at end
                //                //    Log("guiding resumed");
                //                //else
                //                //{
                //                //    if (resumeAttempts < 3)
                //                //    {
                //                //        resumeAttempts++;
                //                //        resumePHD2();
                //                //    }
                //                //    else
                //                //        Log("resume guiding failed");
                //                //}
                //            }
                //            else
                //            {
                //                if (backgroundWorker2.IsBusy == false)
                //                Log("PHD autostar select failed");
                //                resumePHD2();
                //            }

                //            break;
                //        case "0"://no looping, paused or guiding
                //            delay(2);
                //            PHDcommand(19);//start looping
                //            delay(2);
                //            string status2 = PHDcommand(17);//check status
                //            switch (status2)
                //            {
                //                case "101"://looping no star selected       
                //                    delay(2);
                //                    if (PHDcommand(14) == "1")
                //                    {
                //                        if (backgroundWorker2.IsBusy == false)
                //                        Log("PHD Autostar selected");
                //                        delay(2);
                //                        PHDcommand(20);//start guiding
                //                    }
                //                    else
                //                        break;
                //                    break;
                //                case "1":
                //                    delay(2);
                //                    PHDcommand(20);
                //                    break;
                //                case "100":
                //                    delay(2);
                //                    PHDcommand(2);
                //                    break;
                //                case "3":
                //                    break;
                //            }
                //            //  else
                //            //   resumePHD2();
                //            break;
                //        case "4":
                //            delay(2);
                //            PHDcommand(18);
                //            delay(2);
                //            PHDcommand(19);
                //            delay(2);
                //            PHDcommand(14);
                //            delay(2);
                //            PHDcommand(20);

                //            break;

                //    }
                //    delay(2);
                //    if (PHDcommand(17) == "3")
                //    {
                //        if (backgroundWorker2.IsBusy == false)
                //        Log("resumed guiding");
                //        resumeAttempts = 0;
                //        success = true;

                //    }

                //    else
                //    {
                //        if (backgroundWorker2.IsBusy == false)
                //        Log("resume guiding failed - retry");
                //        resumeAttempts++;
                //    }

                //}

                //if (backgroundWorker2.IsBusy == false)
                //Log("resume success");
                //return;
                       

            }
        // 11-8-16
        private void msdelay(int ms)
        {
            DateTime t = DateTime.Now;
            do
            {
                System.Windows.Forms.Application.DoEvents();
            }
            while (t.AddMilliseconds(ms) > DateTime.Now);
            return;
        }

        private void delay(int seconds)
            {
                DateTime t = DateTime.Now;
                do
                {
                    System.Windows.Forms.Application.DoEvents();
                }
                while (t.AddSeconds(seconds) > DateTime.Now);
                return;
            }

            

         
            private void button44_Click(object sender, EventArgs e)//Calculate offset
            {
                if ((textBox31.Text == "") || (textBox32.Text == ""))
                {
                    MessageBox.Show("X and Y coordinates are blank");
                    return;
                }
            //if (!GlobalVariables.LocalPlateSolve)
            //{
            //    ParseXYRD();
            //}
            //else
            //{

            //    XYfound = false;
            //    parseXY();
            //    Thread.Sleep(2000);
            //    ParseRD();
            //}
            ParseXYRD();
            if (!XYfound || !RDfound)
                {
                    Log("Parsing failed  XYfound = " + XYfound.ToString() + "   RDfound = " + RDfound.ToString());
                    FileLog2("Parsing failed  XYfound = " + XYfound.ToString() + "   RDfound = " + RDfound.ToString());
                    return;
                }
                else
                {
                    cramersRule();
                  //  button44.BackColor = System.Drawing.Color.Lime;
                  //  button44.Text = "Calc Done";
                }
               // cramersRule();
            }

            private void button43_Click(object sender, EventArgs e)//slew to center here
            {
                if (devId != "")
                {

                    if (centerHereDec == 0 || centerHereRA == 0)
                    {
                        MessageBox.Show("'Center postion' X,Y coordinates cannot be zero", "centerfocus");
                        return;
                    }
                    else
                    {
                        scope.SlewToCoordinates(centerHereRA, centerHereDec);
                        Log("Slew to Offset complete - At RA = " + centerHereRA.ToString() + "     Dec = " + centerHereDec.ToString());
                        FileLog2("Slew to Offset complete - At RA = " + centerHereRA.ToString() + "     Dec = " + centerHereDec.ToString());
                    }
                }
                else
                    MessageBox.Show("No ASCOM telescope connected, connect and retry", "centerfocus");
                //  button43.BackColor = System.Drawing.Color.Lime;
                // button43.Text = "At Offset";
            }

            

            private void textBox60_KeyPress(object sender, KeyPressEventArgs e)
            {
                e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
            }

            private void textBox61_KeyPress(object sender, KeyPressEventArgs e)
            {
                string keyin = e.KeyChar.ToString();
                e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !keyin.Equals(".");//allow decimal point
            }

            private void textBox62_KeyPress(object sender, KeyPressEventArgs e)
            {
                string keyin = e.KeyChar.ToString();
                e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !keyin.Equals(".");
            }

            private void trackBar1_MouseUp(object sender, MouseEventArgs e)
            {
               
                if (devId4 != null)
                {
                    FlatFlap.SetSwitchValue(0, (double)trackBar1.Value);
                    textBox44.Text = trackBar1.Value.ToString();
                }
                else
                {
                    MessageBox.Show("Not connected to ASCOM switch", "scopefocus");
                    return;
                }
            }

            private void button48_Click(object sender, EventArgs e)
            {
                InstallUpdateSyncWithInfo();
            }


            private void InstallUpdateSyncWithInfo()
            {
                UpdateCheckInfo info = null;
                
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;

                    try
                    {
                        info = ad.CheckForDetailedUpdate();

                    }
                    catch (DeploymentDownloadException dde)
                    {
                        MessageBox.Show("The new version of the application cannot be downloaded at this time. \n\nPlease check your network connection, or try again later. Error: " + dde.Message);
                        return;
                    }
                    catch (InvalidDeploymentException ide)
                    {
                        MessageBox.Show("Cannot check for a new version of the application. The ClickOnce deployment is corrupt. Please redeploy the application and try again. Error: " + ide.Message);
                        return;
                    }
                    catch (InvalidOperationException ioe)
                    {
                        MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " + ioe.Message);
                        return;
                    }
                    if (info.UpdateAvailable == false)
                    {
                        MessageBox.Show("The application is up to date", "scopefocus");
                        return;
                    }

                    if (info.UpdateAvailable)
                    {
                        Boolean doUpdate = true;

                        if (!info.IsUpdateRequired)
                        {
                            DialogResult dr = MessageBox.Show("An update is available. Would you like to update the application now?", "Update Available", MessageBoxButtons.OKCancel);
                            if (!(DialogResult.OK == dr))
                            {
                                doUpdate = false;
                            }
                        }
                        else
                        {
                            // Display a message that the app MUST reboot. Display the minimum required version.
                            MessageBox.Show("This application has detected a mandatory update from your current " +
                                "version to version " + info.MinimumRequiredVersion.ToString() +
                                ". The application will now install the update and restart.",
                                "Update Available", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }

                        if (doUpdate)
                        {
                            try
                            {
                                ad.Update();
                                MessageBox.Show("The application has been upgraded, and will now restart.");
                                System.Windows.Forms.Application.Restart();
                            }
                            catch (DeploymentDownloadException dde)
                            {
                                MessageBox.Show("Cannot install the latest version of the application. \n\nPlease check your network connection, or try again later. Error: " + dde);
                                return;
                            }
                        }
                    }
                }
            }

            private void numericUpDown40_ValueChanged(object sender, EventArgs e)
            {
                numericUpDown2.Value = numericUpDown40.Value * 4;
                numericUpDown8.Value = numericUpDown40.Value;
            }

            private void checkBox30_CheckedChanged(object sender, EventArgs e)
            {
                if (checkBox30.Checked)
                FileLog2("Monitor for PHD lost star - On:  " + "Lost start count = " + numericUpDown41.Value.ToString() );
                else
                    FileLog2("Monitor for PHD lost star - Off");
            }

            private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
            {
                //if (ExcelFilename == null)
                //{
                // //   DialogResult result = saveFileDialog1.ShowDialog();
                    ExcelFilename = saveFileDialog1.FileName.ToString();
                    textBox33.Text = ExcelFilename.ToString();
                //}
                using (SqlCeConnection con = new SqlCeConnection(conString))
                {
                    con.Open();
                    using (SqlCeDataAdapter a = new SqlCeDataAdapter("SELECT * FROM table1", con))
                    {
                        DataTable t = new DataTable();
                        a.Fill(t);
                        // dataGridView1.DataSource = t;
                        a.Update(t);
                        object missing = System.Reflection.Missing.Value;
                        if (ExcelFilename.IndexOf(@".") == 0)//if extension not on filename
                        {
                            MessageBox.Show("Filename extension not selected, must select text or Excel file.  Export Aborted", "scopefocus");
                            return;
                        }

                        if (Path.GetExtension(ExcelFilename.Substring(1)) == ".txt" || Path.GetExtension(ExcelFilename.Substring(1)) == ".TXT")
                            TextCSV_FromDataTable(t, ExcelFilename);
                        else
                            Excel_FromDataTable(t);

                    }
                    con.Close();
                  //  textBox33.Text = ExcelFilename.ToString();
                }
            }

            private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
            {
                _importPath = openFileDialog1.FileName.ToString();
                textBox34.Text = _importPath.ToString();
                importDataFromExcel();
                button17.PerformClick();//update
            }

            private void checkBox12_CheckedChanged(object sender, EventArgs e)
            {
                if (checkBox12.Checked == true)
                {
                    if ((textBox13.Text == "") || (textBox27.Text == "") || (textBox28.Text == "") || (textBox36.Text == ""))
                    {
                        MessageBox.Show("There are blank messaging parameters, complete and retry", "scopefocus");
                        checkBox12.Checked = false;
                    }
                }

            }

            private void checkBox31_CheckedChanged(object sender, EventArgs e)
            {
                if (checkBox31.Checked)
                    MessageBox.Show("Make sure your internal filter is set to the lowest number starting position (0 or 1)", "scopefocus");
            }

            private void textBox8_Leave(object sender, EventArgs e)
            {
                if ((textBox8.Text == "") || ((Convert.ToInt16(textBox8.Text) < 0)) || ((Convert.ToInt16(textBox8.Text) > travel)))
                {
                    MessageBox.Show("Invalid backlash setting", "scopefocus");
                    //   textBox1.Text = focuser.Position.ToString();
                    return;
                }
                // DialogResult result1 = MessageBox.Show("Manually set focuser motor position to " + pos + "?", "scopefocus", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                //  if (result1 == DialogResult.OK)
                //   {

                 //   }
                else
                {
                    backlash = Convert.ToInt16(textBox8.Text);
                    return;
                }
            }
        
            private void textBox1_Leave(object sender, EventArgs e)
            {
                
                //******************this was commented out 1-18-15 doubtful any ASCOM focuser would implement.  
                //*****************t opoperly use must first check the focuser supported action arraylist to see if implemented then proceed or abort if not implelemtned. 
                //    also changed position textbox to read only 




                //string pos = textBox1.Text;
                //if ((pos == "") || (Convert.ToInt32(pos) < 0) || (Convert.ToInt32(pos) > travel))
                //{
                //    MessageBox.Show("Invalid position", "scopefocus");
                //    textBox1.Text = focuser.Position.ToString();
                //    return;
                //}
                //DialogResult result1 = MessageBox.Show("Manually set focuser motor position to " + pos + "?", "scopefocus", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                //if (result1 == DialogResult.OK)
                //{
                    
                //    focuser.Action("setPosition", pos);
                //}
                //else
                //{
                //    textBox1.Text = focuser.Position.ToString();
                //    return;
                //}
            }

            private void textBox22_Click(object sender, EventArgs e)
            {
                FindNebCamera();
            }
       
        private void radioButton5_astrometry_CheckedChanged(object sender, EventArgs e)
        {

         //this will alsywa change w/ the other change so nothing needed.  
        }

        private void radioButton_local_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_local.Checked == true)
                GlobalVariables.LocalPlateSolve = true;
            else
                GlobalVariables.LocalPlateSolve = false;
        }

        //private void textBox50_TextChanged(object sender, EventArgs e)//focus image path
        //{
        //    DialogResult result = openFileDialog2.ShowDialog();
        //    GlobalVariables.FocusImage = openFileDialog2.FileName.ToString();
        //    textBox50.Text = GlobalVariables.FocusImage;
        //    //  var sourceDir = @"c:\sourcedir";
        //    //if (GlobalVariables.LocalPlateSolve)
        //    //{
        //    //    var destDir = @"c:\cygwin\home\astro";
        //    //    // var pattern = "*.csv";
        //    //    // var file = solveImage;
        //    //    //   var sourceDir = @"c:\cygwin\home\astro";
        //    //    var destfile = "solve.fit";
        //    //    if (GlobalVariables.SolveImage != Path.Combine(destDir, Path.GetFileName(GlobalVariables.SolveImage)))
        //    //    {
        //    //        foreach (var files in new DirectoryInfo(destDir).GetFiles("*.*"))
        //    //        {
        //    //            if (files.Name != "tablist.exe")
        //    //                files.Delete();//empty the directory
        //    //        }
        //    //        // File.Copy(solveImage, Path.Combine(destDir, Path.GetFileName(solveImage)));
        //    //        File.Copy(GlobalVariables.SolveImage, Path.Combine(destDir, destfile));
        //    //    }
        //    //}
        //}
       
        //private void textBox51_TextChanged(object sender, EventArgs e) // target image path
        //{
        //    DialogResult result = openFileDialog2.ShowDialog();
        //    GlobalVariables.TargetImage = openFileDialog2.FileName.ToString();
        //    // GlobalVariables.SolveImage = openFileDialog2.FileName.ToString();
        //    textBox51.Text = GlobalVariables.TargetImage;
        //    //  var sourceDir = @"c:\sourcedir";
        //    //if (GlobalVariables.LocalPlateSolve)
        //    //{
        //    //    var destDir = @"c:\cygwin\home\astro";
        //    //    // var pattern = "*.csv";
        //    //    // var file = solveImage;
        //    //    //   var sourceDir = @"c:\cygwin\home\astro";
        //    //    var destfile = "solve.fit";
        //    //    if (GlobalVariables.SolveImage != Path.Combine(destDir, Path.GetFileName(GlobalVariables.SolveImage)))
        //    //    {
        //    //        foreach (var files in new DirectoryInfo(destDir).GetFiles("*.*"))
        //    //        {
        //    //            if (files.Name != "tablist.exe")
        //    //                files.Delete();//empty the directory
        //    //        }
        //    //        // File.Copy(solveImage, Path.Combine(destDir, Path.GetFileName(solveImage)));
        //    //        File.Copy(GlobalVariables.SolveImage, Path.Combine(destDir, destfile));
        //    //    }
        //    //}
        //}

        private void textBox50_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog2.ShowDialog();
            GlobalVariables.FocusImage = openFileDialog2.FileName.ToString();
            textBox50.Text = GlobalVariables.FocusImage;

            // added 11-9-16
            if (GlobalVariables.LocalPlateSolve)
            {
                // var destDir = @"c:\cygwin\home\astro";
                var destDir = (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)) + @"\cygwin_ansvr\tmp";
                // var pattern = "*.csv";
                // var file = solveImage;
                //   var sourceDir = @"c:\cygwin\home\astro";
                //    var destfile = "solve.fit"; // changed 5-1-16
                var destfile = Path.GetFileName(GlobalVariables.FocusImage);

                if (GlobalVariables.SolveImage != Path.Combine(destDir, Path.GetFileName(GlobalVariables.FocusImage)))
                {
                    // remd 10-22-16

                    foreach (var files in new DirectoryInfo(destDir).GetFiles("*.*"))
                    {
                        if (files.Name != "tablist.exe")
                            files.Delete();//empty the directory
                    }
                    //  File.Copy(solveImage, Path.Combine(destDir, Path.GetFileName(solveImage)));
                    File.Copy(GlobalVariables.FocusImage, Path.Combine(destDir, destfile));
                }
            }


        }

        private void textBox51_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog2.ShowDialog();
            GlobalVariables.TargetImage = openFileDialog2.FileName.ToString();
            // GlobalVariables.SolveImage = openFileDialog2.FileName.ToString();
            textBox51.Text = GlobalVariables.TargetImage;

            // added 11-9-16
            if (GlobalVariables.LocalPlateSolve)
            {
                // var destDir = @"c:\cygwin\home\astro";
                var destDir = (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)) + @"\cygwin_ansvr\tmp";
                // var pattern = "*.csv";
                // var file = solveImage;
                //   var sourceDir = @"c:\cygwin\home\astro";
                //    var destfile = "solve.fit"; // changed 5-1-16
                var destfile = Path.GetFileName(GlobalVariables.TargetImage);

                if (GlobalVariables.SolveImage != Path.Combine(destDir, Path.GetFileName(GlobalVariables.TargetImage)))
                {
                    // remd 10-22-16

                    foreach (var files in new DirectoryInfo(destDir).GetFiles("*.*"))
                    {
                        if (files.Name != "tablist.exe")
                            files.Delete();//empty the directory
                    }
                    //  File.Copy(solveImage, Path.Combine(destDir, Path.GetFileName(solveImage)));
                    File.Copy(GlobalVariables.TargetImage, Path.Combine(destDir, destfile));
                }
            }
        }

        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (checkBox22.Checked == true) // only need when using ff metric  
            {
                if (e.RowIndex < 0)
                    return;
                selectedID = (int)dataGridView1.SelectedRows[0].Cells[0].Value;
                scrollIndex = dataGridView1.FirstDisplayedScrollingRowIndex;
                IsSelectedRow = true;
                GetAvg();
                //if (checkBox22.Checked)
                //{
                //    textBox12.Text = _enteredPID.ToString();
                //    textBox3.Text = _enteredSlopeDWN.ToString();
                //    textBox10.Text = _enteredSlopeUP.ToString();
                //    textBox16.Clear();
                //    textBox14.Clear();
                //    textBox15.Clear();
                //}
            }
        }

        private void dataGridView1_DataBindingComplete_1(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (checkBox22.Checked == true) // only need for using FF metric focus. 
            {
                if (IsSelectedRow)
                {

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Cells[0].Value.ToString().Equals(selectedID.ToString()))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }
                    dataGridView1.ClearSelection();
                    dataGridView1.Rows[rowIndex].Selected = true;
                }
            }
        }
        private void FocusGroupCalc()
        {
            //if (numericUpDown38.Value == 0) // remd 10-24-16
            //    MessageBox.Show("Value cannot be zero", "scopefoucs");
            //if (numericUpDown38.Value < 4) // remd 10-24-16
            //    MessageBox.Show("Confrim low refocus per sub value of " + numericUpDown38.Value.ToString(), "scopefocus");
            if (!checkBox17.Checked) // rest if unchecked.  
            {
                FocusPerSubGroupCount = 0;
                for (int i=0; i<6; i++)
                {
                    FocusGroup[i] = 0;
                }

            }
            else
            {
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
                if (checkBox5.Checked == true)
                {
                    if (numericUpDown20.Value == 0)
                    {
                        MessageBox.Show("Position 5 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    FocusGroup[4] = (double)numericUpDown20.Value / FocusPerSub;
                    if (FocusGroup[4] != Math.Round(FocusGroup[4]))
                    {
                        MessageBox.Show("Total pos. 4 subs must be multiple of Focus per sub", "scopefocus");
                        numericUpDown38.Value = 1;
                    }
                    else
                        SubsPerFocus[4] = (int)numericUpDown20.Value / (int)FocusGroup[4];
                    if (FocusGroup[4] == 1)//then just use the filter change to refocus
                    {
                        SubsPerFocus[4] = 0;
                        FocusGroup[4] = 0;
                    }
                }
                if (checkBox9.Checked == true)
                {
                    if (numericUpDown29.Value == 0)
                    {
                        MessageBox.Show("Position 6 is selected but has 0 subs", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    FocusGroup[5] = (double)numericUpDown29.Value / FocusPerSub;
                    if (FocusGroup[5] != Math.Round(FocusGroup[3]))
                    {
                        MessageBox.Show("Total pos. 4 subs must be multiple of Focus per sub", "scopefocus");
                        numericUpDown38.Value = 1;
                    }
                    else
                        SubsPerFocus[5] = (int)numericUpDown29.Value / (int)FocusGroup[5];
                    if (FocusGroup[5] == 1)//then just use the filter change to refocus
                    {
                        SubsPerFocus[5] = 0;
                        FocusGroup[5] = 0;
                    }
                }
            }
            Log("Focus groups set: " + FocusGroup[0] + FocusGroup[1] + FocusGroup[2] + FocusGroup[3] + FocusGroup[4] + FocusGroup[5]);
            FileLog2("FocusGroupCount: " + FocusGroup[0] + FocusGroup[1] + FocusGroup[2] + FocusGroup[3] + FocusGroup[4] + FocusGroup[5]);
            FocusPerSubGroupCount = (int)(FocusGroup[0] + FocusGroup[1] + FocusGroup[2] + FocusGroup[3] + FocusGroup[4] + FocusGroup[5]);
        }
    


        private void numericUpDown38_ValueChanged(object sender, EventArgs e)
        {
            // 11-18-16  moved to numericupdown28_leave

        //    if (checkBox17.Checked)
        //    {
        //        //if (numericUpDown38.Value == 0) // remd 10-24-16
        //        //    MessageBox.Show("Value cannot be zero", "scopefoucs");
        //        //if (numericUpDown38.Value < 4) // remd 10-24-16

        //        //    MessageBox.Show("Confrim low refocus per sub value of " + numericUpDown38.Value.ToString(), "scopefocus");
        //        FocusGroupCalc();  // 11-11-16
        //    }
        }



        // 12-2-16 test PHD2 event monitoring
        private void button25_Click(object sender, EventArgs e)
        {
            textBox9.Text = PHD2comm.PHDAppState;


        }
      //  PHD2comm.AppState state = new PHD2comm.AppState();
        private void button52_Click(object sender, EventArgs e)
        {
            ph.StopCapture();


            //ph.getAppState();
            //if (ph.getAppState().ToString() != null)
            //    MessageBox.Show(ph.getAppState().ToString());




                //while (ph.getAppState().ToString() != "Stopped")
                //    Thread.Sleep(100);

                //if (PHD2comm.State != null)
                //    MessageBox.Show(PHD2comm.State.ToString());



                //  ph.StartEventMonitor();
                //   ph.StopCapture();
                //   ph.Loop();
                //  ph.PHD2connect();
        }

        private void button53_Click(object sender, EventArgs e)
        {
             ph.Guide();
        }
      
        private void button36_Click_1(object sender, EventArgs e)
        {


            if (button36.Text == "Connect")
            {
                string result = ph.Connect();
                if (result != "Failed")
                {
                    textBoxPHDstatus.Text = "Connect Ver " + ph.Connect();
                    button36.BackColor = Color.Lime;
                    button36.Text = "DisCnct";
                    Log("Connected to PHD2 version " + result);
                    FileLog("Connected to PHD2 version " + result);
                    
                    
                }
                else
                {
                    textBoxPHDstatus.Text = "No Connection";
                    button36.BackColor = Color.WhiteSmoke;
                    button36.Text = "Connect";
                    Log("PHD2 Connection failed");
                    
                }
            }
            else
            {
                ph.DisConnect();
                textBoxPHDstatus.Text = "No Connection";
                button36.BackColor = Color.WhiteSmoke;
                button36.Text = "Connect";
                Log("PHD2 Disconnected");
              
                FileLog("PHD2 Disconnected");
            }
                // if (textBoxPHDstatus.Text == "Connect Ver ")
                //  textBoxPHDstatus.Text = ph.show("PHDVersion");
            }

      
        private void button51_Click_1(object sender, EventArgs e)
        {
            Handles H = new Handles();
            Callback myCallBack = new Callback(H.EnumChildGetValue);

            H.FindHandles();
            //   int hWnd;


            if (Handles.PHDhwnd == 0)//this can really just serve as reminded that PHD is needed for some functions...not necessary for focusing.  
            {
                DialogResult result1;
                result1 = MessageBox.Show("PHD Not Found - Open and 'Retry', 'Ignore' or 'Abort' to Close", "scopefocus",
                     MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);

                if (result1 == DialogResult.Ignore)
                    this.Refresh();
                if (result1 == DialogResult.Retry)
                {
                    H.FindHandles();
                    this.Refresh();
                }
                if (result1 == DialogResult.Abort)
                    System.Environment.Exit(0);
            }
            else
                Log("PHD version " + Handles.PHDVNumber.ToString());


            if (Handles.NebhWnd == 0)
            {
                DialogResult result;
                result = MessageBox.Show("Nebulosity Not Found - Open or Close & Restart and hit 'Retry' or 'Ignore' to continue",
                   "scopefocus", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);//change so ok moves focus
                if (result == DialogResult.Ignore)
                    this.Refresh();
                if (result == DialogResult.Retry)
                {
                    H.FindHandles();
                    this.Refresh();
                }
                if (result == DialogResult.Abort)
                    System.Environment.Exit(0);

            }
            else
            {

                EnumChildWindows(Handles.NebhWnd, myCallBack, 0);
                Log("Nebulosity Version " + Handles.NebVNumber.ToString());
            }
        }
        // disconnect neb socket
        //private void button52_Click(object sender, EventArgs e)
        //{
        //    // this works 
        //    serverStream = clientSocket.GetStream();
        //    byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
        //    serverStream.Write(outStream, 0, outStream.Length);
        //    serverStream.Flush();

        //    Thread.Sleep(3000);
        //    serverStream.Close();
        //    clientSocket.Close();
            
        //}


        //Neb Abort Button
       // private void button53_Click_1(object sender, EventArgs e)
       // {
       //  //   Thread.Sleep(3000);
       //   //  serverStream.Close();
       //     SetForegroundWindow(Handles.NebhWnd);
       //     Thread.Sleep(1000);
       //     PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
       //     Thread.Sleep(1000);
       ////     NebListenOn = false;
       //     clientSocket.Close();

       //     //  clientSocket.GetStream().Close();//added 5-17-12
       //     //   clientSocket.Client.Disconnect(true);//added 5-17-12
       //     //clientSocket.Close(); // 10-25-16 try rem this....doesn't matter
       // }



        //NEb socket listen
        //private void button54_Click(object sender, EventArgs e)
        //{
        //    NebListenStart(Handles.NebhWnd, SocketPort);
        //}



        //test add ASCOM camera
        //string devId5;
        //private void button25_Click(object sender, EventArgs e)
        //{
        //    if (camIsConnected)
        //    {
        //        cam.Connected = false;
        //        Log(devId5 + " disconnected");
        //        button25.BackColor = System.Drawing.Color.WhiteSmoke;
        //        return;
        //    }
        //    ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
        //    chooser.DeviceType = "Camera";
        //    devId5 = chooser.Choose();
        //    if (!string.IsNullOrEmpty(devId5))
        //    {
        //        cam = new ASCOM.DriverAccess.Camera(devId5);
        //        cam.Connected = true;
        //        Thread.Sleep(200);
        //                   }

        //    else
        //    {
        //        return;
        //    }
        //    if (SwitchIsConnected)
        //        button12.BackColor = System.Drawing.Color.Lime;
        //    Log("connected to " + devId5);
        //    FileLog2("connected to " + devId5);
        //}



        //private bool camIsConnected
        //{
        //    get
        //    {
        //        return ((this.cam != null) && (cam.Connected == true));
        //    }

        //}

        //private void button56_Click(object sender, EventArgs e)
        //{

        //  //  Clipboard.SetText("//NEB CaptureSingle Metric");
        ////    Clipboard.SetText("//NEB Listsen 0");
        //    // cam.StartExposure(5, true);
        //}



        private void button50_Click(object sender, EventArgs e)
        {
            if (!File.Exists((Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\cygwin_ansvr\tmp\" + "tablist.exe")))
            {
                MessageBox.Show("you must place tablist.exe in " + @"c:\users\'username'\appdata\local\cygwin_ansvr\tmp");
                return;
            }

            openFileDialog2.Filter = "All files (*.*)|*.*";
       //     DialogResult result = openFileDialog2.ShowDialog();
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
               
                string file = openFileDialog2.FileName.ToString();
                string path = Path.GetDirectoryName(file);
                if (Path.GetDirectoryName(file) != (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\cygwin_ansvr\tmp"))
                    File.Copy(file, (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\cygwin_ansvr\tmp" + Path.GetFileName(file)), true);

                tablistViewer(Path.GetFileName(file));
                //  cramersRule();
                Log("table.txt saved to " + (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\cygwin_ansvr"));
            }
            else
                return;
        }

       
    }

    //**********this seems to cause error at end of sequence when ? socket close()????
    static class SocketExtensions
    {
        public static bool IsConnected(this Socket socket)
        {
            try
            {
                return !(socket.Poll(1, SelectMode.SelectRead) && socket.Available == 0);
            }
            catch (SocketException) { return false; }

        }

    }

}