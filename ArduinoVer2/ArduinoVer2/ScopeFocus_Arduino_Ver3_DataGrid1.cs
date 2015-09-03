//Arduino version
//added standby to abort v-curve 9-23-11
//this is latest as of 10-12-11
//10-13-11 added gets Max form filename.  still need to use if 2 minHFRs are equally low

//*******need to change default path in setting to c:windows\temp  prior to publish.  
//******this is temporary just for installation and can be changed on first use 


using System;
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




namespace Pololu.Usc.ScopeFocus
{
    public partial class MainWindow : Form
    {
       
        SerialPort port;
        public MainWindow()
        {
            InitializeComponent();
            foreach 
                (var page in tabControl1.TabPages.Cast<TabPage>())
                {
                page.CausesValidation = true; 
                page.Validating += new CancelEventHandler(OnTabPageValidating);
                }
            

                string[] portlist = SerialPort.GetPortNames();
                foreach (String s in portlist)
                {
                    comboBox1.Items.Add(s);
                }
                if (comboBox1.Items.Count == 1)
                {
                    portautoselect = true;
                   
                }

            fileSystemWatcher1.EnableRaisingEvents = false;
            fileSystemWatcher2.EnableRaisingEvents = false;
            fileSystemWatcher3.EnableRaisingEvents = false;
            textBox11.Focus();        
}
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
        int fineVrepeatDone;
        bool fineVrepeatOn = false;
        int fineVrepeat;
        string equip;
        int rows;
        bool GotoFocusOn = false;
        int intersectPos;
        double slopeHFRdwn = 0;
        double slopeHFRup = 0;
        double XintUP;
        double XintDWN;
        int PID;
        int HFRarraymin = 9999;
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
        float backlash =0;
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
        int travel = 4860;//for tsa-120 direct rack and pinion, quarter step
        int connect = 0;
        int temp = 100;
        int portopen = 0;
        int vcurveenable = 0;
        int waittime = 1; // this needs to be the same as EEPROM wait time on arduino, syncs w/ connect
        int closing = 0;
        int syncval;
        int templog;
        int MoveDelay; //helps ensure no focus movement during capture
        int roundto;
        int vN;
        string conString = WindowsFormsApplication1.Properties.Settings.Default.MyDatabase_2ConnectionString;

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
            if (port != null)
            {
                port.Close();
                port.Dispose();
            }
                port = new SerialPort(comboBox1.SelectedItem.ToString(), 9600, Parity.None, 8, StopBits.One);
                port.Open();
                watchforOpenPort();
                if (portopen == 1)
                {
                    Log("Connected to Arduino on " + comboBox1.SelectedItem.ToString());
                    portselected = comboBox1.SelectedItem.ToString();
                }
        }

       void watchforOpenPort()
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

        void positionbar()
        {         
            progressBar2.Maximum = travel;
            progressBar2.Minimum = 0;
            progressBar2.Increment(10);
            progressBar2.Value = count;
        }
        #region logging
        private void Log(Exception e)
        {
            Log(e.Message);
        }

        private void Log(string text)
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
            ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + textBox13.Text.ToString() + "'";
        }

//analyze SQL data  

        void GetAvg()
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
                            textBox12.Text = numb5.ToString();
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
                            textBox3.Text = numb5.ToString();
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
                            textBox10.Text = numb5.ToString();
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
                using (SqlCeCommand com4 = new SqlCeCommand("SELECT AVG(FocusPos) FROM table1 WHERE Equip = @equip", con))
                {
                    com4.Parameters.AddWithValue("@equip", equip);
                    SqlCeDataReader reader4 = com4.ExecuteReader();
                    while (reader4.Read())
                    {
                        if (!reader4.IsDBNull(0))
                        {
                            int numb7 = reader4.GetInt32(0);
                            textBox16.Text = numb7.ToString();
                        }
                    }
                    reader4.Close();
                }

                
                con.Close();
            }
            
        }
        private void WriteSQLdata()
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

      
        private void fileSystemWatcher2_Changed(object sender, FileSystemEventArgs e)
        {
           
          vcurve();
       
        }      
 
//std dev and avg

           private void fileSystemWatcher3_Changed(object sender, FileSystemEventArgs e)
           {
             
               watchforOpenPort();
               if (portopen == 1)
               {
                   int nn = 10;
                   if (vProgress < 10)
                   {
                       int[] list = new int[nn];

                       string[] filePaths = Directory.GetFiles(path2.ToString(), "*.bmp");    
   //remd to use path2 settins    string[] filePaths = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.bmp");
                       roundto = (int)numericUpDown1.Value; 
                       int current = GetFileHFR(filePaths, roundto);
                       list[vProgress] = current;
                           abc[vProgress] = current;
                           sum = sum + list[vProgress];
                           avg = sum / nn;
                       string stdev = Math.Round(GetStandardDeviation(abc), 2).ToString();
                       textBox9.Text = stdev.ToString();
                       textBox14.Text = avg.ToString();
                       Log("Std Dev :\t  N " + (vProgress + 1).ToString() + "\tHFR" + abc[vProgress].ToString() +  "\t  Avg " + avg.ToString() + "\t  StdDev " + stdev.ToString());
                       string strLogText = "Std Dev" + "\t  " + abc[vProgress].ToString() + "\t  " + avg.ToString() + "\t" + (vProgress + 1).ToString() + "\t" + stdev.ToString();
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
                   }
                    if (vProgress == 10)
                   {
                       if (GotoFocusOn == true)
                       {
                           if (radioButton2.Checked == true)//use upslope
                           {
                               GetAvg();
                               double EnteredSlopeUP;
                               int EnteredPID = Convert.ToInt32(textBox12.Text);
                               EnteredSlopeUP = Convert.ToDouble(textBox10.Text);
                               textBox14.Text = avg.ToString();
                               double BestPos = count - (avg / EnteredSlopeUP) + (EnteredPID / 2);
                               gotopos(Convert.ToInt16(BestPos)); 
                               fileSystemWatcher3.EnableRaisingEvents = false;
                               GotoFocusOn = false;
                               Log("Goto Focus Position" + BestPos.ToString());
                               textBox4.Text = Convert.ToInt16(BestPos).ToString();
                           }
                           if (radioButton3.Checked == true)//use downslope
                           { 
                               GetAvg();
                               double EnteredSlopeDWN;
                               int EnteredPID = Convert.ToInt32(textBox12.Text);
                               EnteredSlopeDWN = Convert.ToDouble(textBox3.Text);
                               textBox14.Text = avg.ToString();
                               double BestPos = count - (avg / EnteredSlopeDWN) - (EnteredPID / 2);
                               gotopos(Convert.ToInt16(BestPos)); 
                               fileSystemWatcher3.EnableRaisingEvents = false;
                               GotoFocusOn = false;
                               Log("Goto Focus Position" + BestPos.ToString());
                               textBox4.Text = Convert.ToInt16(BestPos).ToString();
                           }
                       }  
   
                   }
                   
               }
           }        

        void standby()
        { 
            fileSystemWatcher2.EnableRaisingEvents = false;
            fileSystemWatcher3.EnableRaisingEvents = false;
            fileSystemWatcher1.EnableRaisingEvents = false;
            port.DiscardOutBuffer();
            port.DiscardInBuffer();
            positionbar();
            progressBar1.Value = 0;
 
        }


//standby
        private void button4_Click(object sender, EventArgs e)
        {
            standby();
        }
      
        void gotopos(Int16 value)
        {
            
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
                    if (zz % 8 ==0)
                    {
                        Thread.Sleep(5);
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
                       return  ((int)((param2 + (0.5 * round)) / round) * round);
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

         void vcurve()
         {
             
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
                     textBox17.Text = (fineVrepeatDone+1).ToString();
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
                     progressBar1.Value = vProgress+1;
                 }
                 if (vProgress == vN)
                 {
                     vDone = 1;;
                 }             
                 int avg = 0;
                 int current = 0;
                 int closestHFR = 1;
                 long closestMax = 1;
                 if (vDone != 1)
                 {
                     int[] templist = new int[((vN * repeatTotal) + 1)];

                     string[] filePaths = Directory.GetFiles(path2.ToString(), "*.bmp");
                     closestHFR = GetFileHFR(filePaths, roundto);
                     closestMax =  GetFileMax(filePaths,roundto);
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
 
                     //setup positions array for HFR
                     minHFRpos[vProgress] = count;
                     maxMaxPos[vProgress] = count;
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
                         Log("V-Curve: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t\tPos" + count.ToString() + " \t  HFRavg" + avg.ToString() + "\tMaxavg" + avgMax.ToString());
                         string strLogTextA = "V-Curve: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + count.ToString() + " \t" + avg.ToString() + "\t" + avgMax.ToString();
                         string strLogText = "V-curve" + "\t" + temp.ToString() + "\t" + count.ToString() + "\t" + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + current.ToString() + "\t" + min.ToString() + "\t" + posMin.ToString();
                         string strLogText2 = "TempCal" + "\t " + tempavg.ToString() + "\t " + posMin.ToString();
                         string strLogText3 = "Fine-V: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + count.ToString() + " \t" + avg.ToString() + "\t" + avgMax.ToString() + "\t" + min.ToString() + "\t" + maxMax.ToString() + "\t" + posMin.ToString();
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
                                     if (vProgress <(vN-1))
                                     {
                                         gotopos(Convert.ToInt16(count - step));
                                     }
                                     vProgress++;
                                     vv++;
                                     Thread.Sleep(MoveDelay);
                                 }
                                 if (backlashOUT == true)
                                 {
                                     if (vProgress < (vN-1))
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
                             if  (((apexHFR > (vN / 2 + vN/4)) || (apexHFR < (vN / 2 - vN/4))) & (ffenable ==1) & (backlashDetermOn == false))
                             {
                                 fileSystemWatcher2.EnableRaisingEvents = false;
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
                                     intersectPos = (int)GetIntersectPos((double)minHFRpos[apexHFR - vN/4], (double)minHFRpos[apexHFR + vN/4], (double)list[apexHFR - vN/4], (double)list[apexHFR + vN/4], slopeHFRdwn, slopeHFRup);
                                     //all vN/2 above were 5
 
                                     textBox4.Text = posMin.ToString();
                                     textBox15.Text = HFRarraymin.ToString();
                                     WriteSQLdata();;
                                     FillData();
                                 Log("Slope: N " + (vProgress + 1).ToString() + "\tslopeUP" + slopeHFRup.ToString() + " \tSlopeDWN" + slopeHFRdwn.ToString() + "\tIntersect" + intersectPos.ToString() + "\tPID" + PID.ToString());

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
                         if (ffenable ==1)
                         {
                             for (int xx = 0; xx < vN; xx++)
                             {
                                 if (xx == 0)
                                 {
                                     log.WriteLine(DateTime.Now);
                                 }
                                 log.WriteLine("HFR " + list[xx].ToString() + "\tPosition " + minHFRpos[xx].ToString());
                             }
                             logData.WriteLine(DateTime.Now + "\t" +  vN.ToString() + "\t" + slopeHFRdwn.ToString() + "\t" + slopeHFRup.ToString() + "\t" + XintDWN.ToString() + "\t" + XintUP.ToString() + "\t" + PID.ToString() + "\t" + apexHFR.ToString());          
                         }
                         log.Close();
                         logData.Close();
                         if ((vDone == 1) & (ffenable == 1))
                         {

                             fileSystemWatcher1.EnableRaisingEvents = false;
                             fileSystemWatcher2.EnableRaisingEvents = false;
                             //handle repeated fine v curves
                             if (fineVrepeat > 1)
                             {
                                 button3.PerformClick();
                                 
                             }
                             else
                             {
                                 fineVrepeatDone = 0;
                                 fineVrepeatOn = false;
                                 standby();
                             }

                         }
                        //end course v-curve and goes to rough focus point
                         if ((vDone == 1) & (tempon == 0) & (backlashDetermOn == false) & (ffenable !=1))
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
                         backlashCount ++;
                         if (backlashCount == backlashN-1)
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
                     
         }
         private void chart1_Click(object sender, EventArgs e)
         {
             chart1.Series[1].Points.AddXY(Convert.ToDouble(min), Convert.ToDouble(count));
       
         }

        public void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            string path2 = folderBrowserDialog1.SelectedPath.ToString();         
            textBox11.Text = path2.ToString();
        }

//added to sync for manual changes
            void sync()
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

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String strVersion = System.Reflection.Assembly.GetCallingAssembly().GetName().Version.ToString(); 

            MessageBox.Show("           Arduino Version" + strVersion.ToString()+  "\r" + "           Kevin Sipprell MD\r           www.scopefocus.info\r      ", "scopefocus for nebulosity",
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
        
void tempcal ()
        {
           
            fileSystemWatcher1.EnableRaisingEvents = true;
            port.DiscardInBuffer();
                ffenable = 1;
                tempon = 1;
                vcurve();
        }

private void button4_Click_1(object sender, EventArgs e)
{
    standby();
}

        public static int GetBestHFRPos(int Xa, int Ya, int Yb, int ma)
        {
            return (Xa + (Yb - Ya)/ma);
        }


        public static double GetIntersectPos(double upX, double downX, double upY, double downY, double upSlope, double downSlope)
        {
            return (((upSlope * upX) - (downSlope * downX) + downY - upY) / (upSlope - downSlope));
        }

public static double GetPositionHFR(int Xa, int Ya, int Xb,  int Yb, int Xc, int Yc, int Xd, int Yd)
{
            double m1 = (Ya - Yb) / (Xa - Xb);
            double m2 =(Yc-Yd)/(Xc-Xd);  
            return (((m1 * Xa) - (m2 * Xc) + Yc - Ya) / (m1 - m2));
}

public static double GetPositionMax(double Xa, double Ya, double Xb, double Yb, double Xc, double Yc, double Xd, double Yd)
{   
    double m1 = (Ya - Yb) / (Xa - Xb);
    double m2 = (Yc - Yd) / (Xc - Xd);
    return (((m1 * Xa) - (m2 * Xc) + Yc - Ya) / (m1 - m2));
}

//goes to near focus point uses std dev routine, takes 10 exposures, calculates best focus and goes there
public void gotoFocus()
{
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
    fileSystemWatcher3.EnableRaisingEvents = true;
    fileSystemWatcher1.EnableRaisingEvents = false;

    GotoFocusOn = true;//~line 440
    min = 1;
    sum = 0;
    avg = 0;
    vDone = 0;
    vProgress = 0;
    vv =0;
    list = new int[10];
    abc = new double[10];
}


private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
{
   
}

//update button
private void button13_Click(object sender, EventArgs e)
{
    GetAvg();
    FillData();  
}

private void listView1_SelectedIndexChanged(object sender, EventArgs e)
{

}

private void MainWindow_Load_1(object sender, EventArgs e)
{
    radioButton2.Checked = true;
   //5 lines below for path2 settings
    path2 = WindowsFormsApplication1.Properties.Settings.Default.path.ToString();
    textBox11.Text = path2.ToString();
    fileSystemWatcher1.Path = path2.ToString();
    fileSystemWatcher2.Path = path2.ToString();
    fileSystemWatcher3.Path = path2.ToString();

    setarraysize();
    string[] str = new string[6];
    str[0] = WindowsFormsApplication1.Properties.Settings.Default.setup1;
    str[1] = WindowsFormsApplication1.Properties.Settings.Default.setup2;
    str[2] = WindowsFormsApplication1.Properties.Settings.Default.setup3;
    str[3] = WindowsFormsApplication1.Properties.Settings.Default.setup4;
    str[4] = WindowsFormsApplication1.Properties.Settings.Default.setup5;
    str[5] = WindowsFormsApplication1.Properties.Settings.Default.setup6;

    for (int i = 0; i < 6; i++)
    {
        listView1.Items.Insert(i, str[i], str[i], 0);
    }
    this.tabControl1.SelectedTab = tabPage2;
    comboBox1.Focus();//sets active window at port selection

    //this removes the default listivew items not sure how else to do this
   
    int listcount = this.listView1.Items.Count;
    
    if (listcount > 6)
    {
        for (int i = listcount; i > 6; i--)
        {
            listView1.Items.RemoveAt(i-1);
        }
    }
    FillData();
    
}

private void ButtonDisable_Click_1(object sender, EventArgs e)
{
    WindowsFormsApplication1.Properties.Settings.Default.port = portselected;
    WindowsFormsApplication1.Properties.Settings.Default.path = path2.ToString();
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

private void button4_Click_2(object sender, EventArgs e)
{
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
    setarraysize();
    watchforOpenPort();
    if (portopen == 1)
    {     
        if ((textBox20.Text =="" || textBox18.Text == "") & (radioButton1.Checked == true))
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
        enteredMaxHFR  = Convert.ToInt16(textBox20.Text.ToString());  
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
        if ((finegoto-100) < 0)
        {
            DialogResult result1 = MessageBox.Show("Goto Exceeds Full In\rAdditional backfocus needed to take-up backlash and center V-curve", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (result1 == DialogResult.OK)
                {
                  return;                  
                }
            }
        gotopos(Convert.ToInt16(finegoto - 100));//take up backlash     
        Thread.Sleep(1000);
        gotopos(Convert.ToInt16(finegoto - ((int)numericUpDown8.Value)*4));
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
        HFRarraymin = 500;     
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
        fileSystemWatcher2.EnableRaisingEvents = true;
        fileSystemWatcher3.EnableRaisingEvents = false;
        fileSystemWatcher1.EnableRaisingEvents = false;
    }
}

void fineVrepeatcounter()
{
    if ((fineVrepeatOn == true) & (fineVrepeat>1))
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

private void button1_Click_1(object sender, EventArgs e)
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

private void button9_Click_1(object sender, EventArgs e)
{
    sync();
    count = syncval;
}

private void button7_Click_1(object sender, EventArgs e)
{   
    if (portopen == 1)
    {
        fileSystemWatcher2.EnableRaisingEvents = false;
        fileSystemWatcher3.EnableRaisingEvents = false;
        int go2 = (int)numericUpDown6.Value;
        if (go2 < 0)
        {
                DialogResult result1 = MessageBox.Show("Goto Exceeds Full In", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (result1 == DialogResult.OK)
                {
                    return;
                }
            }
        gotopos(Convert.ToInt16(go2));
        Thread.Sleep(20);        
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
//std dev button
private void Button2000_Click_1(object sender, EventArgs e)
{
    port.DiscardInBuffer();
    port.DiscardOutBuffer();
    fileSystemWatcher2.EnableRaisingEvents = false;
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


private void listView1_DoubleClick_1(object sender, EventArgs e)
{
    if (this.listView1.SelectedItems.Count == 1)
    {
        this.listView1.SelectedItems[0].BeginEdit();
    }
}

private void listView1_Validated(object sender, EventArgs e)
{
    string test0 = this.listView1.Items[0].Text.ToString();
    WindowsFormsApplication1.Properties.Settings.Default.setup1 = test0.ToString();
    string test1 = this.listView1.Items[1].Text.ToString();
    WindowsFormsApplication1.Properties.Settings.Default.setup2 = test1.ToString();
    string test2 = this.listView1.Items[2].Text.ToString();
    WindowsFormsApplication1.Properties.Settings.Default.setup3 = test2.ToString();
    string test3 = this.listView1.Items[3].Text.ToString();
    WindowsFormsApplication1.Properties.Settings.Default.setup4 = test3.ToString();
    string test4 = this.listView1.Items[4].Text.ToString();
    WindowsFormsApplication1.Properties.Settings.Default.setup5 = test4.ToString();
    string test5 = this.listView1.Items[5].Text.ToString();
    WindowsFormsApplication1.Properties.Settings.Default.setup6 = test5.ToString();
    WindowsFormsApplication1.Properties.Settings.Default.Save();

   
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
        }
    }
    else
    {
        comboBox1.Focus();
    }
}


private void button8_Click_2(object sender, EventArgs e)
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
    if (connect == 0)
    {
        //added for maxtravel sync
        int syncfailure2 = 0;
        string numbers2 = "";
        int numb2 = 0;
        string str2 = "";
        textBox2.Text = str2;
        Thread.Sleep(20);
        port.DiscardOutBuffer();

        port.DiscardInBuffer();
        Thread.Sleep(20);
        port.Write("T");
        Thread.Sleep(200);
        str2 = port.ReadExisting();
        Thread.Sleep(500);
        port.DiscardInBuffer();
        Thread.Sleep(50);
        port.DiscardOutBuffer();
        Thread.Sleep(50);
        numbers2 = string.Join(null, System.Text.RegularExpressions.Regex.Split(str2, "[^\\d]"));
        if (numbers2 == "")
        {
            DialogResult result;
            syncfailure2 = 1;
            Log("Sync Failed -- Retry");
            result = MessageBox.Show("Sync Failure -- Retry ", "Sync Failure",
                        MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
            if (result == DialogResult.Retry)
            {
                button8.PerformClick();

            }
        }
        numb2 = Int32.Parse(numbers2);
        if ((numb2 < 1000) || (numb2 > 100000))
        {
            DialogResult result;
            syncfailure2 = 1;

            Log("Data Sync Failed -- Set to Default");
            result = MessageBox.Show("Data Sync Out of Range -- Set to Default ", "Data Sync Failure",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {

                textBox2.Text = travel.ToString();

            }
        }
        if (syncfailure2 != 1)
        {

            travel = numb2;
            textBox2.Text = travel.ToString();
        }
        //stepper delay sync
        Thread.Sleep(500);

        int syncfailure3 = 0;
        string numbers3 = "";
        int numb3 = 0;
        textBox5.Clear();
        string str3 = "";
        Thread.Sleep(20);
        port.DiscardOutBuffer();
        port.DiscardInBuffer();
        Thread.Sleep(20);
        port.Write("S");
        Thread.Sleep(200);
        str3 = port.ReadExisting();
        Thread.Sleep(50);
        port.DiscardInBuffer();
        numbers3 = string.Join(null, System.Text.RegularExpressions.Regex.Split(str3, "[^\\d]"));
        numb3 = Int32.Parse(numbers3);
        if (numbers3 == "")
        {
            DialogResult result;
            syncfailure3 = 1;
            port.Write("E");
            Log("Sync Failed -- Retry");
            result = MessageBox.Show("Sync Failure -- Retry ", "Sync Failure",
                        MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
            if (result == DialogResult.Retry)
            {
                button8.PerformClick();

            }
        }
        if ((numb3 < 1) || (numb3 > 100))
        {
            DialogResult result;
            syncfailure3 = 1;
            Log("Data Sync Failed -- Set to Default");
            result = MessageBox.Show("Data Sync Out of Range -- Set to Default 20", "Data Sync Failure",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                waittime = 20;
                textBox5.Text = waittime.ToString();
            }
        }
        if (syncfailure3 != 1)
        {
            waittime = numb3;
            textBox5.Text = waittime.ToString();
        }
    }
    else
    {
        DialogResult result2 = MessageBox.Show("Already Connected", "Arduino scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    connect = 1;
    WindowsFormsApplication1.Properties.Settings.Default.port = port.ToString();
    WindowsFormsApplication1.Properties.Settings.Default.Save();
}
 
//this ensures only one listview item(equip) is checked
bool isUnchecked;
bool whenUnchecked;
bool isChecking;
bool canCheck;

private void listView1_MouseClick_1(object sender, MouseEventArgs e)
{
    if (!listView1.GetItemAt(e.X, e.Y).Checked)
    {
        canCheck = true;
        listView1.GetItemAt(e.X, e.Y).Checked = true;
    }
    else
        isUnchecked = true;
}

private void listView1_ItemCheck_1(object sender, ItemCheckEventArgs e)
{
    if (isUnchecked)
    {
        isUnchecked = false;
        isChecking = true;
        listView1.Items[e.Index].Checked = false;
        e.NewValue = CheckState.Unchecked;
        //new code for no items checked
        isChecking = false;
        whenUnchecked = true;
    }

    if (!isChecking && canCheck)
    {
        isChecking = true;
        foreach (ListViewItem item in listView1.Items)
        {
            item.Checked = false;
        }
        listView1.Items[e.Index].Checked = true;
        e.NewValue = CheckState.Checked;
        canCheck = false;
        isChecking = false;
    }
    else
    {
        if (isChecking || whenUnchecked)
        {
            e.NewValue = CheckState.Unchecked;
            whenUnchecked = false;
        }
        else
        {
            e.NewValue = e.CurrentValue;
        }
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

}

 //reset defaults 
private void button14_Click(object sender, EventArgs e)
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
         WindowsFormsApplication1.Properties.Settings.Default.port = @"COM";
         WindowsFormsApplication1.Properties.Settings.Default.Save();
     }
}

private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
{
    portselected = comboBox1.SelectedItem.ToString();
}

private void button15_Click(object sender, EventArgs e)
{
    WindowsFormsApplication1.Properties.Settings.Default.port = portselected;
    WindowsFormsApplication1.Properties.Settings.Default.path = path2.ToString();
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
   
    WindowsFormsApplication1.Properties.Settings.Default.port = portselected;
    WindowsFormsApplication1.Properties.Settings.Default.path = path2.ToString();
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

}

private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
{
    if ((listView1.CheckedItems.Count != 0) && (tabControl1.SelectedTab == tabPage1))
    {
        tabControl1.SelectedTab = tabPage1;
        int n = listView1.CheckedItems.Count;
        ListView.CheckedListViewItemCollection checkedItems = listView1.CheckedItems;
        ListView.CheckedIndexCollection k = listView1.CheckedIndices;
        equip = listView1.Items[k[n - 1]].Text;
        textBox13.Text = equip.ToString();
        //next line filters gridview by equip
        ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + textBox13.Text.ToString() + "'";
    }
    else if ((listView1.CheckedItems.Count == 0) && (tabControl1.SelectedTab == tabPage1))
    {
        MessageBox.Show("Must Select Equipment First", "scopefocus");
        tabControl1.SelectedTab = tabPage2;

    }
}
//deletes all SQL data for this Equip
private void button16_Click(object sender, EventArgs e)
{
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
//View All button
private void button17_Click(object sender, EventArgs e)
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
    fileSystemWatcher3.EnableRaisingEvents = false;
    fileSystemWatcher1.EnableRaisingEvents = true;
}
//tempcal button
private void button5_Click(object sender, EventArgs e)
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
    fileSystemWatcher3.EnableRaisingEvents = false;
    fileSystemWatcher1.EnableRaisingEvents = true;
    //  tempcal();
}
//deletes selected rows
private void button18_Click(object sender, EventArgs e)
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










    }




}