//Arduino version
//added standby to abort v-curve 9-23-11
//this is latest as of 10-12-11
//10-13-11 added gets Max form filename.  still need to use if 2 minHFRs are equally low
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


namespace Pololu.Usc.ScopeFocus
{
    public partial class MainWindow : Form
    {
        SerialPort port;
        public MainWindow()
        {
            
        
            InitializeComponent();
   //rem below if using auto open port
            
            string[] portlist = SerialPort.GetPortNames();
            foreach (String s in portlist)
            {
                comboBox1.Items.Add(s);
           }
   
//to here

//Un-rem next line below for auto port open com11
           // PortOpen();
           
            //    MessageBox.Show("Make Sure Neb is NOT capturing in Fine Focus \rMake sure focus is IN from rough guess", "Focus Control",
            //                MessageBoxButtons.OK, MessageBoxIcon.Information);
            fileSystemWatcher1.EnableRaisingEvents = false;
            fileSystemWatcher2.EnableRaisingEvents = false;
            fileSystemWatcher3.EnableRaisingEvents = false;
            DialogResult result = folderBrowserDialog1.ShowDialog();
            string path2 = folderBrowserDialog1.SelectedPath.ToString();         
            textBox11.Text = path2.ToString();
            fileSystemWatcher1.Path = folderBrowserDialog1.SelectedPath;
            fileSystemWatcher2.Path = folderBrowserDialog1.SelectedPath;
            fileSystemWatcher3.Path = folderBrowserDialog1.SelectedPath;

           
}
        int MultiVcurveProgress = 0;
        double BestHFR = 0;
        double slopeHFRdwn = 0;
        double slopeHFRup = 0;
        double[] AVGslopeUP = new double[20];
        double[] AVGslopeDWN = new double[20];
        double[] AVGintersectPOS = new double[20];
        double[] AVGintersectHFR = new double[20];
        double[] AVGPID = new double[20];
        double[] AVGHFR = new double[20];
        double[] AVGxintUP = new double[20];
        double[] AVGxintDWN = new double[20];
        double[] AVGintersecPOSdiff = new double[20];
        double XintUP;
        double XintDWN;
        double PID;
        bool MaxtooHi = false;
        int HFRarraymin = 9999;
        double maxarrayMax = 1;
        int apexHFR;
     //   int apexMax;
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
   //     int c = 0;
    //    int d = 0;
        int sum = 0;
        int avg = 0;
        long sumMax = 0;
        long avgMax = 0;
        int ffenable = 0;
        int repeatTotal = 0;
        int vN = 0;
        int[] list = new int[100];
        double[] listMax = new double[100];
        double[] abc = new double[100];
        int[] peat = new int[10];
        double[] peatMax = new double[10];
        int[] minHFRpos = new int[100];
        double[] maxMaxPos = new double[100];
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
//THis will auto open com11 used only for dell studio 15 laptop, top left USB port
       void PortOpen()
        {
            if (port != null)
            {
                port.Close();
                port.Dispose();
            }
           //use below for comm11 auto open
         //  port = new SerialPort("com11", 9600, Parity.None, 8, StopBits.One);
            port = new SerialPort(comboBox1.SelectedItem.ToString(), 9600, Parity.None, 8, StopBits.One);
           //use above for selecting from list
            port.Open();
            watchforOpenPort();
            if (portopen == 1)
            {
                Log("Connected to Arduino on " + comboBox1.SelectedItem.ToString());
            }
          //  port.Write(go.ToString());
        }

       void watchforOpenPort()
       {
          if (port == null)
           {
               portopen = 0;
               DialogResult result2 = MessageBox.Show("Arduino Not Connected", "Arduino scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
              // if (result2 == DialogResult.OK)
                   
               

               }
               else
               {
                   portopen = 1;
               }
           
       }

 
        void positionbar()
        {
            watchforNegFocus();
        //    travel = (int)numericUpDown3.Value;
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
       
        //addition 12-22
 
        void watchforNegFocus()
        {
            if (count < 0)
            {
                count = 0;
                DialogResult result1 = MessageBox.Show("Goto Exceeds Full In", "Focus Control", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result1 == DialogResult.Cancel)
                {
                    count = 0;
                    return;
                }
            }
        }
      



// Fine Focus
     
        //reset V-curve
        private void button6_Click(object sender, EventArgs e)
        {  

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
                watchforNegFocus();
                maxMax = 0;
                fileSystemWatcher2.EnableRaisingEvents = true;
                fileSystemWatcher3.EnableRaisingEvents = false;
                fileSystemWatcher1.EnableRaisingEvents = false;
                avgMax = 0;
                
                sumMax = 0;
               
            }
         }

        private void fileSystemWatcher2_Changed(object sender, FileSystemEventArgs e)
        {
           vcurve();
       
        }      
 

    


//std dev and avg
           void Button2000_Click(object sender, EventArgs e)
           {
               port.DiscardInBuffer();
               port.DiscardOutBuffer();
               fileSystemWatcher2.EnableRaisingEvents = false;
               fileSystemWatcher3.EnableRaisingEvents = true;
               fileSystemWatcher1.EnableRaisingEvents = false;


               posMin = 0;
               min = 1;
               sum = 0;
               avg = 0;
               vDone = 0;
               vProgress = 0;
           }
           private void fileSystemWatcher3_Changed(object sender, FileSystemEventArgs e)
           {
             
               watchforOpenPort();
               if (portopen == 1)
               {
                   int nn = 10;
                   if (vProgress < 10)
                   {
                       int[] list = new int[nn];
                       string[] filePaths = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.bmp");
                       foreach (string filename in filePaths)
                       {
                           int first = filename.IndexOf(@".");
                           int second = filename.IndexOf(".", first + 1);
                           string size = filename.Substring(first - 2, 1);
                           int num;
                           int current = -1;
                           bool isNumeric = int.TryParse(size, out num);
                           if (isNumeric)
                           {
                               string current2 = filename.Substring(first - 2, second - first);
                               string replace = current2.Replace(".", "");
                               current = Convert.ToInt32(replace);

                               //end addition
                           }
                           else
                           {
                               // value is not a number
                               string current2 = filename.Substring(first - 1, second - first - 1);
                               //added to convert to int
                               string replace = current2.Replace(".", "");
                               current = Convert.ToInt32(replace);
                               //end addition
                           }
                           list[vProgress] = current;
                           abc[vProgress] = current;
                           sum = sum + list[vProgress];
                           avg = sum / nn;
                       }
                       string stdev = Math.Round(GetStandardDeviation(abc), 2).ToString();
                       textBox9.Text = stdev.ToString();
                       textBox14.Text = avg.ToString();
                       Log("Std Dev :\t  N " + (vProgress + 1).ToString() + "\tHFR" + abc[vProgress].ToString() +  "\t  Avg " + avg.ToString() + "\t  StdDev " + stdev.ToString());
                       string strLogText = "Std Dev" + "\t  " + abc[vProgress].ToString() + "\t  " + avg.ToString() + "\t" + (vProgress + 1).ToString() + "\t" + stdev.ToString();
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
                       fileSystemWatcher3.EnableRaisingEvents = false;
                   }
               }
           }        

      
       //Close button
        void ButtonDisable_Click(object sender, EventArgs e)
        {
            closing = 1;
            standby();
          //  gotopos(0);
            if (port == null)
            {
                Application.Exit();
            }
            else
            {
                
                    gotopos(0);
                    port.Close();
                }

                Application.Exit();
            }
     
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            
        }

        int roundto;

        private void Form1_Load(object sender, EventArgs e)
        {

            roundto = (int)numericUpDown1.Value;

        }
        //forward button
        private void button1_Click(object sender, EventArgs e)
        {
            watchforOpenPort();
            if (portopen == 1)
            {
                watchforNegFocus();
                int step = (int)numericUpDown2.Value;
                gotopos(Convert.ToInt16(count - step));

                textBox1.Text = count.ToString();
                textBox4.Text = posMin.ToString();
                textBox3.Text = temp.ToString();
                positionbar();
            }
        }
        //reverse button
        private void button2_Click(object sender, EventArgs e)
        {
            watchforOpenPort();
            if (portopen == 1)
            {
                watchforNegFocus();

                int step = (int)numericUpDown2.Value;
                gotopos(Convert.ToInt16(count + step));

                textBox1.Text = count.ToString();
                textBox4.Text = posMin.ToString();
                textBox3.Text = temp.ToString();
                positionbar();
            }
        }
        
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void LogTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void MainWindow_Load(object sender, EventArgs e)
        {

        }

       

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            
        }
   
      

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {
            
        }
        
        private void label12_Click(object sender, EventArgs e)
        {

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
      
        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            
        }

        void gotopos(Int16 value)
        {
            port.DiscardInBuffer();
            Thread.Sleep(20);
            watchforNegFocus();
            port.Write(value.ToString());
            Thread.Sleep(20);//was 100, then 5, ? too fast so made 20 10-11-11
            port.DiscardOutBuffer();
            Thread.Sleep(20); //was 500 ,then 5 (as above change)

            //goto progress bar
            int diff = Math.Abs(value - count);
            if (closing == 0) // allows for no progress bar during closure while stepper returns to zero
            {
                for (int zz = 0; zz < diff; zz++)
                {
                    if (zz % 8 ==0)
                    {
                        Thread.Sleep(1);
                    }
                    if (vcurveenable != 1)
                    {
                        
                        progressBar1.Maximum = diff;
                        progressBar1.Minimum = 0;
                        progressBar1.Increment(10);
                        progressBar1.Value = zz;
                    }
                    count = count + (value - count);
                    textBox1.Text = count.ToString();//was count
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
 
        //gotopos button
         private void button7_Click(object sender, EventArgs e)
        {
            watchforOpenPort();
            if (portopen == 1)
            {

                fileSystemWatcher2.EnableRaisingEvents = false;
                fileSystemWatcher3.EnableRaisingEvents = false;
                int go2 = (int)numericUpDown6.Value;
                gotopos(Convert.ToInt16(go2));
                Thread.Sleep(20);//was 200
            }
             
         }
        
         private void numericUpDown6_ValueChanged(object sender, EventArgs e)
         {
            
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

/*
        public static long GetFileHFR(string[] filePaths)
        {
            int current;
            int currentHFR;
            long param2;
         //   string[] filePaths = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.bmp");
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
                             // value is not a number
                             string current2 = filename.Substring(first - 1, second - first - 1);
                             //added to convert to int
                             string replace = current2.Replace(".", "");
                             current = Convert.ToInt32(replace);
                             //end addition
                         }
                          param2 = Convert.ToInt32(current);
                          return (param2);
                         //was closestHFR = 
                     //  return  ((int)((param2 + (0.5 * roundto)) / roundto) * roundto);
                     }

                 //    return (param2);  
                   
        }
*/

         void vcurve()
         {
           //  int[] backlashList = new int[10];
             MoveDelay = (int)numericUpDown9.Value;
             Thread.Sleep(MoveDelay);//added to prevent exposure during movement
             vcurveenable = 1;
          
           //  int maxHFR = 1;
                 if (tempon == 1)
                 {
                     fileSystemWatcher1.EnableRaisingEvents = true;
                 }
                 if ((vProgress == 0) & (tempon == 1))
                 {
                     int finegoto = posMin - ((((int)numericUpDown3.Value) / 2) * ((int)numericUpDown8.Value));
                     //fine v-curve goes to N/2 * step size in from the focus position -- V should be centered
                     gotopos(Convert.ToInt16(finegoto));
                   //  gotopos(Convert.ToInt16(posMin - 10));
                 }

                 if (ffenable == 1)
                 {
                     
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
                 
                 if (vProgress == vN)
                 {
                     vDone = 1;
                    // textBox3.Text = vDone.ToString();
                   //  Thread.Sleep(1000);
                 }
                 int roundto = (int)numericUpDown1.Value;             
                 int avg = 0;
                 int current = 0;
                 int closestHFR= 1;
                 long closestMax=1;
                 if (vDone != 1)
                 {
                     
                  //   int[] Maxlist = new int[vN];
                 //    int[] list = new int[vN];
                  //   int[] peat = new int[repeatTotal];
                 //    long[] peatMax = new long[repeatTotal];
                     int[] templist = new int[((vN * repeatTotal) + 1)];
                     string[] filePaths = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.bmp");
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
                             // value is not a number
                             string current2 = filename.Substring(first - 1, second - first - 1);
                             //added to convert to int
                             string replace = current2.Replace(".", "");

                             current = Convert.ToInt32(replace);
                             //end addition
                         }
                         long param = Convert.ToInt32(current);


                         closestHFR = (int)((param + (0.5 * roundto)) / roundto) * roundto;
                         //  long closest2 = Convert.ToInt32(closestHFR);

                         //try getting max value start insert her
                         foreach (string filenameMax in filePaths)
                         {
                             int firstMax = filename.IndexOf(@"_");
                             int secondMax = filename.IndexOf("_", firstMax + 1);
                             string sizeMax = filename.Substring(firstMax + 1, ((secondMax - firstMax) - 1));

                             long paramMax = Convert.ToInt32(sizeMax);

                             long roundtoMax = (roundto * 10);
                             closestMax = (long)((paramMax + (0.5 * roundtoMax)) / roundtoMax) * roundtoMax;
                             //   long closest2Max = Convert.ToInt32(closestMax);
                         }
                         //end addition
//determine maxMax
                     }
                    // list[0] = closestHFR;
                    // listMax[0] = closestMax;
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
                     //  maxMax = closestMax;
                     if (repeatDone == 1)
                     {
                         listMax[vProgress] = avgMax;

                     }
//non-array method - remd since appears duplicate
           //          if (listMax[vProgress] > maxMax)
            //         {
             //            maxMax = listMax[vProgress];
              //           posmaxMax = count;
//                     }         



                     //end addition

//next 5 rems test
                   //  maxHFR = closestHFR;
                  //   if (list[vProgress] > maxHFR)
                 //    {
                  //       maxHFR = list[vProgress];
                 //    }


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

 //determine maxMax - old non-array method
                         /*
                         if (listMax[vProgress] > maxMax)
                         {
                            maxMax = (int)listMax[vProgress];
                            posmaxMax = count;
                         }

                         if (list[vProgress] < min)
                         {
                             min = list[vProgress];
                             posMin = count;
                         }
            */
                         // repeatDone = 0;
                         //reset repeat
                         // repeatProgress = 0;
                         //sum = 0;
                         //end Max addition
                         //check if min HFR

//add test array method  
                        
                      
                     

                         templist[vv] = temp;
                         tempsum = tempsum + templist[vv];
                         tempavg = tempsum / (vv + 1);
                         total = total + 1;
                         textBox1.Text = count.ToString();
                         textBox3.Text = temp.ToString();
                         textBox4.Text = posMin.ToString();
                         //progress current is textbx 7 of total N is textbox6
                         textBox6.Text = vN.ToString();
                         textBox7.Text = (vProgress + 1).ToString();
                         textBox10.Text = tempavg.ToString();



//old log                         Log("V-Curve: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\tPos" + count.ToString() + " \t  HFRavg" + avg.ToString() + "\tMaxavg" + avgMax.ToString() + "\t  minHFR" + min.ToString() + "\t maxMax" + maxMax.ToString() + "\t posMin" + posMin.ToString());
                         Log("V-Curve: N  " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t\tPos" + count.ToString() + " \t  HFRavg" + avg.ToString() + "\tMaxavg" + avgMax.ToString());
//old                         string strLogTextA = "V-Curve: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + count.ToString() + " \t" + avg.ToString() + "\t" + avgMax.ToString() + "\t" + min.ToString() + "\t" + maxMax.ToString() + "\t" + posMin.ToString();
                         string strLogTextA = "V-Curve: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + count.ToString() + " \t" + avg.ToString() + "\t" + avgMax.ToString();
                         string strLogText = "V-curve" + "\t" + temp.ToString() + "\t  " + count.ToString() + "\t" + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + current.ToString() + "\t" + min.ToString() + "\t" + posMin.ToString();
                         string strLogText2 = "TempCal" + "\t " + tempavg.ToString() + "\t " + posMin.ToString();
                         string strLogText3 = "Fine-V: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + count.ToString() + " \t" + avg.ToString() + "\t" + avgMax.ToString() + "\t" + min.ToString() + "\t" + maxMax.ToString() + "\t" + posMin.ToString();
                         positionbar();
                         if ((tempon != 1) & (ffenable != 1) & (repeatDone == 1))
                         {
                             chart1.Series[0].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg));
                         }
                         if ((ffenable == 1) & (tempon == 0) & (repeatDone == 1) & (backlashDetermOn == false))
                         {
                             chart1.Series[1].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg));
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
                                 log.WriteLine("Fine-V" + "\tTemp" + "\t  Pos" + "\tN" + "\tHFR" + "\tminHFR" + "\tposmin");
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
                     



//******consider trying avg of maxMax and minHFR*******
                //        posMin = (posmaxMax + posminHFR) / 2;

// tie braker 
/*                         if (vProgress != 0)
                         {
                             if (min == list[vProgress - 1])
                             {
                                 posMin = posmaxMax;
                             }
                         }
 */
//next 5 re = test

//rem to test array                         
                    //     if (list[vProgress] < min)
                    //     {
                    //         min = list[vProgress];
                    //         posMin = count;
                    //     }
                         
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
                                 // count = count + step;
                                 vProgress++;
                                 vv++;

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
                                 }
                                 if (backlashOUT == true)
                                 {
                                     if (vProgress < (vN-1))
                                     {
                                         gotopos(Convert.ToInt16(count + step));
                                     }
                                     vProgress++;
                                     vv++;
                                 }
                             }

                         }
                     }
                         progressBar1.Maximum = (vN);
                         progressBar1.Minimum = 0;
                         progressBar1.Increment(1);
                         progressBar1.Value = vProgress;
                       //  log.Close();

                         //V=1...skips to here.
                     }
                    



                        
                         if (vDone == 1)
                     {

                             for (int arraycount = 0; arraycount < vProgress; arraycount++)
                             {
                                 if (list[arraycount] < HFRarraymin)
                                 {
                                     HFRarraymin = list[arraycount];
                                     posminHFR = minHFRpos[arraycount];
                                     min = HFRarraymin;
                                     apexHFR = arraycount;
                                 }
                             }

                           //remd 10-21-11  posMin = posminHFR;
                         
                             for (int arraycount = 0; arraycount < vProgress; arraycount++)
                             {
                                 if (listMax[arraycount] > 35000)
                                 {
                                     MaxtooHi = true;
                                 }
                                 if (listMax[arraycount] > maxarrayMax)
                                 {
                                     maxarrayMax = listMax[arraycount];
                                     posmaxMax = (int)maxMaxPos[arraycount];
                                     if ((posmaxMax == posminHFR) & (MaxtooHi = false))
                                     {
                                         posMin = posmaxMax;
                                     }

                                 }
                             }




                                 if ((ffenable == 1) & (radioButton1.Checked)) 
                             {
   //need to fix max slope to set max value exception
                                 //***try new posmin determination using 2 points each side of apex****
                                 //      posmaxMax = (int)GetPositionMax(maxMaxPos[apexHFR - 1], listMax[apexHFR - 1], maxMaxPos[apexHFR - 2], listMax[apexHFR - 2], maxMaxPos[apexHFR + 1], listMax[apexHFR + 1], maxMaxPos[apexHFR + 2], listMax[apexHFR + 2]);
                                 //disable for sim camera (radio button for testing only
                    

                                 //***try new posmin determination using 2 points each side of apex****
//remd to try focusmax method       posminHFR = (int)GetPositionHFR(minHFRpos[apexHFR - 1], list[apexHFR - 1], minHFRpos[apexHFR - 2], list[apexHFR - 2], minHFRpos[apexHFR + 1], list[apexHFR + 1], minHFRpos[apexHFR + 2], list[apexHFR + 2]);
                           
                                     //  posMin = (posmaxMax + posminHFR) / 2;
                             //Unrem if above not working    posMin = posminHFR;

//try multiple point slope determination then uses 2 points up from apex on both sides to calc posMin                        

                                 //need to setup 2 arrays for each HFR and Max array
                                 //test GetSlope
                                 //hfr array     
                                    
                                     int downN = apexHFR -1;
                                 double[] HFRdown = new double[downN];
                                 double[] HFRup = new double[vN - apexHFR-2];
                                 double[] HFRposdwn = new double[downN];
                                 double[] HFRposup = new double[vN - apexHFR-2];
                                 for (int dwn = 0; dwn < downN; dwn++)
                                 {
                                     HFRdown[dwn] = list[dwn];
                                     HFRposdwn[dwn] = minHFRpos[dwn];
                                 }
                                 for (int up = 0 ; up < (vN - apexHFR-2); up++)
                                 {
                                     HFRup[up] = list[apexHFR +1 + up];
                                     HFRposup[up] = minHFRpos[apexHFR +1 + up];
                                 }
                                 slopeHFRdwn = GetSlope(HFRdown, HFRposdwn);
                                 textBox3.Text = slopeHFRdwn.ToString();
                                 slopeHFRup = GetSlope(HFRup, HFRposup);
                                 textBox10.Text = slopeHFRup.ToString();
                                 XintDWN = minHFRpos[apexHFR - 5] - list[apexHFR - 5] / slopeHFRdwn;
                                 XintUP = minHFRpos[apexHFR + 5] - list[apexHFR + 5] / slopeHFRup;
                                 PID = XintDWN - XintUP;




                                 /*   ***this for multiple v-curves***
                                     AVGslopeDWN[MultiVcurveProgress] = slopeHFRdwn;
                                     AVGslopeUP[MultiVcurveProgress] = slopeHFRup;
                                     AVGintersectPOS[MultiVcurveProgress] = posMin;
                                     AVGHFR[MultiVcurveProgress] = HFRarraymin;
                                    
                                     AVGxintUP[MultiVcurveProgress] = minHFRpos[apexHFR + 5] - list[apexHFR + 5] / slopeHFRup;
                                     AVGxintDWN[MultiVcurveProgress] = minHFRpos[apexHFR - 5] - list[apexHFR - 5] / slopeHFRdwn;
                                     AVGPID[MultiVcurveProgress] = AVGxintDWN[MultiVcurveProgress] - AVGxintUP[MultiVcurveProgress];
                                     double averageUP = AVGslopeUP.Average();
                                     double averageDWN = AVGslopeDWN.Average(); 
                          */
                                     

                                 




/*
                                 //max array
                                 //   int downN = apexHFR;
                                 double[] maxdown = new double[downN];
                                 double[] maxup = new double[vN - apexHFR];
                                 double[] maxposdwn = new double[downN];
                                 double[] maxposup = new double[vN - apexHFR];
                                 for (int dwn = 0; dwn < apexHFR; dwn++)
                                 {
                                     maxdown[dwn] = listMax[dwn];
                                     maxposdwn[dwn] = maxMaxPos[dwn];
                                 }
                                 for (int up = apexHFR+1; up < vN; up++)
                                 {
                                     maxup[up] = listMax[up];
                                     maxposup[up] = maxMaxPos[up];
                                 }
                                 double slopemaxdwn = GetSlope(maxdown, maxposdwn);
                                 textBox12.Text = slopemaxdwn.ToString();
                                 double slopemaxup = GetSlope(maxup, maxposup);
                                 textBox13.Text = slopemaxup.ToString();
*/


                                 posMin = (int)GetIntersectPos((double)minHFRpos[apexHFR - 5], (double)minHFRpos[apexHFR + 5], (double)list[apexHFR - 5], (double)list[apexHFR + 5], slopeHFRdwn, slopeHFRup);
                                 textBox4.Text = posMin.ToString();
                                // BestHFR = (slopeHFRup * (posMin - count) +avg);
                               //  textBox15.Text = BestHFR.ToString();
                                 Log("Slope: N  " + (vProgress + 1).ToString() + "\t\tslopeUP" + slopeHFRup.ToString() + " \t  SlopeDWN" + slopeHFRdwn.ToString() + "\tBestPOS" + posminHFR.ToString() + "\tPID" + PID.ToString());

                                  }
                                 else
                             {
                                 posMin = posminHFR;
                             }
                         
                         
                         
                         textBox3.Text = posminHFR.ToString();
                         textBox10.Text = min.ToString();
                         textBox12.Text = maxarrayMax.ToString();
                         textBox13.Text = posmaxMax.ToString();
                         textBox4.Text = posMin.ToString();
                         textBox1.Text = current.ToString();

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
                         log.WriteLine("Posmin: " + posMin.ToString() + "\tminHFR " + min.ToString() + "\tminHFRpos " + posminHFR.ToString() + "\tmaxMAx " + maxarrayMax.ToString() + "\tmaxMaxPos " + posmaxMax.ToString());
                         log.Close();

                         if ((vDone == 1) & (ffenable == 1))
                         {
                             gotoFocus();
                         }

                         if ((vDone == 1) & (tempon == 0) & (backlashDetermOn == false))
                         {
                             fileSystemWatcher1.EnableRaisingEvents = false;
                             gotopos(Convert.ToInt16(posMin));
                             standby();
                         }
                         if ((tempon == 1) & (vDone != 1))
                         {
                             repeatDone = 0;
                             repeatTotal = 0;
                             fileSystemWatcher1.EnableRaisingEvents = true;
                         }

                         if ((tempon == 1) & (vDone == 1))
                         {
                             tempsum = 0;
                             vDone = 0;
                             vv = 0;
                             min = 500;
                             vProgress = 0;
                             fileSystemWatcher1.EnableRaisingEvents = false;
                             chart1.Series[2].Points.AddXY(Convert.ToDouble(posMin), Convert.ToDouble(tempavg));
                             gotopos(Convert.ToInt16(posMin));
                             tempcal();
                             repeatProgress = 0;
                             vcurveenable = 0;
                         }
                         vcurveenable = 0;//added 9-12
                     }
                    // textBox3.Text = vDone.ToString();
                     
                        
                     
                     if ((backlashDetermOn == true) & (vDone == 1))
                     {
                         
                       //  vcurveenable = 0;
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
                           //  vcurve();
                         }

                         if ((backlashCount == 2) || (backlashCount == 4) || (backlashCount == 6) || (backlashCount == 8) || (backlashCount == 10))
                         {

                             backlash = Math.Abs(backlashPosIN - backlashPosOUT);
                             // backlashList[backlashCount] = backlash;
                             textBox8.Text = backlash.ToString();
                             chart1.Series[2].Points.AddXY(Convert.ToDouble(backlashCount), Convert.ToDouble(backlash));
                             backlashSum = backlash + backlashSum;
                             backlashAvg = backlashSum / (backlashCount / 2);
                             Log("Backlash: N " + backlashCount.ToString() + "\tPosOUT " + backlashPosOUT.ToString() + "\tPosIN " + backlashPosIN.ToString() + "\tBacklash " + backlash.ToString() + "\tAvg " + backlashAvg.ToString());
                             
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
                         Array.Clear(listMax, 0, 100);
                         Array.Clear(list, 0, 100);
                         HFRarraymin = 999;
                         maxarrayMax = 1;
                     }


                              
         }





         private void textBox8_TextChanged(object sender, EventArgs e)
         {

         }

         private void toolTip3_Popup(object sender, PopupEventArgs e)
         {

         }

         private void chart1_Click(object sender, EventArgs e)
         {
             chart1.Series[1].Points.AddXY(Convert.ToDouble(min), Convert.ToDouble(count));
       
         }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        public void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            string path2 = folderBrowserDialog1.SelectedPath.ToString();         
            textBox11.Text = path2.ToString();
        }
//max travel set to 1360 default for tak TSA-120
        private void numericUpDown3_ValueChanged_1(object sender, EventArgs e)
        {

        }
//connect button
        private void button8_Click(object sender, EventArgs e)
        {
            //rem to port.open() if using auto open
            /*
                        if (port != null)
                        {
                            port.Close();
                            port.Dispose();
                        }
                        port = new SerialPort(comboBox1.SelectedItem.ToString(), 115200, Parity.None, 8, StopBits.One);
                        port.Open();
            */
            //un-rem if using auto port oen
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
                Thread.Sleep(200);//was 500
                //    port.DiscardOutBuffer();

                //    Thread.Sleep(100);
                str2 = port.ReadExisting();
                Thread.Sleep(500);//was 500

                //  int i = str2.IndexOf('P');
                //   string cutstr2i = str2.Substring(i + 10);

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
                        
                        textBox5.Text = travel.ToString();

                    }
                }
                if (syncfailure2 != 1)
                {

                    travel = numb2;
                    //  travel = (int)numericUpDown3.Value;
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
                    // numbers = travel.ToString();
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
                    //  travel = (int)numericUpDown3.Value;
                    textBox5.Text = waittime.ToString();
                }
            }
            else
            {
                DialogResult result2 = MessageBox.Show("Already Connected", "Arduino scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            connect = 1;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
//added to sync for manual changes

        private void button9_Click(object sender, EventArgs e)
        {
            sync();
            count = syncval;
        }

            void sync()
            {
            //travel = (int)numericUpDown3.Value;
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
           // port.DiscardOutBuffer();

          //  Thread.Sleep(100);
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
                // Write to the file:

                log.WriteLine(strLogText4);
                log.Close();
                watchforNegFocus();
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String strVersion = System.Reflection.Assembly.GetCallingAssembly().GetName().Version.ToString(); 

            MessageBox.Show("           Arduino Version" + strVersion.ToString()+  "\r" + "           Kevin Sipprell MD\r           www.scopefocus.info\r      ", "scopefocus for nebulosity",
            MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged_2(object sender, EventArgs e)
        {
         
        }
        //Fine V-curve, 
        private void button3_Click(object sender, EventArgs e)
        {
             watchforOpenPort();
            if (portopen == 1)
            {
                backlashDetermOn = false;
                port.DiscardInBuffer();
                watchforNegFocus();
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                chart1.Series[2].Points.Clear();
                int finegoto = posMin - ((((int)numericUpDown3.Value) / 2) * ((int)numericUpDown8.Value));
                //fine v-curve goes to N/2 * step size in from the focus position -- V should be centered
                gotopos(Convert.ToInt16(finegoto));
                ffenable = 1;
                repeatProgress = 0;
                repeatDone = 0;
                vProgress = 0;
                vDone = 0;
                repeatTotal = 0;
                vN = 0;
                sum = 0;
                min = 500;
                fileSystemWatcher2.EnableRaisingEvents = true;
                fileSystemWatcher3.EnableRaisingEvents = false;
                fileSystemWatcher1.EnableRaisingEvents = false;
                tempon = 0;
                sumMax = 0;
                avgMax = 0;
                maxMax = 0;
                vv = 0;
                Array.Clear(listMax, 0, 100);
              //  vcurve();
            }
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
            
      //      Thread.Sleep(100);
                ffenable = 1;
                tempon = 1;
                vcurve();
                
           //     ffenable = 1;

        }
//tempcal button
private void button5_Click(object sender, EventArgs e)
{
    chart1.Series[0].Points.Clear();
    chart1.Series[1].Points.Clear();
    chart1.Series[2].Points.Clear();
    fileSystemWatcher2.EnableRaisingEvents = false;
    fileSystemWatcher3.EnableRaisingEvents = false;
    fileSystemWatcher1.EnableRaisingEvents = true;
    //  min = 1;
    sum = 0;
    avg = 0;
    vDone = 0;
    vProgress = 0;
    tempon = 1;
    tempsum = 0;
    vv = 0;
    templog = 0;
  //  tempcal();
}

private void button4_Click_1(object sender, EventArgs e)
{
    standby();
}

private void numericUpDown5_ValueChanged_1(object sender, EventArgs e)
{

}

private void textBox11_TextChanged(object sender, EventArgs e)
{

}

private void toolTip3_Popup_1(object sender, PopupEventArgs e)
{

}

private void toolTip1_Popup(object sender, PopupEventArgs e)
{

}

private void label8_Click_1(object sender, EventArgs e)
{

}
//backlash determine button
private void button10_Click(object sender, EventArgs e)
{
    backlash = 0;
    backlashSum = 0;
    backlashCount = 0;
    chart1.Series[0].Points.Clear();
    chart1.Series[1].Points.Clear();
    chart1.Series[2].Points.Clear();
    fileSystemWatcher2.EnableRaisingEvents = false;
    fileSystemWatcher3.EnableRaisingEvents = false;
    fileSystemWatcher1.EnableRaisingEvents = true;
    //  min = 1;
    sum = 0;
    avg = 0;
    vDone = 0;
    vProgress = 0;
    tempon = 0;
    tempsum = 0;
    vv = 0;
    templog = 0;
    backlashDetermOn = true;
    ffenable = 1;
         int finegoto = posMin - ((((int)numericUpDown3.Value) / 2) * ((int)numericUpDown8.Value));
        //fine v-curve goes to N/2 * step size in from the focus position -- V should be centered
        gotopos(Convert.ToInt16(finegoto));
        backlashOUT = true;//identifies current v curve direction is going out(reverse)
    
   

}

private void textBox3_TextChanged(object sender, EventArgs e)
{

}

private void label4_Click(object sender, EventArgs e)
{

}


        public static int GetBestHFRPos(int Xa, int Ya, int Yb, int ma)
        {
            return (Xa + (Yb - Ya)/ma);
        }


        public static double GetIntersectPos(double upX, double downX, double upY, double downY, double upSlope, double downSlope)
        {

            return (((upSlope * upX) - (downSlope * downX) + downY - upY) / (upSlope - downSlope));
            //return ((Y2 - Y1 + m2 * X2 + m1 * X1) / (m2 - m1));
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

public void Sim()
{
    string path = textBox11.Text.ToString();
    string fullpath = path + @"\log.txt";
    string[] filePaths = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.bmp");
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



}

private void radioButton1_CheckedChanged(object sender, EventArgs e)
{

}

private void button11_Click(object sender, EventArgs e)
{
    
}

public void gotoFocus()
{
    port.DiscardInBuffer();
    port.DiscardOutBuffer();
    fileSystemWatcher2.EnableRaisingEvents = false;
    fileSystemWatcher3.EnableRaisingEvents = true;
    fileSystemWatcher1.EnableRaisingEvents = false;


    posMin = 0;
    min = 1;
    sum = 0;
    avg = 0;
    vDone = 0;
    vProgress = 0;

    if (vProgress == 0)
    {
        gotopos(Convert.ToInt16(minHFRpos[apexHFR + 5]));//not converting to correct int  and ?  not stopping 
    }
    textBox14.Text = avg.ToString();
    

}


    }
}