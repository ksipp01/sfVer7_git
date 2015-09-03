//Arduino version


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

            fileSystemWatcher2.EnableRaisingEvents = false;
            fileSystemWatcher3.EnableRaisingEvents = false;
            DialogResult result = folderBrowserDialog1.ShowDialog();
            string path2 = folderBrowserDialog1.SelectedPath.ToString();         
            textBox11.Text = path2.ToString();

            fileSystemWatcher2.Path = folderBrowserDialog1.SelectedPath;
            fileSystemWatcher3.Path = folderBrowserDialog1.SelectedPath;

           
        }
        
        //abitrarily set min = 999 to start with
        int total = 0;
        int count = 0;
        int min = 999;
        int posMin = 0;
        int a = 0;
        int r = 0;
        int b = 0;
        int v = 0;
        int c = 0;
        int d = 0;
        int sum = 0;
        int avg = 0;
  //      int go = 0;
        int ffenable = 0;
        int repeat = 0;
        int n = 0;
        int[] list = new int[10];
        double[] abc = new double[10];
        int tempon = 0;
        int tempsum = 0;
        int tempavg = 0;
        int vv = 0;
        int relpos = 1;
        int travel = 10000;
        int connect = 0;
        int temp = 100;
        int test = 0;
        int portopen = 0;
        int vcurveenable = 0;
        int waittime = 1; // this needs to be the same as EEPROM wiat time on arduino
        int closing = 0;
        string valuesend;
        int valueconfirm;
        int syncval;
  //      int close;
  //      int current;
//THis will auto open com11 used only for dell studio 15 laptop, top left USB port
       void PortOpen()
        {
            if (port != null)
            {
                port.Close();
                port.Dispose();
            }
            port = new SerialPort("com11", 9600, Parity.None, 8, StopBits.One);
            port.Open();
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
            LogTextBox.Text += DateTime.Now.ToLongTimeString() + "\t" + text;
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
                vcurveenable = 1;
                a = 0;
                r = 0;
                b = 0;
                v = 0;
                c = 0;
                d = 0;
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
                fileSystemWatcher2.EnableRaisingEvents = true;
                fileSystemWatcher3.EnableRaisingEvents = false;
               
            }
         }

        private void fileSystemWatcher2_Changed(object sender, FileSystemEventArgs e)
        {
           vcurve();
       
        }      
 

    


//std dev and avg
           void Button2000_Click(object sender, EventArgs e)
           {
               fileSystemWatcher2.EnableRaisingEvents = false;
               fileSystemWatcher3.EnableRaisingEvents = true;



               posMin = 0;
               min = 1;
               sum = 0;
               avg = 0;
               v = 0;
               b = 0;
           }
           private void fileSystemWatcher3_Changed(object sender, FileSystemEventArgs e)
           {
               watchforOpenPort();
               if (portopen == 1)
               {
                   int nn = 10;
                   if (b < 10)
                   {

                       int[] list = new int[nn];



                       //  string[] filePaths = Directory.GetFiles(@"c:\test2\", "*.bmp");
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
                               // Do your code, the value is number
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



                           list[b] = current;
                           abc[b] = current;
                           sum = sum + list[b];
                           avg = sum / nn;
                       }
                       string stdev = Math.Round(GetStandardDeviation(abc), 2).ToString();
                       textBox9.Text = stdev.ToString();

                       Log("Std Dev :\tHFR" + abc[b].ToString() + "\t  avg " + avg.ToString() + "\t  N " + (b + 1).ToString() + "\t  StdDev " + stdev.ToString());
                       string strLogText = "Std Dev" + "\t  " + abc[b].ToString() + "\t  " + avg.ToString() + "\t" + (b + 1).ToString() + "\t" + stdev.ToString();
                       string path = folderBrowserDialog1.ToString();

                       // Create a writer and open the file:
                       StreamWriter log;

                       if (!File.Exists(path))
                       {
                           log = new StreamWriter(path);
                       }
                       else
                       {
                           log = File.AppendText(path);
                       }
                       // Write to the file:
                       if (b == 0)
                       {
                           log.WriteLine(DateTime.Now);
                           log.WriteLine("Type" + "\tHFR" + "\tAvg" + "\tN" + "\tStdDev");
                       }
                       log.WriteLine(strLogText);
                       log.Close();
                       b++;
                   }
                   if (b == 10)
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
        

  
 /*
// try adding 'subroutine'  Originally forward
        void forward()
        {
            port.Write("F");
            Thread.Sleep(10);
            port.DiscardOutBuffer();
        }
//reverse 'subroutine'
        void reverse()
        { 
            port.Write("R");
            Thread.Sleep(10);
            port.DiscardOutBuffer();
        }
        
*/
        
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            //int roundto = (int)numericUpDown1.Value;
            
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
                //       int temp = Convert.ToInt16(servos[0].position);
                Log("Manual Fwd - total " + total.ToString() + "   count " + count.ToString() + "   temp " + temp.ToString() + "     relpos" + relpos.ToString());
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
                //          int temp = Convert.ToInt16(servos[0].position);
                Log("Manual Rev - total " + total.ToString() + "   count " + count.ToString() + "   temp " + temp.ToString() + "     relpos" + relpos.ToString());
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
            // positionbar();

            valueconfirm = value;
            watchforNegFocus();
            string valuestring = value.ToString();
            string start = "G";
            string end = "H";
            valuesend = start + valuestring + end;

            port.Write(valuesend);
            Thread.Sleep(20);
            port.DiscardOutBuffer();
            Thread.Sleep(20);
            port.DiscardInBuffer();
            Thread.Sleep(500);
            confirm();

           

            //goto progress bar
            int diff = Math.Abs(value - count);
            if (closing == 0) // allows for no progress bar during closure while stepper returns to zero
            {
                for (int zz = 0; zz < diff; zz++)
                {
                    Thread.Sleep(waittime * (20 / 8));
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
                }
                progressBar1.Value = 0;
                positionbar();
            }

        }
        void confirm()
    {
        sync();
            if (syncval != count)
            {
              gotopos(Convert.ToInt16(valueconfirm));
             }
         
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
                Thread.Sleep(500);
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
        


         void vcurve()
         {
             test++;
             int max = 1;
                
                
                
                     repeat = (int)numericUpDown4.Value;
                 
                 if (a == (repeat - 1))
                 {
                     r = 1;
                 }
                 c = repeat;

                
                     n = (int)numericUpDown5.Value;
                 
                 if (b == n)
                 {
                     v = 1;
                 }
                 else
                 {
                     d = n;
                 }
                
                 
                 int avg = 0;
                 int current = 0;
                 int closest = 1;
                 int crvsz = (int)numericUpDown5.Value;
                 if (v != 1)
                 {
                     int[] list = new int[d];
                     int[] peat = new int[c];
                     int[] templist = new int[((d*c)+1)];
                 //    string[] filePaths = Directory.GetFiles(@"c:\test2\", "*.bmp");
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
                             // Do your code, the value is number
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
                         long param = Convert.ToInt32(current);
                         //   int two = param / 1;
                         //   short param2 = Convert.ToInt16(two);
                         int roundto = (int)numericUpDown1.Value;
                         closest = (int)((param + (0.5 * roundto)) / roundto) * roundto;
                         long closest2 = Convert.ToInt32(closest);

                     }
                     max = closest;
                     if (list[b] > max)
                     {
                         max = list[b];
                     }
                  //   if (a == 0)
                  //   {
                  //       min = closest;
                 //    }
                     if (a != 0)
                      {
                         peat[a] = closest;
                         sum = sum + peat[a];

                         //   if (r==1)
                         //   if (c == n)

                         avg = sum / (a+1);
                     }
                     if (a ==0)
                     {
                         peat[a] = closest;
                         avg = closest;
                         sum = closest;
                     }

                     //r=1 when repeat done
                     if (r == 1)
                     {
                         list[b] = avg;
                         
                     }
                     // list[b] = current;
                    
                  //   int temp = Convert.ToInt16(servos[0].position);
                     templist[vv] = temp;
                     tempsum = tempsum + templist[vv];
                     tempavg = tempsum / (vv + 1);
                     
                   //  if (min == 999)
                    // {
                     //    min = closest;
                     //    posMin = count;
                    // }
                   
                     total = total + 1;
                     textBox1.Text = count.ToString();
                     textBox3.Text = temp.ToString();
                     textBox4.Text = posMin.ToString();

                     textBox6.Text = d.ToString();
                     textBox7.Text = (b + 1).ToString();
 
                     textBox10.Text = tempavg.ToString();
                     Log("V-Curve:\tPos" + count.ToString() + "\t  HFRavg" + avg.ToString() + "\tHFR" + current.ToString() + "\tminHFR" + min.ToString() + "\tposMin" + posMin.ToString());
                     string strLogText = "V-curve" + "\t" + temp.ToString() + "\t  " + count.ToString() + "\t" + avg.ToString() + "\t" + current.ToString() + "\t" + min.ToString() + "\t" + posMin.ToString();
                     string strLogText2 = "TempCal" + "\t " + tempavg.ToString() + "\t " +  posMin.ToString();
                     string strLogText3 = "Fine-V " + "\t" + temp.ToString() + "\t  " + count.ToString() + "\t" + avg.ToString() + "\t" + current.ToString() + "\t" + min.ToString() + "\t" + posMin.ToString();
                     positionbar();
                     if ((tempon != 1) & (ffenable !=1) & (r==1))
                     {
                         chart1.Series[0].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg));
                     }
                    
                     string path = folderBrowserDialog1.ToString(); 
                   
                     // Create a writer and open the file:
                     StreamWriter log;

                     if (!File.Exists(path))
                     {
                         log = new StreamWriter(path);
                     }
                     else
                     {
                         log = File.AppendText(path);
                     }
                     // Write to the file:
                   //  if (ffenable == 1)
                 //    {
                  //       log.WriteLine(DateTime.Now);
                 //        log.WriteLine("Fine-V" + "\tTemp" + "\tPos" + "\tN" + "\tHFR" + "\tminHFR" + "\tposmin" + "\tTavg");
                //     }
                     if ((total < 2) & (tempon ==0)) 
                     {
                         log.WriteLine(DateTime.Now);
                         log.WriteLine("Type" + "\tTemp" + "\tPos" + "\tN" + "\tHFR" + "\tminHFR" + "\tposmin");
                     }
                    
                     
                     if ((tempon ==0) & (ffenable != 1))
                     {
                         log.WriteLine(strLogText);
                     }
                     
                     log.Close();
                     //             stack = closest;
                     // Thread.Sleep(3000);
                     
                     if (r != 1)
                     {
                         a++;
                         vv++;
                     }
                     if (r == 1)
                     {
                         //check if min HFR
                         if (list[b] < min)
                         {
                             min = list[b];
                             posMin = count;
                         }
                         r = 0;
                         //reset repeat
                         a = 0;
                         sum = 0;
                         // tempsum = 0;
                         if (b < d)
                         {
                             int step = (int)numericUpDown2.Value;
                             if (ffenable == 1)
                             {
                                 //defines Fine-V step size
                                 step = 1;
                             }

                             
                             //rem below to use gotopos instead
                         /*    for (int i = 0; i < step; i++)
                             {
                                 //move focus
                                 
                                 Thread.Sleep(100);
                                 count++;
                                 reverse();
                            
                             }
                             */
                             //change from above to gotopos method
                             gotopos(Convert.ToInt16(count + step));
                            // count = count + step;
                             b++;
                             vv++;

                         }
                     }
                     
                    
                     progressBar1.Maximum = (d);
                     progressBar1.Minimum = 0;
                     progressBar1.Increment(1);
                     progressBar1.Value = b;
                     log.Close();
                   
                //V=1...skips to here.
                 }
                 if ((v == 1) & (tempon == 0))
                 {
                     progressBar1.Value = 0;
                     gotopos(Convert.ToInt16(posMin));
                     standby();
                 }



                 vcurveenable = 0;    //added 9-12         
         }

         void testVcurve()
         {
             Log("V-Curve:\tTemp ");
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
                Thread.Sleep(500);
                //    port.DiscardOutBuffer();

                //    Thread.Sleep(100);
                str2 = port.ReadExisting();
                Thread.Sleep(50);

                //  int i = str2.IndexOf('P');
                //   string cutstr2i = str2.Substring(i + 10);

                port.DiscardInBuffer();

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
                        travel = 10000;
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
                //   Thread.Sleep(500);

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
                Thread.Sleep(500);
                // port.DiscardOutBuffer();

                //   Thread.Sleep(100);
                str3 = port.ReadExisting();
                Thread.Sleep(50);
                //  int ii = str3.IndexOf('P');
                //    string cutstr3 = str3.Substring(ii + 10);
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
                if ((numb3 < 1) || (numb3 > 10))
                {
                    DialogResult result;
                    syncfailure3 = 1;

                    Log("Data Sync Failed -- Set to Default");
                    result = MessageBox.Show("Data Sync Out of Range -- Set to Default ", "Data Sync Failure",
                                MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (result == DialogResult.OK) 
                    {
                        waittime = 1;
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
//hit button after changes then 1 rot enc click then hit again and should sync
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
            port.Write("P");
           Thread.Sleep(100);
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
                Log("Synced to position " + count.ToString());
                textBox1.Text = numb.ToString();
                string strLogText4 = "Synced to position " + count.ToString();
                positionbar();
                string path = folderBrowserDialog1.ToString();
                // Create a writer and open the file:
                StreamWriter log;

                if (!File.Exists(path))
                {
                    log = new StreamWriter(path);
                }
                else
                {
                    log = File.AppendText(path);
                }
                // Write to the file:

                log.WriteLine(strLogText4);
                log.Close();
                watchforNegFocus();
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("           Arduino Version 2.2.1\r           Kevin Sipprell MD\r           www.scopefocus.info\r      ", "scopefocus for nebulosity",
            MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged_2(object sender, EventArgs e)
        {
         
        }

        



    }
}