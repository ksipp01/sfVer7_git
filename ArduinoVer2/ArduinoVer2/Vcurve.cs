using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pololu.Usc.ScopeFocus
{


    
    class Vcurve
    {

       
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


        //  public static int posMin;
        public static int backlashCount;
        public static bool backlashOUT;
     //   MainWindow m = new MainWindow();
        public void vcurve()
        {
            int vLostStar = 0;
            float backlashAvg = 0;


            // int posMin;


            //delete these from above 
            int roundto;
            int vcurveenable;
            int tempon = 0;
            int vProgress = 0;
            int ffenable = 0;
            int repeatTotal;
            int repeatProgress = 0;
            int repeatDone = 0;
            int vN;
            int vDone = 0;
            long avgMax = 0;
            long sumMax = 0;
            int count = 0;
            int maxMax;
            int sum = 0;
            bool backlashDetermOn = true;
            //  int vv = 0;


            int backlashPosIN = 0;
            int backlashPosOUT = 0;
            int backlash;
            //  int tempon = 0;
            int tempsum = 0;
            int tempavg = 0;
            int vv = 0;//total aquisitions for a given v-curve, N * repeat
            int travel;//maxstep in driver
                       //   int connect = 0;
                       //  int connect2 = 0;
            int temp = 100;
            int total = 0;
            int templog = 0;
            int min = 0;

            int posmaxMax;
            int[] list = null;
            double[] listMax = null;
            double[] abc = null;
            int[] peat = null;
            double[] peatMax = null;
            int[] minHFRpos = null;
            double[] maxMaxPos = null;
            //   int posmin;
            //  Data d = new Data();
            //  try
            //  {
            int arraysize2;

          int arraycountright = 0;
        int arraycountleft = 0;
        int enteredMaxHFR;


            int backlashSum;
         int enteredMinHFR;




        double maxarrayMax = 1;
            int backlashN = 10;
            if (MainWindow.CurrentFilter == 1)
                MainWindow.Filtertext = m.comboBox2.Text;
            if (MainWindow.CurrentFilter == 2)
                MainWindow.Filtertext = m.comboBox3.Text;
            if (MainWindow.CurrentFilter == 3)
                MainWindow.Filtertext = m.comboBox4.Text;
            if (MainWindow.CurrentFilter == 4)
                MainWindow.Filtertext = m.comboBox5.Text;
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
            int MoveDelay = (int)m.numericUpDown9.Value;

            vcurveenable = 1;
            if (tempon == 1)
            {
                m.fileSystemWatcher1.EnableRaisingEvents = true;
            }
            if ((vProgress == 0) & (tempon == 1))
            {
                int finegoto = MainWindow.posMin - ((((int)m.numericUpDown3.Value) / 2) * ((int)m.numericUpDown8.Value));
                //fine v-curve goes to N/2 * step size in from the focus position -- V should be centered
                m.gotopos(Convert.ToInt32(finegoto));
            }

            if (ffenable == 1)
            {
                m.textBox17.Text = (GlobalVariables.FineRepeatDone + 1).ToString();//exposure repeat
                repeatTotal = (int)m.numericUpDown7.Value;

            }
            else
            {
                repeatTotal = (int)m.numericUpDown4.Value;
            }
            if (repeatProgress == (repeatTotal - 1))
            {
                repeatDone = 1;
            }
            if (ffenable == 1)
            {
                //define Fine-V total sample number (should be twice goto position[line 914 temp cal or 187 fine-v]/step size [line 1135])
                vN = (int)m.numericUpDown3.Value;
            }

            else
            {
                vN = (int)m.numericUpDown5.Value;
            }
            if (vProgress < vN)
            {
                m.progressBar1.Maximum = (vN);
                m.progressBar1.Minimum = 0;
                m.progressBar1.Increment(1);
                m.progressBar1.Value = vProgress + 1;
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
                if (m.checkBox22.Checked == true)
                {
                    //begin test metriHFR
                    m.metricpath = Directory.GetFiles(GlobalVariables.Path2.ToString(), "metric*.fit");
                    try
                    {
                        testMetricHFR = GetMetric(m.metricpath, roundto);
                        m.textBox25.Text = testMetricHFR.ToString();
                        //  Log("MetricHFR" + testMetricHFR.ToString());
                    }
                    catch (Exception e)
                    {
                        m.standby();
                        MessageBox.Show("Error parsing filename, ensure there are not any files w/ aaaa_bbbb_cccc_dddd_eeee.fit structure in the capture directory", "scoprefocus");
                        m.Log("GetHFR error" + e.ToString());
                        m.Send("GetHFR error" + e.ToString());
                        m.FileLog("GetHFR error" + e.ToString());

                        return;


                    }

                    closestHFR = testMetricHFR;

                }
                //  int[] templist = new int[((vN * repeatTotal) + 1)]; //7-25-14 chesnted to below
                int[] templist = new int[((vN * repeatTotal) + 0)];
                string[] filePaths = Directory.GetFiles(GlobalVariables.Path2.ToString(), "*.bmp");
                if (m.checkBox22.Checked == false)//added for metrichfr
                {
                    closestHFR = GetFileHFR(filePaths, roundto);

                }
                if (m.checkBox22.Checked == false)//add for metricHFR
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
                if (m.checkBox22.Checked == false)//add for metriHFR
                {
                    maxMaxPos[vProgress] = count;
                }
                if (m.checkBox22.Checked == true)//added for metricHFR
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

                if ((ffenable == 1) & (tempon == 0) & (repeatDone == 1) & (backlashDetermOn == false)) //check only for fine V-curve
                {
                    if ((vProgress > 2) && (vProgress < (vN - 2)))//don't use first/last stars
                    {
                        if (vLostStar > 0)//if lost next ones will be high as well
                        {
                            if (list[vProgress] > 600)
                            {
                                vLostStar++;
                                m.Log("HFR " + list[vProgress].ToString() + " lost star?");
                            }
                        }
                        if (vLostStar == 3)
                        {
                            vLostStar = 0;
                            m.fileSystemWatcher2.EnableRaisingEvents = false;
                            m.fileSystemWatcher5.EnableRaisingEvents = false;
                            m.Send("V-curve aborted, lost star?");
                            MessageBox.Show("Possible Lost Star, V-curve aborted");
                        }

                        if ((((double)list[vProgress] / (double)list[vProgress - 1]) > 1.5) || (((double)list[vProgress] / (double)list[vProgress - 1]) < .5)) //find outlier
                        {
                            vLostStar++;
                            m.Log("HFR " + list[vProgress].ToString() + " lost star?");

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
                    m.textBox1.Text = count.ToString();
                    m.textBox4.Text = MainWindow.posMin.ToString();
                    //progress current is textbx 7 of total N is textbox6
                    m.textBox6.Text = vN.ToString();
                    m.textBox7.Text = (vProgress + 1).ToString();
                    m.Log("V-Curve: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t\tPos" + count.ToString() + " \t  HFRavg" + avg.ToString() + "\tMaxavg" + avgMax.ToString() + "\t" + MainWindow.Filtertext);
                    string strLogTextA = "V-Curve: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + count.ToString() + " \t" + avg.ToString() + "\t" + avgMax.ToString();
                    string strLogText = "V-curve" + "\t" + temp.ToString() + "\t" + count.ToString() + "\t" + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + current.ToString() + "\t" + min.ToString() + "\t" + MainWindow.posMin.ToString();
                    string strLogText2 = "TempCal" + "\t " + tempavg.ToString() + "\t " + MainWindow.posMin.ToString();
                    string strLogText3 = "Fine-V: N " + (vProgress + 1).ToString() + "-" + (repeatProgress + 1).ToString() + "\t" + count.ToString() + " \t" + avg.ToString() + "\t" + avgMax.ToString();
                    m.positionbar();

                    if ((tempon != 1) & (ffenable != 1) & (repeatDone == 1))
                    {
                        m.chart1.Series[0].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg)); //chart course v curve
                    }
                    if ((ffenable == 1) & (tempon == 0) & (repeatDone == 1) & (backlashDetermOn == false))
                    {
                        if ((avg < enteredMinHFR) || (avg > enteredMaxHFR) & (m.radioButton1.Checked == true))
                        {
                            m.chart1.Series[2].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg)); // chart data not used for calcs < min and > max
                        }
                        else
                        {
                            m.chart1.Series[1].Points.AddXY(Convert.ToDouble(count), Convert.ToDouble(avg));//charte data used for calcs
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
                        if ((ffenable == 1) & (vProgress == 0))
                        {
                            //    log.WriteLine(DateTime.Now);
                            //    log.WriteLine("Fine-V" + "\tTemp" + "\t Pos" + "\tN" + "\tHFR" + "\tminHFR" + "\tposmin");
                            // }
                            m.FileLog2(DateTime.Now.ToString());
                            //   FileLog2("Fine-V" + "\tTemp" + "\t Pos" + "\tN" + "\tHFR" + "\tminHFR" + "\tposmin");
                        }
                        if ((vProgress == 0) & (tempon == 0))
                        {
                            //log.WriteLine(DateTime.Now);
                            //log.WriteLine("Type  N:" + "\t  Pos" + "\tHFRAvg" + "\tMaxAvg");
                            m.FileLog2(DateTime.Now.ToString());
                            m.FileLog2("Type  N:" + "\t  Pos" + "\tHFRAvg" + "\tMaxAvg");
                        }
                        if ((tempon == 1) & (vProgress == (vN - 1)))
                        {
                            if (templog == 0)
                            {
                                //     FileLog2(DateTime.Now.ToString());
                                m.FileLog2("Type" + "\tAvgTemp" + "\tposMin");
                                // log.WriteLine(strLogText2);
                                templog = 1;
                                // log.Close();
                            }

                            m.FileLog2(strLogText2);//was log.WriteLine and next 2 below

                        }
                        if ((tempon == 0) & (ffenable != 1))
                        {
                            m.FileLog2(strLogTextA);
                        }
                        if ((tempon == 0) & (ffenable == 1))
                        {
                            m.FileLog2(strLogText3);
                        }
                    }

                    //    log.Close();
                    repeatDone = 0;
                    //reset repeat
                    repeatProgress = 0;
                    sum = 0;
                    if (vProgress < (vN - 1))
                    {
                        int step = (int)m.numericUpDown2.Value;
                        if (ffenable == 1)
                        {
                            //defines Fine-V step size
                            step = (int)m.numericUpDown8.Value;
                        }
                        if (backlashDetermOn != true)
                        {
                            vProgress++;
                            vv++;
                            m.fileSystemWatcher2.EnableRaisingEvents = false; //****  11-20-14 this is added to allow focuser to complete movement before using the next exposure
                            m.fileSystemWatcher5.EnableRaisingEvents = false;
                            m.gotopos(Convert.ToInt32(count + step));
                            Thread.Sleep(MoveDelay);//helps prevent focus mvmnt during capture
                            m.fileSystemWatcher2.EnableRaisingEvents = true;
                            m.fileSystemWatcher5.EnableRaisingEvents = true;

                        }



                        //  }
                        if (backlashDetermOn == true)
                        {
                            if (backlashOUT == false)
                            {
                                if (vProgress < (vN - 1))
                                {
                                    m.gotopos(Convert.ToInt32(count - step));
                                }
                                vProgress++;
                                vv++;
                                Thread.Sleep(MoveDelay);
                            }
                            if (backlashOUT == true)
                            {
                                if (vProgress < (vN - 1))
                                {
                                    m.gotopos(Convert.ToInt32(count + step));
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
                m.fileSystemWatcher2.EnableRaisingEvents = false;
                m.fileSystemWatcher5.EnableRaisingEvents = false; //added to test metritHFR     

                for (int arraycount = 0; arraycount < vProgress; arraycount++)
                {
                    if (list[arraycount] < _hFRarraymin)
                    {
                        _hFRarraymin = list[arraycount];
                        _posminHFR = minHFRpos[arraycount];
                        min = _hFRarraymin;
                        _apexHFR = arraycount;
                        m.textBox4.Text = MainWindow.posMin.ToString();
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
                            MainWindow.posMin = posmaxMax;

                        }

                    }
                }

                //make sure v-curve is symetric
                //                                                                                   
                if (m.radioButton1.Checked == true)
                {                                //below was vN/2 +4 and -4 changed to +/- vN/4 11-24
                    if (((_apexHFR > (vN / 2 + vN / 4)) || (_apexHFR < (vN / 2 - vN / 4))) & (ffenable == 1) & (backlashDetermOn == false))
                    {
                        m.fileSystemWatcher2.EnableRaisingEvents = false;
                        m.fileSystemWatcher5.EnableRaisingEvents = false; //added to test metricHFR
                        DialogResult result2 = MessageBox.Show("V-Curve Not Symmetric", "Aborted - Repeat Rough V-Curve", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (result2 == DialogResult.OK)
                        {

                            m.standby();
                        }
                    }
                }
                if ((ffenable == 1) & (m.radioButton1.Checked == true)) // save data on
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
                        MainWindow.posMin = _posminHFR;
                        //use point half way up each side to calc intersect position can be used as relative focus position
                        _intersectPos = (int)GetIntersectPos((double)minHFRpos[_apexHFR - vN / 4], (double)minHFRpos[_apexHFR + vN / 4], (double)list[_apexHFR - vN / 4], (double)list[_apexHFR + vN / 4], _slopeHFRdwn, _slopeHFRup);
                        //all vN/2 above were 5

                        m.textBox4.Text = MainWindow.posMin.ToString();
                        m.textBox15.Text = _hFRarraymin.ToString();
                        m.WriteSQLdata();
                        m.FillData();
                        m.Log("Slope: N " + (vProgress + 1).ToString() + "\tslopeUP" + _slopeHFRup.ToString() + " \tSlopeDWN" + _slopeHFRdwn.ToString() + "\tIntersect" + _intersectPos.ToString() + "\tPID" + _pID.ToString() + "\t" + MainWindow.Filtertext);
                        m.textBox18.Enabled = true;
                        m.textBox20.Enabled = true;

                    }

                    catch
                    {
                        MessageBox.Show("there was an error obtaining v-curve data.  Possibly lost star or too dim, try again", "scopefocus");
                        return;
                    }


                }

                else
                {
                    MainWindow.posMin = _posminHFR;
                }
                m.textBox4.Text = MainWindow.posMin.ToString();
                m.textBox1.Text = m.focuser.Position.ToString();

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
                m.Log("V-Curve: Results:" + "\tPos" + MainWindow.posMin.ToString() + " \t  minHFR" + min.ToString() + "\tmaxMax" + avgMax.ToString() + "\t" + MainWindow.Filtertext);
                // Log("Posmin: " + posMin.ToString() + "\tminHFR " + min.ToString() + "\tminHFRpos " + _posminHFR.ToString());
                m.FileLog2("Posmin: " + MainWindow.posMin.ToString() + "\tminHFR " + min.ToString() + "\tminHFRpos " + _posminHFR.ToString() + "\tmaxMAx " + maxarrayMax.ToString() + "\tmaxMaxPos " + m.posmaxMax.ToString());
                //removed & radiobutton1,checked from below
                if (ffenable == 1)
                {
                    for (int xx = 0; xx < vN; xx++)
                    {
                        if (xx == 0)
                        {
                            //log.WriteLine(DateTime.Now);
                            //log.WriteLine("Spredsheet friendly version of data above");
                            //log.WriteLine(MainWindow.Filtertext);
                            m.FileLog2(DateTime.Now.ToString());
                            m.FileLog2("Spredsheet friendly version of data above");
                            m.FileLog2(MainWindow.Filtertext);

                        }
                        // log.WriteLine("HFR " + list[xx].ToString() + "\tPosition " + minHFRpos[xx].ToString()); // 7-25-14
                        m.FileLog2("HFR " + list[xx].ToString() + "\tPosition " + minHFRpos[xx].ToString());
                    }
                    //   logData.WriteLine(DateTime.Now + "\t" + vN.ToString() + "\t" + _slopeHFRdwn.ToString() + "\t" + _slopeHFRup.ToString() + "\t" + XintDWN.ToString() + "\t" + XintUP.ToString() + "\t" + _pID.ToString() + "\t" + _apexHFR.ToString() + "\t" + MainWindow.Filtertext);
                    m.FileLog2(DateTime.Now + "\t" + vN.ToString() + "\t" + _slopeHFRdwn.ToString() + "\t" + _slopeHFRup.ToString() + "\t" + XintDWN.ToString() + "\t" + XintUP.ToString() + "\t" + _pID.ToString() + "\t" + min.ToString() + "\t" + MainWindow.Filtertext);
                }
                //  log.Close();
                //   logData.Close();
                if ((vDone == 1) & (ffenable == 1))
                {

                    m.fileSystemWatcher1.EnableRaisingEvents = false;
                    m.fileSystemWatcher2.EnableRaisingEvents = false;
                    m.fileSystemWatcher5.EnableRaisingEvents = false; // added to test metric HFR
                                                                      //handle repeated fine v curves
                    if (GlobalVariables.FineVRepeat > 1)
                    {
                        m.button3.PerformClick();

                    }
                    else
                    {
                        if (m.checkBox22.Checked == true)
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
                            if (!m.UseClipBoard.Checked)
                            {

                                //10-11-16  
                                m.serverStream = m.clientSocket.GetStream();
                                byte[] outStream = System.Text.Encoding.ASCII.GetBytes("listenport 0" + "\n");
                                m.serverStream.Write(outStream, 0, outStream.Length);
                                m.serverStream.Flush();

                                Thread.Sleep(3000);
                                m.serverStream.Close();
                                SetForegroundWindow(Handles.NebhWnd);
                                Thread.Sleep(1000);
                                PostMessage(Handles.Aborthwnd, BN_CLICKED, 0, 0);
                                Thread.Sleep(1000);
                                //    NebListenOn = false;
                                // clientSocket.GetStream().Close();//added 5-17-12
                                //  clientSocket.Client.Disconnect(true);//added 5-17-12
                                m.clientSocket.Close();
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

                                m.delay(1);
                                //Clipboard.SetText("//NEB Listen 0");
                                //msdelay(750);
                            }
                            if (m.metricpath != null)
                                File.Delete(m.metricpath[0]);
                            //  currentmetricN = 0;
                            //return;
                            m.fileSystemWatcher5.EnableRaisingEvents = false;
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
                        m.standby();

                    }

                }
                //end coarse v-curve and goes to rough focus point
                if ((vDone == 1) & (tempon == 0) & (backlashDetermOn == false) & (ffenable != 1))
                {
                    m.fileSystemWatcher1.EnableRaisingEvents = false;
                    m.fileSystemWatcher2.EnableRaisingEvents = false;
                    if (m.checkBox29.Checked == false)
                    {

                        m.gotopos(MainWindow.posMin - (int)m.numericUpDown40.Value * 4);


                        // gotopos(posMin - 3000);//  ***** added 1-4-13 **** take up backlash,  3000 works for geared stepper on fsq85, my not need as much for tsa120
                        Thread.Sleep(1000);
                    }
                    m.gotopos(Convert.ToInt32(MainWindow.posMin));
                    vDone = 0;//****  added 11-20-14

                    m.standby();

                }
                //reset for more v -curves for tempcal
                if ((tempon == 1) & (vDone != 1))
                {
                    repeatDone = 0;
                    repeatTotal = 0;
                    m.fileSystemWatcher1.EnableRaisingEvents = true;
                }
                //tempcal done
                if ((tempon == 1) & (vDone == 1))
                {
                    tempsum = 0;
                    vDone = 0;
                    vv = 0;
                    min = 500;
                    vProgress = 0;
                    m.fileSystemWatcher1.EnableRaisingEvents = false;
                    int finegoto = MainWindow.posMin - ((((int)m.numericUpDown3.Value) / 2) * ((int)m.numericUpDown8.Value));
                    if ((finegoto - 100) < 0)
                    {
                        DialogResult result1 = MessageBox.Show("Goto Exceeds Full In", "scopefocus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (result1 == DialogResult.OK)
                        {
                            return;
                        }
                    }
                    m.chart1.Series[3].Points.AddXY(Convert.ToDouble(MainWindow.posMin), Convert.ToDouble(tempavg));
                    if (m.checkBox29.Checked == false)
                    {
                        m.gotopos(Convert.ToInt32(finegoto - 100));//take up backlash                            
                        Thread.Sleep(2000);
                        m.gotopos(Convert.ToInt32(finegoto - ((int)m.numericUpDown8.Value) * 4));
                        Thread.Sleep(1000);
                        m.gotopos(Convert.ToInt32(finegoto - ((int)m.numericUpDown8.Value) * 3));
                        Thread.Sleep(1000);
                        m.gotopos(Convert.ToInt32(finegoto - ((int)m.numericUpDown8.Value) * 2));
                        Thread.Sleep(1000);
                    }
                    m.gotopos(Convert.ToInt32(finegoto - (int)m.numericUpDown8.Value));
                    Thread.Sleep(1000);
                    m.gotopos(Convert.ToInt32(MainWindow.posMin));

                    repeatProgress = 0;
                    vcurveenable = 0;
                    Array.Clear(list, 0, m.arraysize2);
                    Array.Clear(listMax, 0, m.arraysize2);
                    _hFRarraymin = 999;
                    maxarrayMax = 1;
                    m.tempcal();

                }
                vcurveenable = 0;//added 9-12
                                 //   return;  //added 11-20-14
            }




            if ((backlashDetermOn == true) & (vDone == 1))
            {
                backlashCount++;
                if (backlashCount == backlashN - 1)
                {
                    m.standby();
                }
                if (backlashOUT == true)
                {
                    backlashOUT = false;
                    backlashPosOUT = MainWindow.posMin;
                }
                else
                {

                    backlashOUT = true;
                    backlashPosIN = MainWindow.posMin;
                }

                if ((backlashCount == 2) || (backlashCount == 4) || (backlashCount == 6) || (backlashCount == 8) || (backlashCount == 10))
                {

                    backlash = Math.Abs(backlashPosIN - backlashPosOUT);
                    m.textBox8.Text = backlash.ToString();
                    m.chart1.Series[2].Points.AddXY(Convert.ToDouble(backlashCount), Convert.ToDouble(backlash));
                    backlashSum = backlash + backlashSum;
                    backlashAvg = backlashSum / (backlashCount / 2);
                    m.Log("Backlash: N " + backlashCount.ToString() + "\tPosOUT " + backlashPosOUT.ToString() + "\tPosIN " + backlashPosIN.ToString() + "\tBacklash " + backlash.ToString() + "\tAvg " + backlashAvg.ToString());
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
                        m.FileLog2(DateTime.Now.ToString());
                        m.FileLog2("Backlash N" + "\tPosOUT" + "\tPosIN" + "\tBacklash" + "\tAvg");
                    }
                    //  log.WriteLine("         " + backlashCount.ToString() + "\t  " + backlashPosOUT.ToString() + "\t " + backlashPosIN.ToString() + "\t   " + backlash.ToString() + "\t\t" + backlashAvg.ToString());
                    m.FileLog2("         " + backlashCount.ToString() + "\t  " + backlashPosOUT.ToString() + "\t " + backlashPosIN.ToString() + "\t   " + backlash.ToString() + "\t\t" + backlashAvg.ToString());
                    if (backlashCount == backlashN)
                    {
                        //  log.WriteLine("Avg Backlash: " + backlashAvg.ToString());
                        m.FileLog2("Avg Backlash: " + backlashAvg.ToString());
                        m.textBox8.Text = "Avg " + backlashAvg.ToString();

                    }
                    //  log.Close();

                }

                tempsum = 0;
                vDone = 0;
                vv = 0;
                min = 500;

                maxMax = 1;
                m.fileSystemWatcher1.EnableRaisingEvents = true;

                vProgress = 0;
                repeatProgress = 0;
                //?????not sure why below???
                MainWindow.posMin = count;
                Array.Clear(listMax, 0, m.arraysize2);
                Array.Clear(list, 0, m.arraysize2);
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

    }



}
}
