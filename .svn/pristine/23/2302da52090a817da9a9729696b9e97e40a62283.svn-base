using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;


namespace Pololu.Usc.ScopeFocus
{
    class centerHere
    {
        MainMenu L = new MainMenu();
        private double RA;
        private double DEC;
        public double[,] starsXY = new double[10, 2];
        public static string solveImage = "";
        private int sigma;
        private double High;
        private double Low;
        int IndexOfSecond(string theString, string toFind)
        {
            int first = theString.IndexOf(toFind);

            if (first == -1) return -1;

            // Find the "next" occurrence by starting just past the first
            return theString.IndexOf(toFind, first + 1);
        }
        bool XYfound = false;
        public void parseXY()
        {
            try
            {
                MainWindow L = new MainWindow();
                string commandXY = "./tablist " + "solve-indx.xyls";
                tabList(commandXY);
                StreamReader reader = new StreamReader(@"c:\cygwin\home\astro\text2.txt"); //read the cygwin log file
                //  reader = FileInfo.OpenText("filename.txt");
                string line;
                // while ((line = reader.ReadToEnd()) != null) {
                line = reader.ReadToEnd();
                string[] items = line.Split('\n');

                //    string[,] FieldLine = new String[4, 2];
                string find = "   1 ";
                int count = 0;
                L.FileLog("    X     " + "     Y     ");
                foreach (string item in items)
                {

                    if (item.Contains(find) && count < 10)
                    {
                        //  Log("found pos " + pos.ToString());

                        XYfound = true;
                    }

                    if (XYfound && count < 10)
                    {

                        //get and item for each x and y (then RA and DEc...later)
                        // 4 x 2 array with x,y of first 4 stars 
                        starsXY[count, 0] = Convert.ToDouble(item.Substring(12, 15));
                        starsXY[count, 1] = Convert.ToDouble(item.Substring(36, 15));
                        // int nextSpace = IndexOfSecond(item, "        ");
                        //  FieldLine[count, 1] = item.Substring(36, 15);

                        L.FileLog(starsXY[count, 0] + "    " + starsXY[count, 1]);
                        count++;

                    }

                }
                if (!XYfound)
                {
                    L.Log("ParseXY Error - Aborted");
                    return;
                }
                reader.Close();
            }
            catch (Exception e)
            {
                MainWindow L = new MainWindow();
                L.Log("failed to parse X,Y data");
                L.FileLog("failed to parse X,Y data:  " + e.ToString());
            }

        }

        bool RDfound = false;
        public double[,] starsRD = new double[10, 2];
        private void ParseRD()//*****combine w/ parse XY  ***** only diff is command
        {
            try
            {
                MainWindow L = new MainWindow();
                int count;
                string find = "   1 ";
                //    int pos = 0;

                string commandRD = "./tablist " + "solve.rdls";
                tabList(commandRD);
                //not reading a new text2 file w/ RD data...its using the prev XY one....

                StreamReader reader2 = new StreamReader(@"c:\cygwin\home\astro\text2.txt"); //read the cygwin log file
                //  reader = FileInfo.OpenText("filename.txt");
                string line2;
                // while ((line = reader.ReadToEnd()) != null) {
                line2 = reader2.ReadToEnd();
                string[] items2 = line2.Split('\n');

                //   string[,] FieldLine2 = new String[4, 2];
                // string find = "   1 ";
                //   int pos = 0;
                RDfound = false;
                count = 0;
                L.FileLog("     RA     " + "    Dec     ");
                foreach (string item2 in items2)
                {

                    if (item2.Contains(find) && count < 10)
                    {
                        //   Log("found pos " + pos.ToString());
                        //   Log("RA Dec data
                        RDfound = true;
                    }

                    if (RDfound && count < 10)
                    {

                        //get and item for each x and y (then RA and DEc...later)
                        // 4 x 2 array with x,y of first 4 stars 
                        starsRD[count, 0] = Convert.ToDouble(item2.Substring(12, 15));
                        starsRD[count, 1] = Convert.ToDouble(item2.Substring(34, 15));
                        // int nextSpace = IndexOfSecond(item, "        ");
                        //FieldLine2[count, 1] = item2.Substring(34, 15);

                        L.FileLog(starsRD[count, 0] + "    " + starsRD[count, 1]);

                        count++;

                    }

                }
                if (!RDfound)
                {
                    L.Log("ParseRD Error - Aborted");
                    return;
                }

            }
            catch (Exception e)
            {
                MainWindow L = new MainWindow();
                L.Log("failed to parse RA/Dec data");
                L.FileLog("failed to parse RA/Dec data:  " + e.ToString());
            }

        }

        public double centerHereRA = 0;
        public double centerHereDec = 0;

        private void cramersRule()
        {
            try
            {
                MainWindow L = new MainWindow();
                double centerDec = DEC;
                double centerRA = RA;
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
                //Log("a = " + aaa.ToString());
                //Log("b = "  + bbb.ToString());
                //Log("c = " + ccc.ToString());
                //Log("d = " + ddd.ToString());
                //Log("e = " + eee.ToString());
                //Log("f = " + fff.ToString());
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
                //if (checkBox31.Checked == true)
                //{
                //     int StatusstripHandle = FindWindowEx(Handles.NebhWnd, 0, "msctls_statusbar32", null);

                //         IntPtr statusHandle = new IntPtr(StatusstripHandle);
                //         StatusHelper sh = new StatusHelper(statusHandle);

                //    int comma = sh.Captions[2].IndexOf(",");
                //    int length = sh.Captions[2].Length;
                //         xt = Convert.ToDouble(sh.Captions[2].Substring(0,comma));
                //         yt = Convert.ToDouble(sh.Captions[2].Substring(comma,length));
                //    Log("X = " + xt.ToString() + "     Y = " + yt.ToString());
                //    }

                //         else
                //         {

                xt = Convert.ToDouble(textBox2.Text);//arbitrary for now //input x pixel for target 
                yt = Convert.ToDouble(textBox1.Text);
                L.FileLog("Offset X = " + xt.ToString() + "     Y = " + yt.ToString());
                L.FileLog("Using solved Image CenterRA = " + centerRA.ToString() + "    CenterDec = " + centerDec.ToString());
                //  }
                var sigma1 = Math.Sqrt(s / (n - 3));
                var sigma2 = Math.Sqrt(ds / (n - 3));
                var sigma3 = (sigma1 / FL) * 3600 / (dr * 15 * Math.Cos(d0));
                var sigma4 = (sigma2 / FL) * 3600 / dr;
                var sigma5 = Math.Sqrt(res2 / (n - 3));
                var sigma6 = (sigma5 / FL) * 3600 / dr;
                L.Log("RMS residual RA = " + sigma3.ToString() + " arcsec");
                L.FileLog("RMS residual RA = " + sigma3.ToString() + " arcsec");
                L.Log("RMS residual Dec = " + sigma4.ToString() + " arcsec");
                L.FileLog("RMS residual Dec = " + sigma4.ToString() + " arcsec");
                L.Log("RMS residual = " + sigma6.ToString() + " arcsec");
                L.FileLog("RMS residual = " + sigma6.ToString() + " arcsec");
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
                centerHereRA = a6 / (dr * 15);//was * 15

                L.Log("Calculated RA = " + centerHereRA.ToString());

                // result = sexagesimal(v);  //use this to convert to hh:mm:ss
                //  var ia1=result.hours;
                //  var ia2=result.minutes;
                //  var a3=result.seconds;
                // Find declination of target
                double d6 = Math.Atan((sd + yy * cd) / g);
                centerHereDec = d6 / dr;
                L.Log("Calculated Dec = " + centerHereDec.ToString());
                L.FileLog("Calculated Offset:  Dec = " + centerHereDec.ToString() + "     RA = " + centerHereRA.ToString());
                // result = sexagesimal(v);
                //  var id1=result.hours;
                //   var id2=result.minutes;
                //  var d3=result.seconds;
                //  form.radect.value = ia1+" "+ia2+" "+a3+" "+id1+" "+id2+" "+d3;

            }
            catch (Exception e)
            {
                MainWindow L = new MainWindow();
                L.Log("Calculation failed");
               L. FileLog("Calculation Failed :  " + e.ToString());
            }

        }

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
                proc.StartInfo.FileName = @"C:\cygwin\bin\mintty.exe";
                // @"c:/cygwin/bin/mintty.exe";

                proc.StartInfo.Arguments = "--log /home/astro/text2.txt -i /Cygwin-Terminal.ico -";
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
                SendKeys.Send("cd" + " " + "/home/astro");
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
                MainWindow L = new MainWindow();

                L.Log("tablist.exe failed");
                L.FileLog("tablist.exe failed:  " + e.ToString());
            }

        }
    }
}
