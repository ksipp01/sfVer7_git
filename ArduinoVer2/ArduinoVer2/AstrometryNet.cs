using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using nom.tam.fits; 
using System.Threading;

using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;

using Microsoft.Win32;
using System.Web.Script.Serialization;
using System.Runtime.Serialization.Json;
using System.Net;
using Newtonsoft.Json;

using System.Json;
using System.Text.RegularExpressions;


namespace Pololu.Usc.ScopeFocus
{
    public partial class AstrometryNet : Form
    {
        public AstrometryNet()
        {
            InitializeComponent();
            
        }
    
        private static double raCenter;
        //public static double RaCenter
        //{
        //    get { return raCenter; }
        //    set { raCenter = value; }
        //}
        private static double decCenter;
        //public static double DecCenter
        //{
        //    get { return DecCenter; }
        //    set { DecCenter = value; }
        //}
        internal string apikey = "ckylhfaafccqhumf";  // ? to settings
        internal string session = "";
        private string URI = "http://nova.astrometry.net";
        private string path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData);
        System.Timers.Timer timer = new System.Timers.Timer();
        internal string jobid = "";
        private string status = "pending";
        private int attempts = 6;
        private int attemptNumber = 0;

        string subid = "";

       
        //   MainWindow mw = new MainWindow();

        // to do:
        //fix loggin in lofg window fi possible 
        //continue online solve in scopefus not to use the data ....




        public void Log(string text)
        {
            if (LogTextBox.InvokeRequired)
            {
                Invoke(new Action<string>(Log), new object[] { text });
                return;
            }


            if (LogTextBox.Text != "")
                LogTextBox.Text += Environment.NewLine;
            LogTextBox.Text += DateTime.Now.ToLongTimeString() + "  " + text;
            LogTextBox.SelectionStart = LogTextBox.Text.Length;
            LogTextBox.ScrollToCaret();

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
        public void FileLog2(string textlog)  //for non-error file log
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
            return;
        }

        private void SetTimer()
        {
            timer.Interval = 10000;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(timer_Elapsed);
            timer.AutoReset = true;
        }

       

        private async void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (attemptNumber < attempts)
            {
                if ((status == "pending") & (jobid == ""))
                {
                    timer.Start();
                    getJobId(subid);

                }
                if ((jobid != "") & (status == "pending"))
                {
                    attemptNumber = 0;
            //        Log("checking solve status - " + attemptNumber.ToString());
                    timer.Start();
                    checkStatus(jobid); //jobid

                }

                if (status == "success")
                {
                    Log("Plate Solve complete");
                    FileLog2("Plate Solve complete - " + DateTime.Now);
                    timer.Stop();
                    status = "pending";
                    attemptNumber = 0;
                    getCalibData(jobid);
                }
                else
                    timer.Start();
                attemptNumber++;
            }
            else
            {
                Log("Status Check Failed");
                FileLog2("Status Check Failed");
                timer.Stop();
                DialogResult result;
                result = MessageBox.Show("Status Check Failed - Retry?", "scopefocus",
                                MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.Retry)
                {
                    timer.Start();
                    attempts = 0;
                    getJobId(subid);
                }
                    //await GetSession(apikey).ConfigureAwait(false);
             //   GetSession(apikey);
                else
                    return;
            }

        }



        //private void Log(string logtext)
        //{
        //    string fullpath = path + "\\astrometrylog.txt";
        //    StreamWriter log;
        //    if (!File.Exists(fullpath))
        //    {
        //        log = new StreamWriter(fullpath);
        //    }
        //    else
        //    {
        //        log = File.AppendText(fullpath);
        //    }
        //    log.WriteLine(logtext);
        //    log.Close();
        //    return;
        //}


        internal static JsonObject getLoginJSON(string s)
        {
            JsonObject obj = new JsonObject();
            obj["apikey"] = s;
            return obj;
        }

        internal static JsonObject createJson(string id, string value)
        {
            JsonObject obj = new JsonObject();
            obj[id] = value;
            return obj;
        }

        private async Task GetSession(string apikey)
        {
           // MainWindow.AstrometryRunning = true;
            //eventaully add to settings
            JsonObject obj = new JsonObject();
            JsonObject jsonObject = getLoginJSON(apikey);
            String input = "&request-json=" + jsonObject.ToString();
         //   Log("Astrometry.net login started");
         //   FileLog2("Astrometry.net login started - apikey: " + apikey);
            string baseAddress = "http://nova.astrometry.net/api/login";
            try
            {
                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    HttpContent contentPost = new StringContent(input, Encoding.UTF8, "application/x-www-form-urlencoded");

                    using (var response = await httpClient.PostAsync(baseAddress, contentPost).ConfigureAwait(false))
                    {
                        string responseData = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        var serializer = new JavaScriptSerializer();
                        var result2 = serializer.Deserialize<Dictionary<string, string>>(responseData);  //e.g. http://procbits.com/2011/04/21/quick-json-serializationdeserialization-in-c
                        session = result2["session"];
                    }
                }
                if (session != "")
                {

                   Log("Astrometry.net login successful");
                   FileLog2("Astrometry.net login successful,  Session: " + session);
                    PostFile(session, GlobalVariables.SolveImage);
                }
            }
            catch (Exception e)
            {
                    FileLog("Get session error: " + e.ToString());
            }
        }

        internal static JsonObject getUploadJson(string session)
        {
            JsonObject obj = new JsonObject();
            obj["publicly_visible"] = "y";
            obj["allow_modifications"] = "d";
            obj["session"] = session;
            obj["allow_commercial_use"] = "d";
            return obj;
        }

        private async void PostFile(string session, string filename)
        {
            Log("Uploading Image");
            FileLog2("Uploading Image" + DateTime.Now);
       //     string URI = "http://nova.astrometry.net/api/upload";
            JsonObject obj = new JsonObject();
            JsonObject jsonObject = getUploadJson(session);
            String input = jsonObject.ToString();  // was &json-request...
            string baseAddress = "http://nova.astrometry.net/api/upload";
            using (var httpClient = new HttpClient())
            {
                using (var content = new MultipartFormDataContent("form-data"))
                {
                    try
                    {
                        //  var fileContent = new ByteArrayContent(File.ReadAllBytes(Path.GetFileName(MainWindow.solveImage)));
                        var fileContent = new ByteArrayContent(File.ReadAllBytes(GlobalVariables.SolveImage));
                        httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                        var stringContent = new StringContent(jsonObject.ToString());
                        stringContent.Headers.Add("Content-disposition", "form-data; name=\"request-json\"");
                        content.Add(stringContent, "text/plain");
                        string path2 = GlobalVariables.SolveImage; // for testing @"C:\Users\ksipp_000\M81M82.fit";
                        FileStream fs = File.OpenRead(GlobalVariables.SolveImage);

                        var streamContent = new StreamContent(fs);
                        streamContent.Headers.Add("Content-type", "application/octet-strem");
                        streamContent.Headers.Add("Content-disposition", "form-data; name=\"file\"; filename=\"" + Path.GetFileName(path2) + "\"");
                        content.Add(streamContent, "file", Path.GetFileName(path2));

                        using (var response = await httpClient.PostAsync(baseAddress, content))
                        {
                            string responseData = await response.Content.ReadAsStringAsync();

                            var serializer = new JavaScriptSerializer();
                            var result2 = serializer.Deserialize<Dictionary<string, string>>(responseData);  //e.g. http://procbits.com/2011/04/21/quick-json-serializationdeserialization-in-c
                            string sucess = result2["status"];
                            subid = result2["subid"];
                            string hash = result2["hash"];
                            Log("Upload Complete - submission id:   " + subid);
                            FileLog2("Upload Complete - submission id:   " + subid);

                        }

                        SetTimer();
                        timer.Enabled = true;
                        timer.Start();
                    }
                    catch (Exception e)
                    {
          
                              FileLog("Upload File error:   " + e.ToString());
                    }


                }
            }

        }

        internal async virtual void getCalibData(string jobid)
        {

            string result = "";
            StringBuilder sb = new StringBuilder();
            string URI = "http://nova.astrometry.net/api/jobs/" + jobid + "/calibration";
            var httpClient = new HttpClient();
            try
            {
                using (var response = await httpClient.GetAsync(URI))
                {
                    var responseData = await response.Content.ReadAsStreamAsync();
                    using (var reader = new StreamReader(responseData))
                    {

                        string line = "";
                        while (!string.ReferenceEquals((line = reader.ReadLine()), null))
                        {
                            sb.Append(line);
                        }
                        result = sb.ToString();
                        var serializer = new JavaScriptSerializer();
                        var result2 = serializer.Deserialize<Dictionary<string, string>>(result);
                        string parity = result2["parity"];
                        string orientation = result2["orientation"];
                        string pixscale = result2["pixscale"];
                        string radius = result2["radius"];
                        string radecimal = result2["ra"];
                        string decdecimal = result2["dec"];
                        raCenter = double.Parse(radecimal);
                        MainWindow.RA = raCenter / 15;
                        decCenter = double.Parse(decdecimal);
                        MainWindow.DEC = decCenter;

                        if ((raCenter != 0) || (decCenter != 0))
                        {
                           Log("Calibration data success:  RAcenter: " + raCenter/15 + "    DECcenter: " + decCenter);
                           FileLog2("Calibration data:  RA: " + raCenter/15 + "    DEC: " + decCenter + "    " + DateTime.Now);
                          //  if (File.Exists(GlobalVariables.Path2 + "\\axy.fits"))
                            //    File.Delete(GlobalVariables.Path2 + "\\axy.fits");
                          //  if (File.Exists(GlobalVariables.Path2 + "\\rdls.fits"))
                          //      File.Delete(GlobalVariables.Path2 + "\\rdls.fits");
                            if (File.Exists(GlobalVariables.Path2 + "\\corr.fits"))
                                File.Delete(GlobalVariables.Path2 + "\\corr.fits");
                          //  if (File.Exists(GlobalVariables.Path2 + "\\wcs.fits"))
                            //    File.Delete(GlobalVariables.Path2 + "\\wcs.fits");

                          //  GetFile(jobid, "axy.fits", "/axy_file/");
                         //   GetFile(jobid, "rdls.fits", "/rdls_file/");
                          //  GetFile(jobid, "wcs.fits", "/wcs_file/");
                            GetFile(jobid, "corr.fits", "/corr_file/");

                        }
                        else
                        {
                            Log("Error getting calibration data");
                        }



                        ///"WCS.fits" - "/wcs_file/");
                        /// "ANNOTATED_FULL.jpg" -  "/annotated_full/");
                        /// "AXY.fits"  - "/axy_file/");
                        /// "NEW-IMAGE.fits" - "/new_fits_file/");
                        /// "rdls.fits  - "/rdls.fits/";


                    }
                }

            }
            catch (Exception e)
            {
                 FileLog("Calibration data error:   " + e.ToString());
            }
        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    GetSession(apikey);

        //}

        //private void button3_Click(object sender, EventArgs e)
        //{
        //    PostFile(session, "M81M82.fit");

        //}

        //private void button4_Click(object sender, EventArgs e)
        //{
        //    getCalibData("1469005");
        //}


        private async void checkStatus(string id)
        {
            FileLog2("status check - " + DateTime.Now);
            Log("Checking status");

            string result = "";
            StringBuilder sb = new StringBuilder();
            string url = URI + @"/api/jobs/" + id;
            var httpClient = new HttpClient();
            try
            {
                using (var response = await httpClient.GetAsync(url))
                {
                    var responseData = await response.Content.ReadAsStreamAsync();
                    using (var reader = new StreamReader(responseData))
                    {

                        string line = "";
                        while (!string.ReferenceEquals((line = reader.ReadLine()), null))
                        {
                            sb.Append(line);
                        }
                        result = sb.ToString();
                    }
                    //not getting json results, seems like bad url

                    var serializer = new JavaScriptSerializer();
                    var result2 = serializer.Deserialize<Dictionary<string, string>>(result);
                    status = result2["status"];

                    return;

                }
            }
            catch (Exception e)
            {
                FileLog("Check status Error:  " + e.ToString());
            }

        }

        public class RootObject
        {
            public string processing_started { get; set; }
            // public List<List<int>> job_calibrations { get; set; }
            public List<List<string>> job_calibrations { get; set; }
            //  public List<int> jobs { get; set; }
            public List<string> jobs { get; set; }
            public string processing_finished { get; set; }
            public int user { get; set; }
            public List<int> user_images { get; set; }
        }

        private async void getJobId(string subid)
        {

            StringBuilder sb = new StringBuilder();
            string result;
            string url = "http://nova.astrometry.net/api/submissions/" + subid;
            var httpClient = new HttpClient();
            try
            {
                using (var response = await httpClient.GetAsync(url))
                {
                    var responseData = await response.Content.ReadAsStreamAsync();
                    using (var reader = new StreamReader(responseData))
                    {
                        string line = "";
                        while (!string.ReferenceEquals((line = reader.ReadLine()), null))
                        {
                            sb.Append(line);
                        }
                        result = sb.ToString();
                        JsonSerializer serializer = new JsonSerializer();
                        var jsonresults = JsonConvert.DeserializeObject<RootObject>(result); ///this need error handling for the value saved in txt file
                        int count = jsonresults.job_calibrations.Count;
                        //  int count = jsonresults.jobs.Count;
                        if (count > 0)
                            jobid = jsonresults.job_calibrations[0][count - 1];
                        else
                        {
                            Log("Waiting for JobID");
                            return;
                        }

                        //   jobid = jsonresults.jobs[0];
                        if (jobid != "")
                        {
                            Log("Upload succsesful");
                            FileLog2("Upload succsesful.  Jobid: " + jobid + "   " + DateTime.Now);
                            //timer.Enabled = true;
                            //timer.Start();  //check for completion
                        }
                        //else
                        //{

                        //    Log("Failed to get jobid   " + DateTime.Now);
                        //    return;
                        //}
                        //  }
                        // catch (Exception e)
                        //{
                        //    MessageBox.Show("Jobid Error: " + e);
                        //    return;
                        //    Log("Error getting jobid" + e.ToString());
                        //  }


                        return;
                    }
                }
            }
            catch (Exception e)
            {
                FileLog("Get jobid error:  " + e.ToString());
            }
        }
        // was static async ...
        public async Task DownloadAsync(Uri requestUri, string filename)  //filename is save path\name  the filename from the webpage is specified in the requestUri  
        {
            try
            {
                if (filename == null)
                    throw new ArgumentNullException("filename");

                using (var httpClient = new HttpClient())
                {
                    using (var request = new HttpRequestMessage(HttpMethod.Get, requestUri))
                    {
                        using (
                            Stream contentStream = await (await httpClient.SendAsync(request)).Content.ReadAsStreamAsync(),
                            stream = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            await contentStream.CopyToAsync(stream);
                        }
                    }
                }
            }

            catch (Exception e)
            {

                FileLog("Downloadasyc error:  " + e.ToString());
            }
        }



        private async void GetFile(string jid, string file, string suburl) //jid is jobid, file name of file to download (e.g. wcs.fits, suburl is eg "/wcs_file"  see list below 
        {
            try
            {
               // Log("Downloading " + file);
                string newurl = URI + suburl + jid;
                Uri resultUrl = new Uri(newurl, false);
                Log("Downloading " + newurl);
                FileLog2("Downloaded : " + newurl);

                //  DownloadAsync(resultUrl, "c:\\atest3\\wcs.fits"); 
                await DownloadAsync(resultUrl, GlobalVariables.Path2 + @"\" + file);  //source url and destination path(fullname) 


                ///available files and suburls:
                ///
                ///"WCS.fits" - "/wcs_file/");
                /// "ANNOTATED_FULL.jpg" -  "/annotated_full/");
                /// "AXY.fits"  - "/axy_file/");
                /// "NEW-IMAGE.fits" - "/new_fits_file/");
                //   mw.Log("Download Complete");
                //   mw.Refresh();
                //   mw.FileLog2("Download Complete: " + file.ToString());
            }
            catch (Exception e)
            {
               FileLog("Error Downloading: " + e.ToString());
            }
            Log("Files Download to: " + GlobalVariables.Path2);
            Thread.Sleep(1000);
            MakeDataFile();
            MainWindow.AstrometryRunning = false;
        }

        public async Task OnlineSolve(string image)
        {
            //Form2 frm = new Form2();
            //frm.Show();
            //   f2.Show();
            this.Show();
            await GetSession(apikey).ConfigureAwait(false);
        }

        public void MakeDataFile()
        {
            string pathToCorr = "";
            if (GlobalVariables.LocalPlateSolve)
                pathToCorr = @"c:\cygwin\home\astro\solve.corr";
            else
                pathToCorr = GlobalVariables.Path2 + "\\corr.fits";
                
           // Fits f = new Fits(GlobalVariables.Path2 + "\\corr.fits"); // original
           Fits f = new Fits(pathToCorr);
            BinaryTableHDU h = (BinaryTableHDU)f.GetHDU(1);

            //   Object[] row23 = h.GetRow(23);
            Object col_x = h.GetColumn(0);
            Object col_y = h.GetColumn(1);
            Object col_ra = h.GetColumn(2);
            Object col_dec = h.GetColumn(3);
            float[] x = new float[h.NRows];
            float[] y = new float[h.NRows];
            float[] ra = new float[h.NRows];
            float[] dec = new float[h.NRows];
            GlobalVariables.CorrFileLines = h.NRows;
            int i = 0;
            foreach (float item in (dynamic)col_x)
            {
                x[i] = (dynamic)(item);
                i++;
            }
            i = 0;
            foreach (float item2 in (dynamic)col_y)
            {
                y[i] = (dynamic)(item2);
                i++;
            }
            i = 0;
            foreach (float item in (dynamic)col_ra)
            {
                ra[i] = (dynamic)(item)/15;
                i++;
            }
            i = 0;
            foreach (float item2 in (dynamic)col_dec)
            {
                dec[i] = (dynamic)(item2);
                i++;
            }
           
            using (StreamWriter sr = new StreamWriter(GlobalVariables.Path2 + "\\corr2text.txt"))
            {
                // sr.WriteLine("   x   " + "         y   ");
                for (i = 0; i < x.Length; i++)
                {
                    sr.WriteLine(x[i] + "    \t" + y[i] + "    \t" + ra[i] + "    \t" + dec[i]);

                }
                
                sr.Dispose();
            }
            Log("Corr2text file complete in " + GlobalVariables.Path2);
           
        }

    }
}
