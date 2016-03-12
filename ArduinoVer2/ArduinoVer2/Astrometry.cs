using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Windows.Forms;
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
    

    public class Form2: Form
    {
       public TextBox LogTextBox;

        public Form2()
        {
            Text = "Plate Solve";
        }

        private void InitializeComponent()
        {
            this.LogTextBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // LogTextBox
            // 
            this.LogTextBox.BackColor = System.Drawing.SystemColors.MenuBar;
            this.LogTextBox.Location = new System.Drawing.Point(12, 12);
            this.LogTextBox.Multiline = true;
            this.LogTextBox.Name = "LogTextBox";
            this.LogTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.LogTextBox.Size = new System.Drawing.Size(592, 105);
            this.LogTextBox.TabIndex = 17;
            this.LogTextBox.Text = "PLate Solve Log";
            // 
            // Form2
            // 
            this.ClientSize = new System.Drawing.Size(616, 129);
            this.Controls.Add(this.LogTextBox);
            this.Name = "Form2";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }

    class Astrometry
    {

        internal static double raCenter;
        internal static double decCenter;
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

        Form2 f2 = new Form2();
        
       

        //   MainWindow mw = new MainWindow();

        // to do:
        //fix loggin in lofg window fi possible 
        //continue online solve in scopefus not to use the data ....

        


        public void Log (string text)
        {

            if (f2.LogTextBox.InvokeRequired)
            {
                f2.Invoke(new Action<string>(Log), new object[] { text });
                return;
            }
            f2.LogTextBox.Text = text;
        }


        private void SetTimer()
        {
            timer.Interval = 15000;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(timer_Elapsed);
            timer.AutoReset = true;
        }

        public void UpdateTextBox(string text)
        {
          //  mw.AppendText(text);
        }

        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (attemptNumber < attempts)
            {
                if ((status == "pending") & (jobid == ""))
                {
                    timer.Start();
                    //  mw.Log("checking for job id - " + attemptNumber.ToString());
                 //   mw.LogTextBox.Text = "checking for job id - " + attemptNumber.ToString();
                  //  UpdateTextBox("checking for job id - " + attemptNumber.ToString());
                    getJobId(subid);

                }
                if ((jobid != "") & (status == "pending"))
                {
                    attemptNumber = 0;
                  //  mw.Log("checking solve status - " + attemptNumber.ToString());
                    //mw.Refresh();
                    timer.Start();
                    checkStatus(jobid); //jobid

                }

                if (status == "success")
                {
                   // mw.Log("Plate Solve complete");
                  //  mw.FileLog2("Plate Solve complete - " + DateTime.Now);
                   // mw.Refresh();
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
                //mw.Log("Status Check Failed");
              //  mw.FileLog2("Status Check Failed");
                //mw.Refresh();
                timer.Stop();
                DialogResult result;
                result = MessageBox.Show("Status Check Failed - Retry?", "scopefocus",
                                MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.Retry)
                    GetSession(apikey);
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
            //eventaully add to settings
            Log("TASEGAAWEGARHA");
            JsonObject obj = new JsonObject();
            JsonObject jsonObject = getLoginJSON(apikey);
            String input = "&request-json=" + jsonObject.ToString();
          //  mw.FileLog2("Astrometry.net login started - apikey: " + apikey);
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
  
                  //  mw.LogTextBox.Text =" TESTTESTSETSAETASET";
                 //   mw.LogTextBox.Text = "Astrometry.net login successful,  Session: " + session;
                 //   UpdateTextBox("Astrometry.net login successful,  Session: " + session);
                    //   MessageBox.Show("Session number: " + session);
                    //  Log("Astrometry.net login successful,  Session: " + session);
                    //  mw.Log("Astrometry.net login successful,  Session: " + session);
                    //  mw.LogFromClass = "Astrometry.net login successful,  Session: " + session;
                    //   mw.FileLog2("Astrometry.net login successful,  Session: " + session);
                    //  mw.Refresh();
                    PostFile(session, GlobalVariables.SolveImage);
                }
            }
            catch (Exception e)
            {
            //    mw.FileLog("Get session error: " + e.ToString());
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
            string URI = "http://nova.astrometry.net/api/upload";
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
                            //Log("File uploaded, submission id:   " + subid);
                            //mw.Log("File uploaded, submission id:   " + subid);
                            //mw.FileLog2("File uploaded, submission id:   " + subid);
                            //mw.Refresh();
                        }

                        SetTimer();
                        timer.Enabled = true;
                        timer.Start();
                    }
                    catch (Exception e)
                    {
                 //       mw.FileLog2("Upload File error");
                 //       mw.FileLog("Upload File error:   " + e.ToString());
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
                        decCenter = double.Parse(decdecimal);

                        if ((raCenter != 0) || (decCenter != 0))
                        {
                            //  Log("Calibration data success:  RAcenter: " + raCenter + "    DECcenter: " + decCenter);
                          //  mw.Log("Calibration data:  RA: " + raCenter + "    DEC: " + decCenter);
                   //         mw.FileLog2("Calibration data:  RA: " + raCenter + "    DEC: " + decCenter + "    " + DateTime.Now);
                            GetFile(jobid, "axy.fits", "/axy.fits/");
                            GetFile(jobid, "rdls.fits", "/rdls.fits/");
                            GetFile(jobid, "wcs.fits", "/wcs.fits/");
                            
                        }
                        else
                        {
                           // mw.Log("Error getting calibration data");
                          //  mw.Refresh();
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
              //  mw.FileLog2("Calibration data error");
              //  mw.FileLog("Calibration data error:   " + e.ToString());
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
            //mw.FileLog2("status check - " + DateTime.Now);
          //  mw.Log("Checking status");
         //   mw.Refresh();
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
             //   mw.FileLog("Check status Error:  " + e.ToString());
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

                     //       mw.FileLog2("Failed to get jobid   " + DateTime.Now);
                            return;
                        }

                        //   jobid = jsonresults.jobs[0];
                        if (jobid != "")
                        {
                   //         mw.FileLog2("Upload succsesful.  Jobid: " + jobid + "   " + DateTime.Now);
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
            //    mw.FileLog("Get jobid error:  " + e.ToString());
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
               // MainMenu mw = new MainMenu();
            //    mw.FileLog("Downloadasyc error:  " + e.ToString());
            }
        }



        private async void GetFile(string jid, string file, string suburl) //jid is jobid, file name of file to download (e.g. wcs.fits, suburl is eg "/wcs_file"  see list below 
        {
            try
            {
             //   mw.Log("Downloading " + file);
                string newurl = URI + suburl + jid;
                Uri resultUrl = new Uri(newurl, false);

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
            //    mw.FileLog("Error Downloading: " + e.ToString());
            }

        }

        public async Task OnlineSolve(string image)
        {
            //Form2 frm = new Form2();
            //frm.Show();
            f2.Show();
            await GetSession(apikey).ConfigureAwait(false);
        }



    }

}

