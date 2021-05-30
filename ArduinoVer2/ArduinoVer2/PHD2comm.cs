using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Net.Sockets;
using Newtonsoft.Json;
using System.Json;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Web.Script.Serialization;
using System.Net;







namespace Pololu.Usc.ScopeFocus
{


   

    public partial class PHD2comm
    {


       
        TcpClient socket;
         System.Timers.Timer polltimer = new System.Timers.Timer();
        public PHD2comm()
        {
      
            polltimer.Interval = 1000;
            polltimer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimedEvent); 
        }

        //private void Polltimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

    


        private  void OnTimedEvent(object source,System.Timers.ElapsedEventArgs e)
        {

            if (_phd2Connected)
            PHDgetAppState();
        }
        public string Connect()
        {
            try
            {
                // socket = new TcpClient("127.0.0.1", 4400);
                socket = new TcpClient();
                IAsyncResult r = socket.BeginConnect("127.0.0.1", 4400, null, null);

                bool success = r.AsyncWaitHandle.WaitOne(5000, true);

                if (!socket.Connected)
                {
                    // NOTE, MUST CLOSE THE SOCKET

                    socket.Close();
                    throw new ApplicationException("Failed to connect server.");

                }
                else
                {

                    polltimer.Enabled = true;

                    EventMessage msg;
                    //   TcpClient socket = new TcpClient("127.0.0.1", 4400);  
                    // NetworkStream stream = socket.GetStream();
                    var reader = new StreamReader(socket.GetStream());
                    string result = "";
                    StringBuilder sb = new StringBuilder();
                    //    var stm = socket.GetStream();      
                    string line = "";
                    while ((line = reader.ReadLine()) != null)
                    {
                        result = line.ToString();

                        //while (!string.ReferenceEquals((line = reader.ReadLine()), null))
                        //{
                        //    sb.Append(line);
                        //    result = sb.ToString();
                        msg = JsonConvert.DeserializeObject<EventMessage>(result);

                        // if (command == "PHDversion")
                        _phd2Connected = true;
                        return msg.PHDversion;

                        //if (command == "PHDSubver")
                        //    return msg.PHDSubver;
                        //if (command == "Host")
                        //    return msg.Host;
                        //if (command == "State")
                        //    return msg.State;                  
                    }

                    //polltimer.Interval = 1000;
                    //polltimer.Elapsed += PollUpdates;
                    //polltimer.Start();
                    reader.Dispose();

                    return "";

                    //var reader = new StreamReader(socket.GetStream());
                    //string result = "";
                    //StringBuilder sb = new StringBuilder();
                    ////    var stm = socket.GetStream();      
                    //string line = "";
                    //while (!string.ReferenceEquals((line = reader.ReadLine()), null))
                    //{
                    //    sb.Append(line);
                    //    result = sb.ToString();
                    //    //  msg = JsonConvert.DeserializeObject<Jsonrpc>(result);

                    //    //if (command == "State")
                    //    //    return msg.State;                  
                    //}



                    //NetworkStream stream = socket.GetStream();
                    //byte[] inStream = new byte[10025];
                    //stream.Read(inStream, 0, inStream.Length);
                    //string returndata = System.Text.Encoding.ASCII.GetString(inStream);
                    //stream.Flush();

                }
               


                
            }
            catch (Exception)
            {
                //MainWindow m = new MainWindow();
                //m.Log("Failed to connect to PHD2 ");
                //m.FileLog("Failed to connect to PHD2 " + e.ToString());
                return "Failed";
            }
        }

   


        public string DisConnect()
        {

            _phd2Connected = false;
            polltimer.Enabled = false;
            Thread.Sleep(1000);
            //socket.GetStream().Close();
            //socket.Close();
            
            return "closed";
        }

        public class EventMessage
        {
            public string Host { get; set; }
            public string PHDversion { get; set; }
            public string PHDSubver { get; set; }
            //   public string State { get; set; }
        }
        public class Jsonrpc
        {
            public string jsonrpc { get; set; }
            public string result { get; set; }
            public string id { get; set; }
            //   public string State { get; set; }
        }
        public class AppState
        {
            public string State { get; set; }
            public string Event { get; set; }
            //   public string PHDSubver { get; set; }
            //   public string State { get; set; }
        }

        private class Command
        {
            public string method { get; set; }
            //    public string  param { get; set; }
            public int id { get; set; }
        }


        // copy json and use edit-->Paste speicial --> as json. 

        public class Settle         // {"pixels": 1.5, "time": 10, "timeout": 60}
        {
            public double pixels { get; set; }
            public int time { get; set; }
            public int timeout { get; set; }
        }

        public class GuideCommand
        {
            public string method { get; set; }
            public List<object> @params { get; set; }
            public int id { get; set; }
        }

        public void Guide()
        {
            GuideCommand cmd = new GuideCommand();
            Settle stl = new Settle();
            stl.pixels = 1.5;
            stl.time = 10;
            stl.timeout = 60;
            List<object> list = new List<object>();
            cmd.method = "guide";
            cmd.id = 42;
            list.Add(stl);       
            cmd.@params = list;
            NetworkStream stream = socket.GetStream();
            string command = JsonConvert.SerializeObject(cmd) + "\r\n";
            byte[] outStream = System.Text.Encoding.ASCII.GetBytes(command);
            stream.Write(outStream, 0, outStream.Length);
            stream.Flush();
        }





        public void StopCapture()
        {
            Command cmd = new Command();
            cmd.method = "stop_capture";
            //   cmd.param = "";
            //   cmd.id = 0;        
            NetworkStream stream = socket.GetStream();
            //  stream.Flush();
         //   JsonSerializer serializer = new JsonSerializer();
            string command = JsonConvert.SerializeObject(cmd) + "\r\n";
            byte[] outStream = System.Text.Encoding.ASCII.GetBytes(command);
            stream.Write(outStream, 0, outStream.Length);
            stream.Flush();

            //while (getAppState().ToString() != "Stopped")
            //    Thread.Sleep(100);





            //byte[] inStream = new byte[1024];
            //stream.Read(inStream, 0, inStream.Length);
            //string returndata = System.Text.Encoding.ASCII.GetString(inStream);


        }
        // ***** need to build this command and add the Settle parameter within it......


        private static bool _phd2Connected;
        public static bool PHD2Connected
        {
            get { return _phd2Connected; }
            set { _phd2Connected = value; }
        }

        //private static bool _stopped;
        //public static bool Stopped
        //{
        //    get { return _stopped; }
        //    set { _stopped = value; }
        //}

        //private static string _event;
        //public static string Event
        //{
        //    get { return _event; }
        //    set { _event = value; }
        //}
        //private static string _appState;
        //public static string State
        //{
        //    get { return _appState; }
        //    set { _appState = value; }
        //}
        // this will continuously monitor events until (needed) some sort of break.....
        public async void EventMonitor()
        {
            AppState msg;
            string result = "";
            StringBuilder sb = new StringBuilder();
            //    var stm = socket.GetStream();      
            string line = "";
            using (StreamReader reader = new StreamReader(socket.GetStream()))
            {
                //while (!string.ReferenceEquals((line = await reader.ReadLineAsync()), null))
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    sb.Append(line);
                    result = sb.ToString();
                    sb.Clear();
                    msg = JsonConvert.DeserializeObject<AppState>(result);

                    if (msg.Event != null)
                    {
                        //_appState = msg.Event.ToString();
                        break;
                    }

                    //if (msg.Event != null)
                    //    _event = msg.Event.ToString();
                    //if (msg.State != null)
                    //{
                    //  //  socket.Close();
                    // //   reader.Close();
                    //    break;
                    //}

                }
            }
        }
        private static string _PHDappState;
        public static string PHDAppState
        {
            get { return _PHDappState; }
            set { _PHDappState = value; }

        }

        // **** this works! ****


        //public string getAppState()
        //{
        //    PHDgetAppState();
        //    return _PHDappState;
        //}

        public async void PHDgetAppState()  // was getAppState  12-7-16  make async 12-8-16 removed async
        {
            // ****need to stop/cancel the async when form closing.  can prob remove stuff tires thus far....
            try
            {
                Command cmd = new Command();
                cmd.method = "get_app_state";
                //  cmd.method = "get_exposure";
                //   cmd.param = "";
                cmd.id = 1;
                NetworkStream stream = socket.GetStream();
                //  stream.Flush();
                JsonSerializer serializer = new JsonSerializer();
                string command = JsonConvert.SerializeObject(cmd) + "\r\n";
                byte[] outStream = System.Text.Encoding.ASCII.GetBytes(command);
                stream.Write(outStream, 0, outStream.Length);
                // stream.Flush();

                Thread.Sleep(250);


                Jsonrpc msg;
                //   TcpClient socket = new TcpClient("127.0.0.1", 4400);  
                // NetworkStream stream = socket.GetStream();
                var reader = new StreamReader(socket.GetStream());
                string raw = "";
                StringBuilder sb = new StringBuilder();
                //    var stm = socket.GetStream();      
                string line = "";

                //  while (!string.ReferenceEquals((line = reader.ReadLine()), null))
                //   while ((line = reader.ReadLine()) != null) // 12-7-16


                if (reader == null)
                    _phd2Connected = false;

                // while ((line = reader.ReadLine()) != null)
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    raw = line.ToString();
                    //sb.Append(line);
                    //result = sb.ToString();
                    if (raw.Substring(0, 10) == "{\"jsonrpc\"")
                    {
                        msg = JsonConvert.DeserializeObject<Jsonrpc>(raw);
                        if (msg.result != null && msg.result != "0")
                            _PHDappState = msg.result.ToString();
                        else
                            _PHDappState = "Error";
                        //return msg.result.ToString(); // 12-7-16


                    }

                    //if (command == "State")
                    //    return msg.State;                  
                }

                reader.Dispose();

                //  return ""; 12-7-16

            }
            // added 1-9-17 error handle lost phd conneciton
            catch (IOException)
            {
                
                _phd2Connected = false;
            }
            catch (Exception)
            {
                _phd2Connected = false;
            }
            

        }

        public void Loop()
        {
            Command cmd = new Command();
            cmd.method = "loop";
            //   cmd.param = "";
            //   cmd.id = 0;        
            NetworkStream stream = socket.GetStream();
            //  stream.Flush();
            JsonSerializer serializer = new JsonSerializer();
            string command = JsonConvert.SerializeObject(cmd) + "\r\n";
            byte[] outStream = System.Text.Encoding.ASCII.GetBytes(command);
            stream.Write(outStream, 0, outStream.Length);
            stream.Flush();


        }




        public async void StartEventMonitor()
        {

            // *** this reads each individual json sequentially.  
            // ? should be done in background???
            // ?need to check each one for what specifically looking for.  
            EventMessage msg;
            string result = "";
            StringBuilder sb = new StringBuilder();
            //    var stm = socket.GetStream();      
            string line = "";
            using (StreamReader reader = new StreamReader(socket.GetStream()))
            {
                while (!string.ReferenceEquals((line = await reader.ReadLineAsync()), null))
                {
                    sb.Append(line);
                    result = sb.ToString();
                    sb.Clear();
                    msg = JsonConvert.DeserializeObject<EventMessage>(result);


                }
            }

        }




        public string show()
        {
            return "";
        }
        public string show(string command)
        {

            EventMessage msg;
            //   TcpClient socket = new TcpClient("127.0.0.1", 4400);  
            // NetworkStream stream = socket.GetStream();
            var reader = new StreamReader(socket.GetStream());
            string result = "";
            StringBuilder sb = new StringBuilder();
            //    var stm = socket.GetStream();      
            string line = "";
            while ((line = reader.ReadLine()) != null)
            {
                result = line.ToString();

                //while (!string.ReferenceEquals((line = reader.ReadLine()), null))
                //{
                //    sb.Append(line);
                //    result = sb.ToString();
                msg = JsonConvert.DeserializeObject<EventMessage>(result);
                if (command == "PHDversion")
                    return msg.PHDversion;
                if (command == "PHDSubver")
                    return msg.PHDSubver;
                if (command == "Host")
                    return msg.Host;
                //if (command == "State")
                //    return msg.State;                  
            }

            reader.Dispose();

            return "";

            //var stm = socket.GetStream();
            //byte[] resp = new byte[2048];
            //var memStream = new MemoryStream();
            //int bytesread = stm.Read(resp, 0, resp.Length);
            //stm.ReadTimeout = 4000;
            //var bytes = 0;

            //do
            //{
            //    try
            //    {
            //        bytes = stm.Read(resp, 0, resp.Length);
            //        memStream.Write(resp, 0, bytes);
            //    }
            //    catch (IOException ex)
            //    {
            //        // if the ReceiveTimeout is reached an IOException will be raised...
            //        // with an InnerException of type SocketException and ErrorCode 10060
            //        var socketExept = ex.InnerException as SocketException;
            //        if (socketExept == null || socketExept.ErrorCode != 10060)
            //            // if it's not the "expected" exception, let's not hide the error
            //            throw ex;
            //        // if it is the receive timeout, then reading ended
            //        bytes = 0;
            //    }
            //} while (bytes > 0);
            //return getVersion(Encoding.GetEncoding(1252).GetString(memStream.ToArray())).ToString();


            //while (bytesread > 0)
            //{
            //    memStream.Write(resp, 0, bytesread);
            //    bytesread = stm.Read(resp, 0, resp.Length);
            //}
            //return Encoding.GetEncoding(1252).GetString(memStream.ToArray());



            //  stream.ReadTimeout = 2000;
            //    Int32 bytes = stream.Read(data, 0, data.Length);

            //   int response = Convert.ToInt32(data[0]);
            //   responseData = response.ToString();

            //   stream.Flush();

            //     return getVersion(str).ToString();

            //return responseData;

        }





    }



}
