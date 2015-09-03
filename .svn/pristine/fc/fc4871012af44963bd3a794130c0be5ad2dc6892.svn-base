using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Pololu.Usc.ScopeFocus
{
    class Handles
    {


        static int _nebhWnd;
        static int _loadScripthwnd;
        static int _nebVNumber;
        static int _phdVNumber;
        static int _pHDhwnd;
        static int _panelhwnd;
        static int _editfound;
        static int _hwndDuration;
        static int _nebCamera;
        static int _captureMainhWnd;
        static int _aborthwnd;
        static int _framehwnd;
        static int _finehwnd;
        static int _panelhwnd2;
        static int _hwndDuration2;
        static int _aborthwnd2;
        static int _framehwnd2;
        static int _finehwnd2;
        static int _loadScripthwnd2;
        static int _gotofocushwnd;
        static int _capturehwnd;
        static int _flathwnd;
        static int _pausehwnd;
        static int _slaveStatushwnd;
   //     static int _serverhwnd;
        public static int NebhWnd
        {
            get { return _nebhWnd;}
            set { _nebhWnd = value;}
        }
        public static int LoadScripthwnd
        {
            get { return _loadScripthwnd; }
            set { _loadScripthwnd = value; }
        }
        public static int NebVNumber
        {
            get { return _nebVNumber; }
            set { _nebVNumber = value; }
        }
        public static int PHDVNumber
        {
            get { return _phdVNumber; }
            set { _phdVNumber = value; }
        }
        public static int PHDhwnd
        {
            get { return _pHDhwnd; }
            set { _pHDhwnd = value; }
        }
        public static int Panelhwnd
        {
            get { return _panelhwnd; }
            set { _panelhwnd = value; }
        }
        public static int Editfound
        {
            get { return _editfound; }
            set { _editfound = value; }
        }
        public static int HwndDuration
        {
            get { return _hwndDuration; }
            set { _hwndDuration = value; }
        }
        public static int NebCamera
        {
            get { return _nebCamera; }
            set { _nebCamera = value; }
        }

        public static int CaptureMainhWnd
        {
            get { return _captureMainhWnd; }
            set { _captureMainhWnd = value; }
        }
        public static int Aborthwnd
        {
            get { return _aborthwnd; }
            set { _aborthwnd = value; }
        }
        public static int Framehwnd
        {
            get { return _framehwnd; }
            set { _framehwnd = value; }
        }

        public static int Finehwnd
        {
            get { return _finehwnd; }
            set { _finehwnd = value; }
        }
        public static int Panelhwnd2
        {
            get { return _panelhwnd2; }
            set { _panelhwnd2 = value; }
        }
        public static int HwndDuration2
        {
            get { return _hwndDuration2; }
            set { _hwndDuration2 = value; }
        }
        public static int Aborthwnd2
        {
            get { return _aborthwnd2; }
            set { _aborthwnd2 = value; }
        }
        public static int Framehwnd2
        {
            get{ return _framehwnd2;}
            set { _framehwnd2 = value; }
        }
        public static int Finehwnd2
        {
            get { return _finehwnd2;}
            set { _finehwnd2 = value; }
        }
        public static int LoadScripthwnd2
        {
            get { return _loadScripthwnd2;}
            set { _loadScripthwnd2 = value; }
        }
        public static int Gotofocushwnd
        {
            get { return _gotofocushwnd;}
            set { _gotofocushwnd = value; }
        }
        public static int Capturehwnd
        {
            get { return _capturehwnd; }
            set { _capturehwnd = value; }
        }
        //public static int NebhWnd
        //{
        //    get { return _nebhWnd; }
        //    set { _nebhWnd = value; }
        //}
        public static int SlaveStatushwnd
        {
            get { return _slaveStatushwnd; }
            set { _slaveStatushwnd = value; }
        }
        public static int Pausehwnd
        {
            get { return _pausehwnd; }
            set { _pausehwnd = value; }
        }
        public static int Flathwnd
        {
            get { return _flathwnd; }
            set { _flathwnd = value; }
        }

        public int Serverhwnd()
        {
            MainWindow L = new MainWindow();
             IntPtr ServerhwndPtr = Handles.SearchForWindow("WindowsForms10", "scopefocus - Main");
             L.Log("scopefocus-server handle found --  " + ServerhwndPtr.ToInt32());
             int _serverhwnd = ServerhwndPtr.ToInt32();
            return _serverhwnd;
        }
        public delegate int Callback(int hWnd, int lParam);
        [DllImport("User32.dll")]
        public static extern Int32 GetWindowText(int hWnd, StringBuilder s, int nMaxCount);
        [DllImport("User32.dll")]
        public static extern Int32 GetClassName(int hWnd, StringBuilder s, int nMaxCount);//added to try to get edit box 
        [DllImport("User32.dll")]
        public static extern Int32 FindWindow(String lpClassName, String lpWindowName);
        [DllImport("User32.dll")]
        public static extern Boolean EnumChildWindows(int hWndParent, Delegate lpEnumFunc, int lParam);
        [DllImport("User32.dll")]
        public static extern int SendMessage(int hWnd, int msg, int wparam, StringBuilder text);
       // [DllImport("USER32.DLL", CharSet = CharSet.Auto)]
      //  public static extern int SendMessage(int hWnd, int msg, int wparam, int lparam);
        
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


        public static IntPtr SearchForWindow(string wndclass, string title)
        {
            SearchData sd = new SearchData { Wndclass = wndclass, Title = title };
            EnumWindows(new EnumWindowsProc(EnumProc), ref sd);
            return sd.hWnd;
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


                EnumChildWindows(_loadScripthwnd, myCallBack, 0);
            }
            catch (Exception e)
            {
                MainWindow L = new MainWindow();
                
               L.Log("GetHandles Error" + e.ToString());
               L.Send("GetHandles Error" + e.ToString());
               L.FileLog("GetHandles Error" + e.ToString());

            }

        }


      //  private bool FindHandlesDone = false;
        public void FindHandles()
        {
        string MainWindowName;
            //******5-30 this might be problem, maybe this should just be done once...not sure if 
            //finding the loadscript though. consider doing seperate for the 2 neb windows.  one 
            //with (spawned)
            try
            {
                MainWindow L = new MainWindow();
                MainWindowName = "Nebulosity";
                IntPtr hWnd2 = SearchForWindow("wxWindow", MainWindowName);
               L.Log("Neb Handle Found -- " + hWnd2.ToInt32());
                _nebhWnd = hWnd2.ToInt32();

                if (_nebhWnd != 0)
                {
                   
                    string NebVersion;
                    //finds the Neb version (the number after the "v")
                    StringBuilder sb = new StringBuilder(1024);
                    SendMessage(_nebhWnd, MainWindow.WM_GETTEXT, 1024, sb);
                    //  GetWindowText(camera, sb, sb.Capacity);
                  //  L.Log("test");
                    L.Log(sb.ToString());//*****^%%$  None of the L.Log Works  ****$%%^
                    NebVersion = sb.ToString();
                    int NebVposNumber = NebVersion.IndexOf("v");
                    string NebVNumberAfter = NebVersion.Substring(NebVposNumber + 1, 1);
                    _nebVNumber = Convert.ToInt16(NebVNumberAfter);
                }
                if (!L.SlaveModeEnabled())//don't need for slave mode
                {
                    IntPtr PHDhwnd2 = SearchForWindow("wxWindow", "PHD2 Guiding");
                    L.Log("PHD Handle Found --  " + PHDhwnd2.ToInt32());
                    _pHDhwnd = PHDhwnd2.ToInt32();

                    string PHDVersion;
                    //finds the Neb version (the number after the "v")
                    StringBuilder sb = new StringBuilder(1024);
                    SendMessage(_pHDhwnd, MainWindow.WM_GETTEXT, 1024, sb);
                    //  GetWindowText(camera, sb, sb.Capacity);
                    //  L.Log("test");
                    L.Log(sb.ToString());//*****^%%$  None of the L.Log Works  ****$%%^
                    PHDVersion = sb.ToString();
                    int PHDVposNumber = PHDVersion.IndexOf("g");
                    string NebVNumberAfter = PHDVersion.Substring(PHDVposNumber + 2, 1);
                    _phdVNumber = Convert.ToInt16(NebVNumberAfter);

                }
                // PHDhwnd = FindWindow(null, "PHD Guiding 1.13.0b  -  www.stark-labs.com (Log active)");
                _loadScripthwnd = FindWindow(null, "Load script");
                //   Log("PHD " + PHDhwnd.ToString());
                //   Log("Load script " + LoadScripthwnd.ToString());



            }

            catch (Exception e)
            {
              //  MainWindow L = new MainWindow();
             //  L.Log("FindHandles Error" + e.ToString());
        //       L.Send("FindHandles Error" + e.ToString());
             //  L.FileLog("FindHandles Error" + e.ToString());

            }

        }
        bool _setupWindowFound = false;
        public int EnumChildGetValue(int hWnd, int lParam)
        {
            
         
            MainWindow L = new MainWindow();
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
                    _panelhwnd = hWnd;

                }
                if (classtext == "Edit")
                {
                    _editfound++;
                    if (_editfound == 2)
                    {
                        _hwndDuration = hWnd;
                    }
                }

                if (editText == "Advanced")
                {
                    int Advancedhwnd = hWnd;
                }
                //*****************************need to test after writing auto camera find stuff ***********************************         
                if (editText == _nebCamera + " Setup")
                {
                    if (_setupWindowFound == false)//picks first one
                    {
                        int Advhwnd = hWnd;
                        //   Log("Adv" + Advhwnd.ToString());
                        _setupWindowFound = true;
                    }
                }
                if (editText == "Capture Series")
                    _captureMainhWnd = hWnd;
                if (editText == "Abort")
                {

                    _aborthwnd = hWnd;
                    //   Log("Abort " + Aborthwnd.ToString());
                }
                if (editText == "Frame and Focus")
                {

                    _framehwnd = hWnd;
                    // Log("Frame " + Framehwnd.ToString());
                }
                if (editText == "Fine Focus")
                {

                    _finehwnd = hWnd;
                    //  Log("Fine " + Finehwnd.ToString());
                }
                if (editText == "Load script")//not finding it
                {
                    _loadScripthwnd = hWnd;
                    //   Log("LoadScript " + LoadScripthwnd.ToString());
                }
            }
            if (lParam == 2)//added to find handles of second instence of Neb
            {
                if (editText == "panel")//doesn't work w/ msctls_statusbar32  either
                {
                    _panelhwnd2 = hWnd;

                }
                if (classtext == "Edit")
                {
                    _editfound++;
                    if (_editfound == 2)
                    {
                        _hwndDuration2 = hWnd;
                    }
                }

                if (editText == "Advanced")
                {
                    int Advancedhwnd2 = hWnd;
                    //    Advancedhwnd2 = hWnd;
                }
                //*****************************need to test after writing auto camera find stuff ***********************************         
                if (editText == _nebCamera + " Setup")
                {
                    if (_setupWindowFound == false)//picks first one
                    {
                        int Advhwnd2 = hWnd;
                        //   Log("Adv" + Advhwnd.ToString());
                        _setupWindowFound = true;
                    }
                }

                if (editText == "Abort")
                {

                    _aborthwnd2 = hWnd;
                    //   Log("Abort " + Aborthwnd.ToString());
                }
                if (editText == "Frame and Focus")
                {

                    _framehwnd2 = hWnd;
                    // Log("Frame " + Framehwnd.ToString());
                }
                if (editText == "Fine Focus")
                {

                    _finehwnd2 = hWnd;
                    //  Log("Fine " + Finehwnd.ToString());
                }
                if (editText == "Load script")//not finding it
                {
                    _loadScripthwnd2 = hWnd;
                    //   Log("LoadScript " + LoadScripthwnd.ToString());
                }
            }

            if (L.ServerEnabled())
            {
                if (editText == "GotoFocus")
                {
                    _gotofocushwnd = hWnd;
                  //  Log("Slave Gotofocus handle found  " + _gotofocushwnd.ToString());
                }
                if (editText == "Capture")
                {
                    _capturehwnd = hWnd;
                //    Log("Slave Capture Handle Found  " + _capturehwnd.ToString());

                }
                if (editText == "Flat")
                {
                    _flathwnd = hWnd;
                  //  Log("Slave Flat handle found  " + _flathwnd.ToString());
                }
                if (editText == "Pause")
                {
                    _pausehwnd = hWnd;
                  //  Log("SlavePause Handle found  " + _pausehwnd.ToString());
                }
                if (editText == "Not Connected")
                {
                    _slaveStatushwnd = hWnd;
                  //  Log("Status Handle " + _slaveStatushwnd.ToString());
                }

            }

            



            //MessageBox.Show("Contains text of control "+ editText);
            return 1;

        }






    }
}
