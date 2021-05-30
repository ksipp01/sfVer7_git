using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ASCOM.DriverAccess;

namespace Pololu.Usc.ScopeFocus
{
    class Mount
    {

        public static Telescope scope;
        private static string devId;
        public static string DevId
        {
            get { return devId; }
            set { devId = value; }
        }


        public static void Chooser()
        {
            ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
            chooser.DeviceType = "Telescope";
            devId = chooser.Choose();
            if (devId != "")
                scope = new ASCOM.DriverAccess.Telescope(devId);
            else
                return;
            //  ASCOM.DriverAccess.Telescope scope = new ASCOM.DriverAccess.Telescope(devId);
            //    Log("connected to " + devId);
            //    FileLog2("connected to " + devId);
                scope.Connected = true;
            //    if (scope.Connected)
            //    {
            //        timer2.Enabled = true;
            //        timer2.Start();
            //    }
            //    usingASCOM = true;
            //    button49.BackColor = System.Drawing.Color.Lime;
            //    //   button49.Text = "Connected";
            //}
            //  bool usingASCOMFocus = false;
        }


    }
}
