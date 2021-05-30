using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using ASCOM.DriverAccess;

namespace Pololu.Usc.ScopeFocus
{
    class Flap
    {

        //public static Switch FlatFlap;
        public static CoverCalibrator FlatFlap;
        private static string devId4;
        public static string DevId4
        {
            get { return devId4; }
            set { devId4 = value; }
        }
        // changed to coverCal 5/30/2021
        public static void FlapChooser()
        {
            ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
            chooser.DeviceType = "CoverCalibrator";
                devId4 = chooser.Choose();
                if (!string.IsNullOrEmpty(devId4))
                {
                   // FlatFlap = new ASCOM.DriverAccess.Switch(devId4);
                FlatFlap = new ASCOM.DriverAccess.CoverCalibrator(devId4);
                FlatFlap.Connected = true;
                    Thread.Sleep(200);
                  //  FlatFlap.SetSwitch(0, false);
                }
}
}
}
