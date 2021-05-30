using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using ASCOM.DriverAccess;

namespace Pololu.Usc.ScopeFocus
{
    class Rot
    {
        public static Rotator Rotate;
        private static string devId5;
        public static string DevId5
        {
            get { return devId5; }
            set { devId5 = value; }
        }
        private static float skyAngleCorrection;
        public static float SkyAngleCorrection
        {
            get { return skyAngleCorrection; }
            set { skyAngleCorrection = value; }
        }
        public static void Chooser()
        {

            ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
            chooser.DeviceType = "Rotator";
            devId5 = chooser.Choose();
            //  ASCOM.DriverAccess.Focuser focuser = new ASCOM.DriverAccess.Focuser(devId2);
            if (devId5 != "")
                Rotate = new ASCOM.DriverAccess.Rotator(devId5);
            else
                return;
            Rotate.Connected = true;

            //ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
            //chooser.DeviceType = "Rotator";
            //devId5 = chooser.Choose();
            //if (!string.IsNullOrEmpty(devId5))
            //{
            //    Rotate = new ASCOM.DriverAccess.Rotator(devId5);

            //    Rotate.Connected = true;
            //    Thread.Sleep(200);

            //}


        }
    }
}
