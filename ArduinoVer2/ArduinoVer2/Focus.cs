using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ASCOM.DriverAccess;

namespace Pololu.Usc.ScopeFocus
{
    class Focus
    {

        public static Focuser focuser;
        private static string devId2;
        public static string DevId2
        {
            get { return devId2; }
            set { devId2 = value; }
        }

        public static void FocusChooser()
        {
            ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
            chooser.DeviceType = "Focuser";
            devId2 = chooser.Choose();
            //  ASCOM.DriverAccess.Focuser focuser = new ASCOM.DriverAccess.Focuser(devId2);
            if (devId2 != "")
                focuser = new ASCOM.DriverAccess.Focuser(devId2);
            else
                return;



            focuser.Connected = true;
            //***************** I think this needs to be changes so it SETs the value not GETs***************
            //go back to previous method for storing the maxtravel in settings, then when gets it 
            //after selecting equipement it sets the value. 
            //****************************************************************************************
            //travel = focuser.MaxStep;
            //textBox2.Text = travel.ToString();
            //count = focuser.Position;
            //textBox1.Text = count.ToString();
            //Log("connected to " + devId2);
            //FileLog2("connected to " + devId2);
            //button8.BackColor = System.Drawing.Color.Lime;
            ////  button8.Text = "Connected";
            ///*
            //if (focuser.Connected)
            //{
            //    MessageBox.Show("ASCOM Focuser connected");
            //}
            // */
            //numericUpDown6.Value = focuser.Position;
            //    usingASCOMFocus = true;

            //   focuser.CommandString("C", true);
        }


    }
}
