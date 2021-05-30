using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ASCOM.DriverAccess;

namespace Pololu.Usc.ScopeFocus
{
    public class Filter
    {
       public static FilterWheel filterWheel;
        private static string devId3;
        public static string DevId3
        {
            get { return devId3; }
            set { devId3 = value; }
        }

        public static void FilterWheelChooser()
        {

            ASCOM.Utilities.Chooser chooser = new ASCOM.Utilities.Chooser();
            chooser.DeviceType = "FilterWheel";
            devId3 = chooser.Choose();
            //  ASCOM.DriverAccess.Focuser focuser = new ASCOM.DriverAccess.Focuser(devId2);
            if (devId3 != "")
                filterWheel = new FilterWheel(devId3);
            else
                return;
            filterWheel.Connected = true;
            //Log("filterwheel connected " + devId3);
            //FileLog2("filterwheel connected " + devId3);
            //buttonFilterConnect.BackColor = System.Drawing.Color.Lime;
            //if (!checkBox31.Checked)
            //    filterWheel.Position = 0;
            //DisplayCurrentFilter();
            //if (!checkBox31.Checked)
            //{
            //    foreach (string filter in filterWheel.Names)
            //        comboBox6.Items.Add(filter);
            //    comboBox6.SelectedItem = filterWheel.Position;
            //    ComboBoxFill();
            
        }

    }
}
