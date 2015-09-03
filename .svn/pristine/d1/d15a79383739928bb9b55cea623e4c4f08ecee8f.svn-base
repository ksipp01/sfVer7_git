using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ASCOM.DeviceInterface;
using ASCOM.Utilities;

namespace ASCOM.CGEM
{
    [ComVisible(false)]					// Form not registered for COM!
    public partial class SetupDialogForm : Form
    {


      //  internal const string driverID = "ASCOM.CGEM.Telescope";
     //   private const string driverDescription = "ASCOM TEST Telescope Driver for CGEM.";
        public SetupDialogForm()
        {
            InitializeComponent();
          //  textBox1.Text = Properties.Settings.Default.CommPort;
        }
        
        private void cmdOK_Click(object sender, EventArgs e)
        {
        using (ASCOM.Utilities.Profile p = new Utilities.Profile())
            {
   
                p.DeviceType = "Telescope";
                p.WriteValue(Telescope.driverID, "Comport",
                (string)comboBox1.SelectedItem);
            }
            
         //   Properties.Settings.Default.CommPort = comboBox1.SelectedItem.ToString();
            Dispose();
            Close();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BrowseToAscom(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("http://ascom-standards.org/");
            }
            catch (System.ComponentModel.Win32Exception noBrowser)
            {
                if (noBrowser.ErrorCode == -2147467259)
                    MessageBox.Show(noBrowser.Message);
            }
            catch (System.Exception other)
            {
                MessageBox.Show(other.Message);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            using (ASCOM.Utilities.Serial serial = new Utilities.Serial())
            {
                foreach (var item in serial.AvailableCOMPorts)
                {
                    comboBox1.Items.Add(item);
                }
            }
        }
       


       
    }
}