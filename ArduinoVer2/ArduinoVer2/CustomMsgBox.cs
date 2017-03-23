using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pololu.Usc.ScopeFocus
{
    public partial class CustomMsgBox : Form
    {
        public CustomMsgBox(string Text)
        {
            InitializeComponent();
            this.label1.Text = Text;
            //MsgBox.TesBtn.Text = YesBtn;
            //MsgBox.NoBtn.Text = NoBtn;
            //MsgBox.IgnorBtn.Text = IgnoreBtn;
            //MsgBox.Text = Caption;
            //MsgBox.button4.Text = button4;
            //MsgBox.button5.Text = button5;
        }

      

        private static bool _mountSlew = false;
        public static bool MountSlew
        {
            get { return _mountSlew; }
            set { _mountSlew = value; }
        }
        private static bool _mountSync = false;
        public static bool MountSync
        {
            get { return _mountSync; }
            set { _mountSync = value; }
        }
        private static bool _rotatorMove = false;
        public static bool RotatorMove
        {
            get { return _rotatorMove; }
            set { _rotatorMove = value; }
        }
        private static bool _rotatorSnyc = false;
        public static bool RotatorSync
        {
            get { return _rotatorSnyc; }
            set { _rotatorSnyc = value; }
        }
        private static bool _close = false;
        public static bool Close
        {
            get { return _close; }
            set { _close = value; }
        }
        //public string LabelText
        //{
        //    get {return this.label1.Text;}
        //    set
        //    {this.label1.Text = value;}
        //}


     //   static CustomMsgBox MsgBox;
        //static DialogResult result = DialogResult.No;
        //public static DialogResult Show (string Text, string Caption, string YesBtn, string NoBtn, string IgnoreBtn, string button4, string button5)
        //{
        //    MsgBox = new CustomMsgBox();
        //    MsgBox.label1.Text = Text;
        //    MsgBox.TesBtn.Text = YesBtn;
        //    MsgBox.NoBtn.Text = NoBtn;
        //    MsgBox.IgnorBtn.Text = IgnoreBtn;
        //    MsgBox.Text = Caption;
        //    MsgBox.button4.Text = button4;
        //    MsgBox.button5.Text = button5;
        // //   result = DialogResult.No;
        //    MsgBox.ShowDialog();
        //    return result;
        //}

      
        private void button1_Click(object sender, EventArgs e)
        {
            _mountSlew = true;

           // result = DialogResult.Yes;// MsgBox.Close();  // Mount slew
        }

        private void button3_Click(object sender, EventArgs e)
        {
            _close = true;
            //   result = DialogResult.No;
         //   MsgBox.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {

            _mountSync = true;
            // result = DialogResult.Ignore; MsgBox.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            _rotatorMove = true;
            
            // result = DialogResult.OK;  // Rotator Go To
        }

        private void button5_Click(object sender, EventArgs e)
        {
            _rotatorSnyc = true;
            //  result = DialogResult.Retry; // Rotator Snyc
        }

        private void CustomMsgBox_Load(object sender, EventArgs e)
        {
            if (!MainWindow.RotatorIsConnected)
            {
                button4.Enabled = false;
                button5.Enabled = false;
            }

        }
    }
}
