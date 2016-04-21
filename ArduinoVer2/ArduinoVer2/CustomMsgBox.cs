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
        public CustomMsgBox()
        {
            InitializeComponent();
        }
        static CustomMsgBox MsgBox; static DialogResult result = DialogResult.No;
        public static DialogResult Show (string Text, string Caption, string YesBtn, string NoBtn, string IgnoreBtn)
        {
            MsgBox = new CustomMsgBox();
            MsgBox.label1.Text = Text;
            MsgBox.TesBtn.Text = YesBtn;
            MsgBox.NoBtn.Text = NoBtn;
            MsgBox.IgnorBtn.Text = IgnoreBtn;
            MsgBox.Text = Caption;
         //   result = DialogResult.No;
            MsgBox.ShowDialog();
            return result;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            result = DialogResult.Yes; MsgBox.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            result = DialogResult.Ignore; MsgBox.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            result = DialogResult.No; MsgBox.Close();
        }

       
    }
}
