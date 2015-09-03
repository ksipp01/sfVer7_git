using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using WindowsFormsApplication1.Properties;

namespace Pololu.Usc.ScopeFocus
{
    public partial class Form1 : Form
    {

        string str0 = WindowsFormsApplication1.Properties.Settings.Default.setup1;
        string str1 = WindowsFormsApplication1.Properties.Settings.Default.setup2;
        string str2 = WindowsFormsApplication1.Properties.Settings.Default.setup3;
        string str3 = WindowsFormsApplication1.Properties.Settings.Default.setup4;
        string str4 = WindowsFormsApplication1.Properties.Settings.Default.setup5;
        string str5 = WindowsFormsApplication1.Properties.Settings.Default.setup6;


       
        public Form1()
        {
            InitializeComponent();
            
        }

        MainMenu secondform = new MainMenu();
        public static TextBox equip = new TextBox();




        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        
        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            if (this.listView1.SelectedItems.Count == 1)
            {
                this.listView1.SelectedItems[0].BeginEdit();
            }

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
       

        private void Form1_Load(object sender, EventArgs e)
        {
            equip = textBox1;
            string [] str = new string[6];
            str[0] = WindowsFormsApplication1.Properties.Settings.Default.setup1;
            str[1] = WindowsFormsApplication1.Properties.Settings.Default.setup2;
            str[2] = WindowsFormsApplication1.Properties.Settings.Default.setup3;
            str[3] = WindowsFormsApplication1.Properties.Settings.Default.setup4;
            str[4] = WindowsFormsApplication1.Properties.Settings.Default.setup5;
            str[5] = WindowsFormsApplication1.Properties.Settings.Default.setup6;
               
            for (int i = 0; i < 6; i++)
            {
                listView1.Items.Insert(i,str[i] , str[i], 0);
            }


        }
     

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            string test0 = this.listView1.Items[0].Text.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.setup1 = test0.ToString();
           
            string test1 = this.listView1.Items[1].Text.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.setup2 = test1.ToString();
          
            string test2 = this.listView1.Items[2].Text.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.setup3 = test2.ToString();
           
            string test3 = this.listView1.Items[3].Text.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.setup4 = test3.ToString();
          
            string test4 = this.listView1.Items[4].Text.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.setup5 = test4.ToString();
           
            string test5 = this.listView1.Items[5].Text.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.setup6 = test5.ToString();
            WindowsFormsApplication1.Properties.Settings.Default.Save();

            
          //  int index  = Convert.ToInt32(listView1.SelectedIndices);
       //     equipment = listView1.Items[index].ToString();
          
           
                 
            
        }

        private void listView1_AfterLabelEdit(object sender, LabelEditEventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
