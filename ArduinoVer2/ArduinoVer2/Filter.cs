using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pololu.Usc.ScopeFocus
{
    class Filter
    {
        private int currentFilter;
        public int Currentfilter { get; set; }

        private void filteradvance()
        {
            
            try
            {
                MainWindow m = new MainWindow();
               MainWindow.FlatCalcDone = false;
                if (MainWindow.SequenceRunning == true)
                {
                   m.DisableUpDwn();
                    // fileSystemWatcher1.EnableRaisingEvents = true;
                }
                toolStripStatusLabel1.Text = "Filter Moving";
                this.Refresh();
                string movedone;
                port.DiscardOutBuffer();
                port.DiscardInBuffer();
                Thread.Sleep(20);
                port.Write("A");
                filterMoving = true;
                Thread.Sleep(50);

                while (filterMoving == true)
                {
                    movedone = port.ReadExisting();
                    if (movedone == "D")
                    {

                        filterMoving = false;

                        toolStripStatusLabel1.Text = "Capturing";
                        this.Refresh();
                        //   toolStripStatusLabel1.Text = " ";
                        //    toolStripStatusLabel1.Text = " ";

                    }
                }
                Thread.Sleep(50);
                port.DiscardInBuffer();

                if (currentfilter == 0)//allows advance without count to set to starting pos. 
                {
                    return;
                }
                /* *********************rem'd 4-11 put currentfilter asignment in filtersequence()
                // if (checkBox5.Checked == false)
                //  {
                if (currentfilter != 7)//***************changed from 4 on 2_29 to ? fix dark2
                {
                    currentfilter++;
                }
                else
                    currentfilter = 1;//reset at end
*/
                //    toolStripStatusLabel1.Text = " ";
                DisplayCurrentFilter();
            }
            catch
            {
                Log("Filter Advance Error - Make Sure Arduino Connected");
                Send("Filter Advance Error - Make Sure Arduino Connected");
                FileLog("Filter Advance Error - Make Sure Arduino Connected");

            }
        }

        public void DisplayCurrentFilter()
        {
            try
            {
                //  if (SequenceRunning == true)
                //  DisableUpDwn();
                //    }
                /*
                            if (checkBox5.Checked == true)
                            {
                                if (currentfilter != 5)
                                {
                                    currentfilter++;
                                }
                                else
                                    currentfilter = 1;
                            }
                 */
                if (currentfilter == 1)
                    Filtertext = comboBox2.Text;
                if (currentfilter == 2)
                    Filtertext = comboBox3.Text;
                if (currentfilter == 3)
                    Filtertext = comboBox4.Text;
                if (currentfilter == 4)
                    Filtertext = comboBox5.Text;
                if (DarksOn == true)
                {
                    Filtertext = "Dark";
                    textBox21.BackColor = System.Drawing.Color.Black;
                    textBox21.ForeColor = System.Drawing.Color.White;
                    toolStripStatusLabel4.ForeColor = System.Drawing.Color.White;
                    toolStripStatusLabel4.BackColor = System.Drawing.Color.Black;
                }
                if (FlatsOn == true)//****************added 4_5 for flats
                {
                    Filtertext = "Flat";
                    textBox21.BackColor = System.Drawing.Color.White;
                    textBox21.ForeColor = System.Drawing.Color.Blue;
                    toolStripStatusLabel4.ForeColor = System.Drawing.Color.Blue;
                    toolStripStatusLabel4.BackColor = System.Drawing.Color.White;
                }
                if (Filtertext == "Blue")
                {
                    textBox21.BackColor = System.Drawing.Color.Blue;
                    textBox21.ForeColor = System.Drawing.Color.White;
                    toolStripStatusLabel4.ForeColor = System.Drawing.Color.White;
                    toolStripStatusLabel4.BackColor = System.Drawing.Color.Blue;
                }
                /*
                        if ((Filtertext != "Blue") || (Filtertext != "Dark"))
                            textBox21.ForeColor = System.Drawing.Color.White;
                */
                if ((DarksOn == false) && (Filtertext != "Blue"))
                {
                    textBox21.ForeColor = System.Drawing.Color.Black;
                    toolStripStatusLabel4.ForeColor = System.Drawing.Color.Black;
                }

                if (Filtertext == "Luminosity")
                {
                    textBox21.BackColor = System.Drawing.Color.White;
                    toolStripStatusLabel4.BackColor = System.Drawing.Color.White;

                }
                if (Filtertext == "Green")
                {
                    textBox21.BackColor = System.Drawing.Color.Lime;
                    toolStripStatusLabel4.BackColor = System.Drawing.Color.Lime;
                }

                if (Filtertext == "Red")
                {
                    textBox21.BackColor = System.Drawing.Color.Red;
                    toolStripStatusLabel4.BackColor = System.Drawing.Color.Red;
                }


                if (Filtertext == "Ha")
                {
                    textBox21.BackColor = System.Drawing.Color.Magenta;
                    toolStripStatusLabel4.BackColor = System.Drawing.Color.Magenta;
                }
                if (Filtertext == "OIII")
                {
                    textBox21.BackColor = System.Drawing.Color.Cyan;
                    toolStripStatusLabel4.BackColor = System.Drawing.Color.Cyan;
                }
                if (Filtertext == "SII")
                {
                    textBox21.BackColor = System.Drawing.Color.Yellow;
                    toolStripStatusLabel4.BackColor = System.Drawing.Color.Yellow;
                }
                //Added to sync filter with equip
                if ((Filtertext == "Ha") || (Filtertext == "OIII") || (Filtertext == "SII"))
                    GlobalVariables.EquipPrefix = comboBox7.Text + "_" + Filtertext;
                if (radioButton9.Checked == true)
                    _equip = GlobalVariables.EquipPrefix + "_E";
                if (radioButton10.Checked == true)
                    _equip = GlobalVariables.EquipPrefix + " _R";
                else
                    _equip = GlobalVariables.EquipPrefix;
                if ((Filtertext == "Luminosity")) //  rem'd 4_19 || (Filtertext == "Red") || (Filtertext == "Green") || (Filtertext == "Blue"))
                //only need this to start at lum w/ NB
                {
                    GlobalVariables.EquipPrefix = comboBox7.Text + "_" + "LRGB";
                    if (radioButton9.Checked == true)
                        _equip = GlobalVariables.EquipPrefix + "_E";
                    if (radioButton10.Checked == true)
                        _equip = GlobalVariables.EquipPrefix + " _R";
                    else
                        _equip = GlobalVariables.EquipPrefix;
                }
                if (checkBox22.Checked == true)
                    _equip = _equip + "_Metric";
                toolStripStatusLabel3.Text = _equip.ToString();
                //end filter/equip sync addition
                toolStripStatusLabel4.Text = Filtertext.ToString();
                toolStripStatusLabel2.Text = "Sub " + subCountCurrent.ToString() + "of " + TotalSubs().ToString();
                //next line filters gridview by equip
                //   Data d = new Data();
                ((DataTable)this.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";
                //  ((DataTable)d.dataGridView1.DataSource).DefaultView.RowFilter = "Equip =" + "'" + toolStripStatusLabel3.Text.ToString() + "'";
                //FillData();//******These next 2 were rem'd ?????????? 8-20-13
                // GetAvg();
                //   textBox27.Text = Filtertext;
                textBox21.Text = Filtertext.ToString();

            }
            catch
            {
                Log("Display Current Filter Error");
            }
        }



    }
}
