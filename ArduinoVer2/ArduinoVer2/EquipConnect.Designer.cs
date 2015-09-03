namespace WindowsFormsApplication1
{
    partial class EquipConnect
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonFocusConnect = new System.Windows.Forms.Button();
            this.buttonMountConnect = new System.Windows.Forms.Button();
            this.buttonFilterConnect = new System.Windows.Forms.Button();
            this.buttonSwitchConnect = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonFocusConnect
            // 
            this.buttonFocusConnect.Location = new System.Drawing.Point(110, 31);
            this.buttonFocusConnect.Name = "buttonFocusConnect";
            this.buttonFocusConnect.Size = new System.Drawing.Size(75, 23);
            this.buttonFocusConnect.TabIndex = 0;
            this.buttonFocusConnect.Text = "Connect";
            this.buttonFocusConnect.UseVisualStyleBackColor = true;
            this.buttonFocusConnect.Click += new System.EventHandler(this.buttonFocusConnect_Click);
            // 
            // buttonMountConnect
            // 
            this.buttonMountConnect.Location = new System.Drawing.Point(110, 60);
            this.buttonMountConnect.Name = "buttonMountConnect";
            this.buttonMountConnect.Size = new System.Drawing.Size(75, 23);
            this.buttonMountConnect.TabIndex = 1;
            this.buttonMountConnect.Text = "Connect";
            this.buttonMountConnect.UseVisualStyleBackColor = true;
            // 
            // buttonFilterConnect
            // 
            this.buttonFilterConnect.Location = new System.Drawing.Point(110, 89);
            this.buttonFilterConnect.Name = "buttonFilterConnect";
            this.buttonFilterConnect.Size = new System.Drawing.Size(75, 23);
            this.buttonFilterConnect.TabIndex = 2;
            this.buttonFilterConnect.Text = "Connect";
            this.buttonFilterConnect.UseVisualStyleBackColor = true;
            // 
            // buttonSwitchConnect
            // 
            this.buttonSwitchConnect.Location = new System.Drawing.Point(110, 118);
            this.buttonSwitchConnect.Name = "buttonSwitchConnect";
            this.buttonSwitchConnect.Size = new System.Drawing.Size(75, 23);
            this.buttonSwitchConnect.TabIndex = 3;
            this.buttonSwitchConnect.Text = "Connect";
            this.buttonSwitchConnect.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(56, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Focuser:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(61, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Mount: ";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(38, 94);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Filter Wheel:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(62, 123);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(42, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Switch:";
            // 
            // EquipConnect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(230, 165);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonSwitchConnect);
            this.Controls.Add(this.buttonFilterConnect);
            this.Controls.Add(this.buttonMountConnect);
            this.Controls.Add(this.buttonFocusConnect);
            this.Name = "EquipConnect";
            this.Text = "EquipConnect";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonFocusConnect;
        private System.Windows.Forms.Button buttonMountConnect;
        private System.Windows.Forms.Button buttonFilterConnect;
        private System.Windows.Forms.Button buttonSwitchConnect;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}