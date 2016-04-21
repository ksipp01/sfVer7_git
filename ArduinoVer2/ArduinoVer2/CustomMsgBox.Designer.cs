namespace Pololu.Usc.ScopeFocus
{
    partial class CustomMsgBox
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.TesBtn = new System.Windows.Forms.Button();
            this.NoBtn = new System.Windows.Forms.Button();
            this.IgnorBtn = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(288, 64);
            this.panel1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "label1";
            // 
            // TesBtn
            // 
            this.TesBtn.Location = new System.Drawing.Point(12, 70);
            this.TesBtn.Name = "TesBtn";
            this.TesBtn.Size = new System.Drawing.Size(75, 23);
            this.TesBtn.TabIndex = 1;
            this.TesBtn.Text = "button1";
            this.TesBtn.UseVisualStyleBackColor = true;
            this.TesBtn.Click += new System.EventHandler(this.button1_Click);
            // 
            // NoBtn
            // 
            this.NoBtn.Location = new System.Drawing.Point(109, 70);
            this.NoBtn.Name = "NoBtn";
            this.NoBtn.Size = new System.Drawing.Size(75, 23);
            this.NoBtn.TabIndex = 2;
            this.NoBtn.Text = "button2";
            this.NoBtn.UseVisualStyleBackColor = true;
            this.NoBtn.Click += new System.EventHandler(this.button2_Click);
            // 
            // IgnorBtn
            // 
            this.IgnorBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.IgnorBtn.Location = new System.Drawing.Point(201, 70);
            this.IgnorBtn.Name = "IgnorBtn";
            this.IgnorBtn.Size = new System.Drawing.Size(75, 23);
            this.IgnorBtn.TabIndex = 3;
            this.IgnorBtn.Text = "button3";
            this.IgnorBtn.UseVisualStyleBackColor = true;
            this.IgnorBtn.Click += new System.EventHandler(this.button3_Click);
            // 
            // CustomMsgBox
            // 
            this.AcceptButton = this.TesBtn;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.IgnorBtn;
            this.ClientSize = new System.Drawing.Size(288, 101);
            this.Controls.Add(this.IgnorBtn);
            this.Controls.Add(this.NoBtn);
            this.Controls.Add(this.TesBtn);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CustomMsgBox";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CustomMsgBox";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button TesBtn;
        private System.Windows.Forms.Button NoBtn;
        private System.Windows.Forms.Button IgnorBtn;
    }
}