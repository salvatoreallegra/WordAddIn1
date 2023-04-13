namespace WordAddIn1
{
    partial class MyUserControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnExecuteStyles = new System.Windows.Forms.Button();
            this.lblCompliance = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnExecuteStyles
            // 
            this.btnExecuteStyles.Location = new System.Drawing.Point(175, 370);
            this.btnExecuteStyles.Name = "btnExecuteStyles";
            this.btnExecuteStyles.Size = new System.Drawing.Size(138, 23);
            this.btnExecuteStyles.TabIndex = 0;
            this.btnExecuteStyles.Text = "Find and Replace";
            this.btnExecuteStyles.UseVisualStyleBackColor = true;
            this.btnExecuteStyles.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblCompliance
            // 
            this.lblCompliance.AutoSize = true;
            this.lblCompliance.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCompliance.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.lblCompliance.Location = new System.Drawing.Point(79, 17);
            this.lblCompliance.Name = "lblCompliance";
            this.lblCompliance.Size = new System.Drawing.Size(319, 31);
            this.lblCompliance.TabIndex = 1;
            this.lblCompliance.Text = "Complaince Made Easy";
            // 
            // MyUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(42)))), ((int)(((byte)(79)))));
            this.Controls.Add(this.lblCompliance);
            this.Controls.Add(this.btnExecuteStyles);
            this.Name = "MyUserControl";
            this.Size = new System.Drawing.Size(500, 500);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExecuteStyles;
        private System.Windows.Forms.Label lblCompliance;
    }
}
