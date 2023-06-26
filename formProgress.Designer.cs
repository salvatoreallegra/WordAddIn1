
namespace WordAddIn1
{
    partial class formProgress
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
            this.cmeProgress = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // cmeProgress
            // 
            this.cmeProgress.Location = new System.Drawing.Point(313, 165);
            this.cmeProgress.Name = "cmeProgress";
            this.cmeProgress.Size = new System.Drawing.Size(100, 23);
            this.cmeProgress.TabIndex = 0;
            // 
            // formProgress
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.cmeProgress);
            this.Name = "formProgress";
            this.Text = "formProgress";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar cmeProgress;
    }
}