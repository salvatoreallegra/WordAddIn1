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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyUserControl));
            this.btnCorrectDocument = new System.Windows.Forms.Button();
            this.lblCompliance = new System.Windows.Forms.Label();
            this.picMainImage = new System.Windows.Forms.PictureBox();
            this.btnClearComments = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.picMainImage)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCorrectDocument
            // 
            this.btnCorrectDocument.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(214)))), ((int)(((byte)(233)))));
            this.btnCorrectDocument.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCorrectDocument.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCorrectDocument.Location = new System.Drawing.Point(171, 327);
            this.btnCorrectDocument.Name = "btnCorrectDocument";
            this.btnCorrectDocument.Size = new System.Drawing.Size(138, 60);
            this.btnCorrectDocument.TabIndex = 0;
            this.btnCorrectDocument.Text = "Correct Document";
            this.btnCorrectDocument.UseVisualStyleBackColor = false;
            this.btnCorrectDocument.Click += new System.EventHandler(this.btnCorrectDocument_Click);
            // 
            // lblCompliance
            // 
            this.lblCompliance.AutoSize = true;
            this.lblCompliance.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCompliance.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.lblCompliance.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.lblCompliance.Location = new System.Drawing.Point(79, 17);
            this.lblCompliance.Name = "lblCompliance";
            this.lblCompliance.Size = new System.Drawing.Size(319, 31);
            this.lblCompliance.TabIndex = 1;
            this.lblCompliance.Text = "Compliance Made Easy";
            this.lblCompliance.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // picMainImage
            // 
            this.picMainImage.Image = ((System.Drawing.Image)(resources.GetObject("picMainImage.Image")));
            this.picMainImage.Location = new System.Drawing.Point(55, 129);
            this.picMainImage.Name = "picMainImage";
            this.picMainImage.Size = new System.Drawing.Size(354, 127);
            this.picMainImage.TabIndex = 2;
            this.picMainImage.TabStop = false;
            // 
            // btnClearComments
            // 
            this.btnClearComments.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(214)))), ((int)(((byte)(233)))));
            this.btnClearComments.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClearComments.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClearComments.Location = new System.Drawing.Point(171, 393);
            this.btnClearComments.Name = "btnClearComments";
            this.btnClearComments.Size = new System.Drawing.Size(138, 60);
            this.btnClearComments.TabIndex = 3;
            this.btnClearComments.Text = "Clear all Comments";
            this.btnClearComments.UseVisualStyleBackColor = false;
            this.btnClearComments.Click += new System.EventHandler(this.btnClearComments_Click);
            // 
            // MyUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(42)))), ((int)(((byte)(79)))));
            this.Controls.Add(this.btnClearComments);
            this.Controls.Add(this.picMainImage);
            this.Controls.Add(this.lblCompliance);
            this.Controls.Add(this.btnCorrectDocument);
            this.Name = "MyUserControl";
            this.Size = new System.Drawing.Size(494, 590);
            ((System.ComponentModel.ISupportInitialize)(this.picMainImage)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCorrectDocument;
        private System.Windows.Forms.Label lblCompliance;
        private System.Windows.Forms.PictureBox picMainImage;
        private System.Windows.Forms.Button btnClearComments;
    }
}
