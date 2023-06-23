using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class formComments : Form
    {

        private TextBox txtBox = new TextBox();
        private Button btnAdd = new Button();
        private ListBox lstBox = new ListBox();
        private CheckBox chkBox = new CheckBox();
        private Label lblCount = new Label();
        public formComments()
        {
            InitializeComponent();
        }

        private void formComments_Load(object sender, EventArgs e)
        {

            //Set up the form.
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;
            this.ForeColor = Color.Black;
            this.Size = new System.Drawing.Size(155, 265);
            this.Text = "Run-time Controls";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;


            //Format controls. Note: Controls inherit color from parent form.
            this.btnAdd.BackColor = Color.Gray;
            this.btnAdd.Text = "Add";
            this.btnAdd.Location = new System.Drawing.Point(90, 25);
            this.btnAdd.Size = new System.Drawing.Size(50, 25);
            //Controls.

            this.txtBox.Text = "Text";
            this.txtBox.Location = new System.Drawing.Point(10, 25);
            this.txtBox.Size = new System.Drawing.Size(70, 20);

            this.lstBox.Items.Add("One");
            this.lstBox.Items.Add("Two");
            this.lstBox.Items.Add("Three");
            this.lstBox.Items.Add("Four");
            this.lstBox.Sorted = true;
            this.lstBox.Location = new System.Drawing.Point(10, 55);
            this.lstBox.Size = new System.Drawing.Size(130, 95);

            this.chkBox.Text = "Disable";
            this.chkBox.Location = new System.Drawing.Point(15, 190);
            this.chkBox.Size = new System.Drawing.Size(110, 30);

            this.lblCount.Text = lstBox.Items.Count.ToString() + " items";
            this.lblCount.Location = new System.Drawing.Point(55, 160);
            this.lblCount.Size = new System.Drawing.Size(65, 15);

            //Add controls to the form.
            this.Controls.Add(btnAdd);
            this.Controls.Add(txtBox);
            this.Controls.Add(lstBox);
            this.Controls.Add(chkBox);
            this.Controls.Add(lblCount);

        }


    }
}
