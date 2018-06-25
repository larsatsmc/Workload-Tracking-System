using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Toolroom_Scheduler
{
    public partial class AddResourceForm : Form
    {
        public AddResourceForm()
        {
            InitializeComponent();
        }

        private void AddResourcesToDB()
        {
            if (((FirstNameTextBox.Text == "" || FirstNameTextBox.Text == "First Name") && (LastNameTextBox.Text == "" || LastNameTextBox.Text == "Last Name")) && MachineNameTextBox.Text == "")
            {
                MessageBox.Show("Please enter a person's full name or a machine name.");

                if (FirstNameTextBox.Text == "First Name" || FirstNameTextBox.Text == "")
                {
                    FirstNameTextBox.Text = "";
                    FirstNameTextBox.BackColor = Color.Red;
                }
                if (LastNameTextBox.Text == "Last Name" || LastNameTextBox.Text == "")
                {
                    LastNameTextBox.Text = "";
                    LastNameTextBox.BackColor = Color.Red;
                }
            }
            else if(MachineNameTextBox.Text == "")
            {

            }
        }

        private void FirstNameTextBox_Click(object sender, EventArgs e)
        {
            FirstNameTextBox.Enabled = true;
            LastNameTextBox.Enabled = true;
            MachineNameTextBox.Enabled = false;
        }

        private void LastNameTextBox_Click(object sender, EventArgs e)
        {
            FirstNameTextBox.Enabled = true;
            LastNameTextBox.Enabled = true;
            MachineNameTextBox.Enabled = false;
        }

        private void MachineNameTextBox_Click(object sender, EventArgs e)
        {
            FirstNameTextBox.Enabled = false;
            LastNameTextBox.Enabled = false;
            MachineNameTextBox.Enabled = true;
        }
    }
}
