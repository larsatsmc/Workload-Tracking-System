using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Toolroom_Scheduler
{
	public partial class ManageResourcesForm : Form
	{
        Database db = new Database();

        public ManageResourcesForm()
		{
			try
			{
				InitializeComponent();

				LoadResourceListBox();

                RoleComboBox.SelectedIndex = 0;
            }
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}
		}

		private void LoadResourceListBox()
		{
            resourceListBox.DataSource = null;

            resourceListBox.DataSource = db.GetResourceList();
		}

        private void LoadRoleListBox()
        {
            roleListBox.DataSource = null;

            if(RoleComboBox.Text != "")
            {
                roleListBox.DataSource = db.GetRoleList(GetRoleFromRoleComboBox());
            }
        }

		private void RemoveResource()
		{
			if (resourceListBox.SelectedItems.Count > 0)
			{
                db.RemoveResource(resourceListBox.SelectedItem.ToString());

                LoadResourceListBox();
                LoadRoleListBox();
            }
            else
			{
                MessageBox.Show("You have not selected a resource to remove.");
            }
        }

		private void AddResource()
		{
            if (addResourceTextBox.Text != "")
            {
                db.InsertResource(addResourceTextBox.Text);

                LoadResourceListBox();

                resourceListBox.SelectedItem = addResourceTextBox.Text;
            }
            else
            {
                MessageBox.Show("You have not typed in a resource to add.");
            }

            // Maybe remove window from program as well?

            //Database di = new Database();
            //using (var form = new AddResourceForm())
            //{
            //    Start:
            //    var result = form.ShowDialog();
            //    if (result == DialogResult.OK)
            //    {

            //        goto Start;
            //    }
            //    else if (result == DialogResult.Cancel)
            //    {
            //        form.Close();
            //    }
            //}
        }

        private void AddRoleToResource()
        {
            if(RoleComboBox.Text != "" || resourceListBox.SelectedItems.Count > 0)
            {
                db.InsertResourceRole(resourceListBox.SelectedItem.ToString(), GetRoleFromRoleComboBox());

                LoadRoleListBox();
            }
        }

        private void RemoveRoleFromResource()
        {
            if(roleListBox.SelectedItems.Count > 0)
            {
                db.RemoveResourceRole(roleListBox.SelectedItem.ToString(), GetRoleFromRoleComboBox());

                LoadRoleListBox();
            }
            else
            {
                MessageBox.Show("You have not selected a resource to remove a role from.");
            }
        }

        private string GetRoleFromRoleComboBox()
        {
            StringBuilder sb = new StringBuilder(RoleComboBox.Text);

            sb.Remove(sb.Length - 1, 1);

            return sb.ToString();
        }

		private void RemoveButton_Click(object sender, EventArgs e)
		{
			RemoveResource();
		}

		private void RoleComboBox_SelectedIndexChanged(object sender, EventArgs e)
		{
            LoadRoleListBox();
		}

        private void AddRoleButton_Click(object sender, EventArgs e)
        {
            AddRoleToResource();
        }

        private void RemoveRoleButton_Click(object sender, EventArgs e)
        {
            RemoveRoleFromResource();
        }

        private void AddResourceButton_Click(object sender, EventArgs e)
        {
            AddResource();
        }

        private void RemoveResourceButton_Click(object sender, EventArgs e)
        {
            RemoveResource();
        }
    }
}
