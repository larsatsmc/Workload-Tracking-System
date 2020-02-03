using ClassLibrary;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace Toolroom_Project_Viewer
{
    public partial class ManageResourcesForm : DevExpress.XtraEditors.XtraForm
    {
        List<DepartmentModel> Departments = new List<DepartmentModel>();

        public ManageResourcesForm()
        {
            try
            {
                InitializeComponent();

                InitializeDepartments();

                LoadResourceListBox();

                resourceListBox.SelectedIndex = 1;
                resourceListBox.SelectedIndex = 0;

                LoadRoleListBox();

                RoleComboBox.SelectedIndex = 0;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
            }
        }

        private void LoadResourceListBox()
        {
            resourceListBox.SelectedValueChanged -= ResourceListBox_SelectedValueChanged;

            resourceListBox.DataSource = Database.GetResourceList();

            resourceListBox.SelectedValueChanged += ResourceListBox_SelectedValueChanged;
        }

        private void LoadRoleListBox()
        {
            roleListBox.DataSource = null;

            if (RoleComboBox.Text != "")
            {
                roleListBox.DataSource = Database.GetRoleList(GetRoleFromRoleComboBox());
            }
        }

        private void RemoveResource()
        {
            if (resourceListBox.SelectedItems.Count > 0)
            {
                Database.RemoveResource(resourceListBox.SelectedItem.ToString());

                LoadResourceListBox();

                LoadRoleListBox();
            }
            else
            {
                MessageBox.Show("You have not selected a resource to remove.");
            }
        }
        private bool ResourceExists(string resource)
        {
            foreach (var item in resourceListBox.Items)
            {
                if (item.ToString() == resource)
                {
                    MessageBox.Show("That resource already exists.");
                    return true;
                }
            }

            return false;
        }
        private void AddResource()
        {
            try
            {
                if (addResourceTextBox.Text != "" && ResourceExists(addResourceTextBox.Text) == false)
                {
                    Database.InsertResource(addResourceTextBox.Text, GetResourceType());

                    LoadResourceListBox();

                    resourceListBox.SelectedItem = addResourceTextBox.Text;
                    addResourceTextBox.Text = "";
                }
                else if (addResourceTextBox.Text == "")
                {
                    MessageBox.Show("You have not typed in a resource to add.");
                }
            }
            catch (Exception ex)
            {
                resourceListBox.SelectedValueChanged += ResourceListBox_SelectedValueChanged;
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
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
        private string GetResourceType()
        {
            if (resourceTypeRadioGroup.SelectedIndex == 0)
            {
                return "Person";
            }
            else
            {
                return "Machine";
            }
        }
        private void AddRoleToResource()
        {
            if (!RoleComboBox.Items.Contains(RoleComboBox.Text))
            {
                MessageBox.Show("Creation of new role is not allowed.");
                return;
            }

            if (RoleComboBox.Text != "" || resourceListBox.SelectedItems.Count > 0)
            {
                string role = GetRoleFromRoleComboBox();

                Database.InsertResourceRole(resourceListBox.SelectedItem.ToString(), role, FindDepartmentID(role));

                LoadRoleListBox();
            }
        }

        private void RemoveRoleFromResource()
        {
            if (roleListBox.SelectedItems.Count > 0)
            {
                Database.RemoveResourceRole(roleListBox.SelectedItem.ToString(), GetRoleFromRoleComboBox());

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

            if (sb.ToString().EndsWith("s"))
            {
                sb.Remove(sb.Length - 1, 1); 
            }

            return sb.ToString();
        }

        private void InitializeDepartments()
        {
            Departments = Database.LoadDepartments();
        }

        private int FindDepartmentID(string role)
        {
            int departmentID;

            if (role == "Design")
            {
                departmentID = 1;
            }
            else if (role == "Rough Programmer")
            {
                departmentID = 2;
            }
            else if (role == "Finish Programmer")
            {
                departmentID = 3;
            }
            else if (role == "Electrode Programmer")
            {
                departmentID = 4;
            }
            else if (role == "Rough Mill")
            {
                departmentID = 5;
            }
            else if (role == "Rough CNC Operator")
            {
                departmentID = 5;
            }
            else if (role == "Finish Mill")
            {
                departmentID = 6;
            }
            else if (role == "Finish CNC Operator")
            {
                departmentID = 6;
            }
            else if (role == "Graphite Mill")
            {
                departmentID = 7;
            }
            else if (role == "Electrode CNC Operator")
            {
                departmentID = 7;
            }
            else if (role == "CMM Operator")
            {
                departmentID = 8;
            }
            else if (role == "Tool Maker")
            {
                departmentID = 9;
            }
            else if (role == "EDM Wire")
            {
                departmentID = 10;
            }
            else if (role == "EDM Wire Operator")
            {
                departmentID = 10;
            }
            else if (role == "EDM Sinker")
            {
                departmentID = 11;
            }
            else if (role == "EDM Sinker Operator")
            {
                departmentID = 11;
            }
            else if (role == "Hole Popper Operator")
            {
                departmentID = 13;
            }
            else
            {
                departmentID = 12;
            }

            return departmentID;
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

        private void ResourceListBox_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                resourceTypeRadioGroup.EditValueChanged -= ResourceTypeRadioGroup_EditValueChanged;
                resourceTypeRadioGroup.EditValue = Database.GetResourceType(resourceListBox.SelectedItem.ToString());
                resourceTypeRadioGroup.EditValueChanged += ResourceTypeRadioGroup_EditValueChanged;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void ResourceTypeRadioGroup_EditValueChanged(object sender, EventArgs e)
        {
            if (resourceListBox.SelectedItems.Count == 1)
            {
                Database.SetResourceType(resourceListBox.SelectedItem.ToString(), resourceTypeRadioGroup.EditValue.ToString());
            }
            else if (resourceListBox.SelectedItems.Count > 1)
            {
                MessageBox.Show("Please select only one item to change resourceType.");
            }
        }
    }
}
