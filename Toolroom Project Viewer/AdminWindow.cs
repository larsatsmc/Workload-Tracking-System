using ClassLibrary;
using ClassLibrary.Models;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Toolroom_Project_Viewer
{
    public partial class AdminWindow : Form
    {
        public BindingList<UserModel> UserList { get; set; }
        public AdminWindow()
        {
            InitializeComponent();
        }
        private void DeleteSelectedUser()
        {
            
        }
        private void AdminWindow_Load(object sender, EventArgs e)
        {
            UserList = new BindingList<UserModel>(Database.GetUsers());
            gridControl1.DataSource = UserList;
        }

        private void userGridView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridView gridView = sender as GridView;

            UserModel user = gridView.GetFocusedRow() as UserModel;

            try
            {
                if (e.Column.FieldName == "IsAdmin")
                {
                    if (user.IsAdmin)
                    {
                        user.CanReadOnly = false;
                        user.CanChangeProjectData = true;
                        user.CanChangeDates = true;
                        user.CanCreateProjects = true;
                        user.CanDeleteProjects = true;
                    }
                }
                else if (e.Column.FieldName == "CanReadOnly")
                {
                    if (user.CanReadOnly)
                    {
                        user.IsAdmin = false;
                        user.CanOnlyChangeDesignWork = false;
                        user.CanChangeDates = false;
                        user.CanChangeProjectData = false;
                        user.CanCreateProjects = false;
                        user.CanDeleteProjects = false;
                    }
                }

                if (!gridView.IsNewItemRow(e.RowHandle))
                {
                    Database.UpdateUser(user, e);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }

        private void userGridView_RowUpdated(object sender, DevExpress.XtraGrid.Views.Base.RowObjectEventArgs e)
        {
            GridView gridView = sender as GridView;

            UserModel user = e.Row as UserModel;

            try
            {
                if (gridView.IsNewItemRow(e.RowHandle))
                {
                    Database.CreateUser(user);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }

        private void userGridView_ValidateRow(object sender, DevExpress.XtraGrid.Views.Base.ValidateRowEventArgs e)
        {
            GridView gridView = sender as GridView;

            UserModel user = e.Row as UserModel;

            if (gridView.IsNewItemRow(e.RowHandle))
            {
                if (UserList.Count(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.IsAdmin == true) == 0)
                {
                    MessageBox.Show("You must be an admin to add users.");
                    e.Valid = false;
                    return;
                }

                if (user.FirstName == null || user.FirstName.Length == 0)
                {
                    gridView.SetColumnError(colFirstName, "Please enter a first name.");
                    e.Valid = false;
                }
                else if (user.LastName == null || user.LastName.Length == 0)
                {
                    gridView.SetColumnError(colLastName, "Please enter a last name.");
                    e.Valid = false;
                }
                else if (user.LoginName == null || user.LoginName.Length == 0)
                {
                    gridView.SetColumnError(colLoginName, "Please enter a login name.");
                    e.Valid = false;
                }
            }
        }

        private void userGridView_ValidatingEditor(object sender, DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs e)
        {
            GridView gridView = sender as GridView;

            GridColumn column = (e as EditFormValidateEditorEventArgs)?.Column ?? gridView.FocusedColumn;

            if (UserList.Count(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.IsAdmin == true) == 0)
            {
                e.ErrorText = "You must be an admin to make changes. Hit ESC to Cancel.";
                e.Valid = false;
            }
        }

        private void userGridView_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Delete && e.Modifiers == Keys.Control)
                {
                    if (UserList.Count(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.IsAdmin == true) == 0)
                    {
                        MessageBox.Show("You must be an admin to delete users.");
                        return;
                    }

                    UserModel user = userGridView.GetFocusedRow() as UserModel;

                    if (Database.DeleteUser(user.ID))
                    {
                        userGridView.DeleteRow(userGridView.FocusedRowHandle);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
