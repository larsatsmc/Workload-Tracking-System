namespace Toolroom_Project_Viewer
{
    partial class AdminWindow
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
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.userGridView = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colFirstName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colLastName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colLoginName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colIsAdmin = new DevExpress.XtraGrid.Columns.GridColumn();
            this.riCheckEdit = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
            this.colCanChangeDates = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colEngineeringNumberVisible = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colCanReadOnly = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colCanOnlyChangeDesignWork = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colCanChangeProjectData = new DevExpress.XtraGrid.Columns.GridColumn();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.userGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.riCheckEdit)).BeginInit();
            this.SuspendLayout();
            // 
            // gridControl1
            // 
            this.gridControl1.Location = new System.Drawing.Point(12, 12);
            this.gridControl1.MainView = this.userGridView;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.riCheckEdit});
            this.gridControl1.Size = new System.Drawing.Size(884, 555);
            this.gridControl1.TabIndex = 0;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.userGridView});
            // 
            // userGridView
            // 
            this.userGridView.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colFirstName,
            this.colLastName,
            this.colLoginName,
            this.colIsAdmin,
            this.colCanChangeDates,
            this.colEngineeringNumberVisible,
            this.colCanReadOnly,
            this.colCanOnlyChangeDesignWork,
            this.colCanChangeProjectData});
            this.userGridView.GridControl = this.gridControl1;
            this.userGridView.Name = "userGridView";
            this.userGridView.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.True;
            this.userGridView.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Top;
            this.userGridView.CellValueChanged += new DevExpress.XtraGrid.Views.Base.CellValueChangedEventHandler(this.userGridView_CellValueChanged);
            this.userGridView.ValidateRow += new DevExpress.XtraGrid.Views.Base.ValidateRowEventHandler(this.userGridView_ValidateRow);
            this.userGridView.RowUpdated += new DevExpress.XtraGrid.Views.Base.RowObjectEventHandler(this.userGridView_RowUpdated);
            this.userGridView.KeyDown += new System.Windows.Forms.KeyEventHandler(this.userGridView_KeyDown);
            this.userGridView.ValidatingEditor += new DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventHandler(this.userGridView_ValidatingEditor);
            // 
            // colFirstName
            // 
            this.colFirstName.Caption = "First Name";
            this.colFirstName.FieldName = "FirstName";
            this.colFirstName.Name = "colFirstName";
            this.colFirstName.Visible = true;
            this.colFirstName.VisibleIndex = 0;
            // 
            // colLastName
            // 
            this.colLastName.Caption = "Last Name";
            this.colLastName.FieldName = "LastName";
            this.colLastName.Name = "colLastName";
            this.colLastName.Visible = true;
            this.colLastName.VisibleIndex = 1;
            // 
            // colLoginName
            // 
            this.colLoginName.Caption = "Login Name";
            this.colLoginName.FieldName = "LoginName";
            this.colLoginName.Name = "colLoginName";
            this.colLoginName.Visible = true;
            this.colLoginName.VisibleIndex = 2;
            // 
            // colIsAdmin
            // 
            this.colIsAdmin.Caption = "Admin";
            this.colIsAdmin.ColumnEdit = this.riCheckEdit;
            this.colIsAdmin.FieldName = "IsAdmin";
            this.colIsAdmin.Name = "colIsAdmin";
            this.colIsAdmin.Visible = true;
            this.colIsAdmin.VisibleIndex = 3;
            // 
            // riCheckEdit
            // 
            this.riCheckEdit.AutoHeight = false;
            this.riCheckEdit.Name = "riCheckEdit";
            // 
            // colCanChangeDates
            // 
            this.colCanChangeDates.Caption = "Change Dates";
            this.colCanChangeDates.ColumnEdit = this.riCheckEdit;
            this.colCanChangeDates.FieldName = "CanChangeDates";
            this.colCanChangeDates.Name = "colCanChangeDates";
            this.colCanChangeDates.Visible = true;
            this.colCanChangeDates.VisibleIndex = 4;
            // 
            // colEngineeringNumberVisible
            // 
            this.colEngineeringNumberVisible.Caption = "Engineering # Visible";
            this.colEngineeringNumberVisible.ColumnEdit = this.riCheckEdit;
            this.colEngineeringNumberVisible.FieldName = "EngineeringNumberVisible";
            this.colEngineeringNumberVisible.Name = "colEngineeringNumberVisible";
            this.colEngineeringNumberVisible.Visible = true;
            this.colEngineeringNumberVisible.VisibleIndex = 5;
            // 
            // colCanReadOnly
            // 
            this.colCanReadOnly.Caption = "Read Only";
            this.colCanReadOnly.ColumnEdit = this.riCheckEdit;
            this.colCanReadOnly.FieldName = "CanReadOnly";
            this.colCanReadOnly.Name = "colCanReadOnly";
            this.colCanReadOnly.Visible = true;
            this.colCanReadOnly.VisibleIndex = 6;
            // 
            // colCanOnlyChangeDesignWork
            // 
            this.colCanOnlyChangeDesignWork.Caption = "Change Design Work";
            this.colCanOnlyChangeDesignWork.ColumnEdit = this.riCheckEdit;
            this.colCanOnlyChangeDesignWork.FieldName = "CanOnlyChangeDesignWork";
            this.colCanOnlyChangeDesignWork.Name = "colCanOnlyChangeDesignWork";
            this.colCanOnlyChangeDesignWork.Visible = true;
            this.colCanOnlyChangeDesignWork.VisibleIndex = 7;
            // 
            // colCanChangeProjectData
            // 
            this.colCanChangeProjectData.Caption = "Change Project Data";
            this.colCanChangeProjectData.ColumnEdit = this.riCheckEdit;
            this.colCanChangeProjectData.FieldName = "CanChangeProjectData";
            this.colCanChangeProjectData.Name = "colCanChangeProjectData";
            this.colCanChangeProjectData.Visible = true;
            this.colCanChangeProjectData.VisibleIndex = 8;
            // 
            // AdminWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(908, 579);
            this.Controls.Add(this.gridControl1);
            this.Name = "AdminWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Admin";
            this.Load += new System.EventHandler(this.AdminWindow_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.userGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.riCheckEdit)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView userGridView;
        private DevExpress.XtraGrid.Columns.GridColumn colFirstName;
        private DevExpress.XtraGrid.Columns.GridColumn colLastName;
        private DevExpress.XtraGrid.Columns.GridColumn colLoginName;
        private DevExpress.XtraGrid.Columns.GridColumn colIsAdmin;
        private DevExpress.XtraGrid.Columns.GridColumn colCanChangeDates;
        private DevExpress.XtraGrid.Columns.GridColumn colEngineeringNumberVisible;
        private DevExpress.XtraGrid.Columns.GridColumn colCanReadOnly;
        private DevExpress.XtraGrid.Columns.GridColumn colCanOnlyChangeDesignWork;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit riCheckEdit;
        private DevExpress.XtraGrid.Columns.GridColumn colCanChangeProjectData;
    }
}