namespace KanBanDataExtractor
{
    partial class MainWindow
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
            this.taskRecordView = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.loadButton = new DevExpress.XtraEditors.SimpleButton();
            this.exportButton = new DevExpress.XtraEditors.SimpleButton();
            this.colProjectNumber = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colJobNumber = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colComponent = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colMaterial = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colQuantity = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colTaskID = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colTaskName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDuration = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colStartDate = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colFinishDate = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colHours = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colNotes = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colInitials = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDate = new DevExpress.XtraGrid.Columns.GridColumn();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.taskRecordView)).BeginInit();
            this.SuspendLayout();
            // 
            // gridControl1
            // 
            this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridControl1.Location = new System.Drawing.Point(12, 58);
            this.gridControl1.MainView = this.taskRecordView;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.Size = new System.Drawing.Size(904, 507);
            this.gridControl1.TabIndex = 0;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.taskRecordView});
            // 
            // taskRecordView
            // 
            this.taskRecordView.Appearance.EvenRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.taskRecordView.Appearance.EvenRow.Options.UseBackColor = true;
            this.taskRecordView.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colProjectNumber,
            this.colJobNumber,
            this.colComponent,
            this.colMaterial,
            this.colQuantity,
            this.colTaskID,
            this.colTaskName,
            this.colDuration,
            this.colStartDate,
            this.colFinishDate,
            this.colHours,
            this.colNotes,
            this.colInitials,
            this.colDate});
            this.taskRecordView.GridControl = this.gridControl1;
            this.taskRecordView.Name = "taskRecordView";
            this.taskRecordView.OptionsView.ColumnAutoWidth = false;
            this.taskRecordView.OptionsView.EnableAppearanceEvenRow = true;
            // 
            // loadButton
            // 
            this.loadButton.Location = new System.Drawing.Point(12, 12);
            this.loadButton.Name = "loadButton";
            this.loadButton.Size = new System.Drawing.Size(70, 33);
            this.loadButton.TabIndex = 1;
            this.loadButton.Text = "Load";
            this.loadButton.Click += new System.EventHandler(this.loadButton_Click);
            // 
            // exportButton
            // 
            this.exportButton.Location = new System.Drawing.Point(88, 12);
            this.exportButton.Name = "exportButton";
            this.exportButton.Size = new System.Drawing.Size(70, 33);
            this.exportButton.TabIndex = 2;
            this.exportButton.Text = "Export";
            this.exportButton.Click += new System.EventHandler(this.exportButton_Click);
            // 
            // colProjectNumber
            // 
            this.colProjectNumber.AppearanceCell.Options.UseTextOptions = true;
            this.colProjectNumber.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colProjectNumber.AppearanceHeader.Options.UseTextOptions = true;
            this.colProjectNumber.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colProjectNumber.Caption = "Project #";
            this.colProjectNumber.FieldName = "ProjectNumber";
            this.colProjectNumber.Name = "colProjectNumber";
            this.colProjectNumber.Visible = true;
            this.colProjectNumber.VisibleIndex = 0;
            // 
            // colJobNumber
            // 
            this.colJobNumber.AppearanceCell.Options.UseTextOptions = true;
            this.colJobNumber.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colJobNumber.AppearanceHeader.Options.UseTextOptions = true;
            this.colJobNumber.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colJobNumber.Caption = "Job #";
            this.colJobNumber.FieldName = "JobNumber";
            this.colJobNumber.Name = "colJobNumber";
            this.colJobNumber.Visible = true;
            this.colJobNumber.VisibleIndex = 1;
            // 
            // colComponent
            // 
            this.colComponent.AppearanceHeader.Options.UseTextOptions = true;
            this.colComponent.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colComponent.Caption = "Component";
            this.colComponent.FieldName = "Component";
            this.colComponent.Name = "colComponent";
            this.colComponent.Visible = true;
            this.colComponent.VisibleIndex = 2;
            // 
            // colMaterial
            // 
            this.colMaterial.AppearanceCell.Options.UseTextOptions = true;
            this.colMaterial.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colMaterial.AppearanceHeader.Options.UseTextOptions = true;
            this.colMaterial.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colMaterial.Caption = "Material";
            this.colMaterial.FieldName = "Material";
            this.colMaterial.Name = "colMaterial";
            this.colMaterial.Visible = true;
            this.colMaterial.VisibleIndex = 3;
            // 
            // colQuantity
            // 
            this.colQuantity.AppearanceCell.Options.UseTextOptions = true;
            this.colQuantity.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colQuantity.AppearanceHeader.Options.UseTextOptions = true;
            this.colQuantity.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colQuantity.Caption = "Quantity";
            this.colQuantity.FieldName = "Quantity";
            this.colQuantity.Name = "colQuantity";
            this.colQuantity.Visible = true;
            this.colQuantity.VisibleIndex = 4;
            // 
            // colTaskID
            // 
            this.colTaskID.AppearanceCell.Options.UseTextOptions = true;
            this.colTaskID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colTaskID.AppearanceHeader.Options.UseTextOptions = true;
            this.colTaskID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colTaskID.Caption = "Task ID";
            this.colTaskID.FieldName = "TaskID";
            this.colTaskID.Name = "colTaskID";
            this.colTaskID.Visible = true;
            this.colTaskID.VisibleIndex = 5;
            // 
            // colTaskName
            // 
            this.colTaskName.AppearanceHeader.Options.UseTextOptions = true;
            this.colTaskName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colTaskName.Caption = "Task Name";
            this.colTaskName.FieldName = "TaskName";
            this.colTaskName.Name = "colTaskName";
            this.colTaskName.Visible = true;
            this.colTaskName.VisibleIndex = 6;
            // 
            // colDuration
            // 
            this.colDuration.AppearanceCell.Options.UseTextOptions = true;
            this.colDuration.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colDuration.AppearanceHeader.Options.UseTextOptions = true;
            this.colDuration.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colDuration.Caption = "Duration";
            this.colDuration.FieldName = "Duration";
            this.colDuration.Name = "colDuration";
            this.colDuration.Visible = true;
            this.colDuration.VisibleIndex = 7;
            // 
            // colStartDate
            // 
            this.colStartDate.AppearanceCell.Options.UseTextOptions = true;
            this.colStartDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colStartDate.AppearanceHeader.Options.UseTextOptions = true;
            this.colStartDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colStartDate.Caption = "Start Date";
            this.colStartDate.FieldName = "StartDate";
            this.colStartDate.Name = "colStartDate";
            this.colStartDate.Visible = true;
            this.colStartDate.VisibleIndex = 8;
            // 
            // colFinishDate
            // 
            this.colFinishDate.AppearanceCell.Options.UseTextOptions = true;
            this.colFinishDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colFinishDate.AppearanceHeader.Options.UseTextOptions = true;
            this.colFinishDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colFinishDate.Caption = "Finish Date";
            this.colFinishDate.FieldName = "FinishDate";
            this.colFinishDate.Name = "colFinishDate";
            this.colFinishDate.Visible = true;
            this.colFinishDate.VisibleIndex = 9;
            // 
            // colHours
            // 
            this.colHours.AppearanceCell.Options.UseTextOptions = true;
            this.colHours.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colHours.AppearanceHeader.Options.UseTextOptions = true;
            this.colHours.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colHours.Caption = "Hours";
            this.colHours.FieldName = "Hours";
            this.colHours.Name = "colHours";
            this.colHours.Visible = true;
            this.colHours.VisibleIndex = 10;
            // 
            // colNotes
            // 
            this.colNotes.AppearanceHeader.Options.UseTextOptions = true;
            this.colNotes.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colNotes.Caption = "Notes";
            this.colNotes.FieldName = "Notes";
            this.colNotes.Name = "colNotes";
            this.colNotes.Visible = true;
            this.colNotes.VisibleIndex = 11;
            // 
            // colInitials
            // 
            this.colInitials.AppearanceCell.Options.UseTextOptions = true;
            this.colInitials.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colInitials.AppearanceHeader.Options.UseTextOptions = true;
            this.colInitials.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colInitials.Caption = "Initials";
            this.colInitials.FieldName = "Initials";
            this.colInitials.Name = "colInitials";
            this.colInitials.Visible = true;
            this.colInitials.VisibleIndex = 12;
            // 
            // colDate
            // 
            this.colDate.AppearanceCell.Options.UseTextOptions = true;
            this.colDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colDate.AppearanceHeader.Options.UseTextOptions = true;
            this.colDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colDate.Caption = "Date";
            this.colDate.FieldName = "DateCompleted";
            this.colDate.Name = "colDate";
            this.colDate.Visible = true;
            this.colDate.VisibleIndex = 13;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(928, 577);
            this.Controls.Add(this.exportButton);
            this.Controls.Add(this.loadButton);
            this.Controls.Add(this.gridControl1);
            this.Name = "MainWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Kan Ban Data Extractor";
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.taskRecordView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView taskRecordView;
        private DevExpress.XtraEditors.SimpleButton loadButton;
        private DevExpress.XtraEditors.SimpleButton exportButton;
        private DevExpress.XtraGrid.Columns.GridColumn colProjectNumber;
        private DevExpress.XtraGrid.Columns.GridColumn colJobNumber;
        private DevExpress.XtraGrid.Columns.GridColumn colComponent;
        private DevExpress.XtraGrid.Columns.GridColumn colMaterial;
        private DevExpress.XtraGrid.Columns.GridColumn colQuantity;
        private DevExpress.XtraGrid.Columns.GridColumn colTaskID;
        private DevExpress.XtraGrid.Columns.GridColumn colTaskName;
        private DevExpress.XtraGrid.Columns.GridColumn colDuration;
        private DevExpress.XtraGrid.Columns.GridColumn colStartDate;
        private DevExpress.XtraGrid.Columns.GridColumn colFinishDate;
        private DevExpress.XtraGrid.Columns.GridColumn colHours;
        private DevExpress.XtraGrid.Columns.GridColumn colNotes;
        private DevExpress.XtraGrid.Columns.GridColumn colInitials;
        private DevExpress.XtraGrid.Columns.GridColumn colDate;
    }
}

