namespace Toolroom_Project_Viewer
{
    partial class ProjectSelectionForm
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
            this.components = new System.ComponentModel.Container();
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.workLoadBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.workload_Tracking_System_DBDataSet = new Toolroom_Project_Viewer.Workload_Tracking_System_DBDataSet();
            this.bandedGridView1 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridView();
            this.colToolNumber = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colMWONumber = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colProjectNumber = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colStage = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colCustomer = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colPartName = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colMoldCost = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colDeliveryInWeeks = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colStartDate = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colID = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colFinishDate = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colAdjustedDeliveryDate = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colEngineer = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colDesigner = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colToolMaker = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colRoughProgrammer = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colElectrodeProgrammer = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colFinishProgrammer = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colManifold = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colMoldBase = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colGeneralNotes = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colGeneralNotesRTF = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.UseSelectedProjectButton = new DevExpress.XtraEditors.SimpleButton();
            this.workLoadTableAdapter = new Toolroom_Project_Viewer.Workload_Tracking_System_DBDataSetTableAdapters.WorkLoadTableAdapter();
            this.colApprentice = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.gridBand1 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            this.gridBand2 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            this.gridBand3 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            this.gridBand4 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.workLoadBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.workload_Tracking_System_DBDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bandedGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // gridControl1
            // 
            this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridControl1.DataSource = this.workLoadBindingSource;
            this.gridControl1.Location = new System.Drawing.Point(12, 46);
            this.gridControl1.MainView = this.bandedGridView1;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.Size = new System.Drawing.Size(1119, 569);
            this.gridControl1.TabIndex = 0;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.bandedGridView1});
            // 
            // workLoadBindingSource
            // 
            this.workLoadBindingSource.DataMember = "WorkLoad";
            this.workLoadBindingSource.DataSource = this.workload_Tracking_System_DBDataSet;
            // 
            // workload_Tracking_System_DBDataSet
            // 
            this.workload_Tracking_System_DBDataSet.DataSetName = "Workload_Tracking_System_DBDataSet";
            this.workload_Tracking_System_DBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // bandedGridView1
            // 
            this.bandedGridView1.Bands.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] {
            this.gridBand1,
            this.gridBand2,
            this.gridBand3,
            this.gridBand4});
            this.bandedGridView1.Columns.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn[] {
            this.colID,
            this.colToolNumber,
            this.colMWONumber,
            this.colProjectNumber,
            this.colStage,
            this.colCustomer,
            this.colPartName,
            this.colEngineer,
            this.colDeliveryInWeeks,
            this.colStartDate,
            this.colFinishDate,
            this.colAdjustedDeliveryDate,
            this.colMoldCost,
            this.colDesigner,
            this.colToolMaker,
            this.colRoughProgrammer,
            this.colElectrodeProgrammer,
            this.colFinishProgrammer,
            this.colApprentice,
            this.colManifold,
            this.colMoldBase,
            this.colGeneralNotes,
            this.colGeneralNotesRTF});
            this.bandedGridView1.GridControl = this.gridControl1;
            this.bandedGridView1.GroupCount = 1;
            this.bandedGridView1.Name = "bandedGridView1";
            this.bandedGridView1.OptionsView.ColumnAutoWidth = false;
            this.bandedGridView1.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.colStage, DevExpress.Data.ColumnSortOrder.Ascending)});
            // 
            // colToolNumber
            // 
            this.colToolNumber.FieldName = "ToolNumber";
            this.colToolNumber.Name = "colToolNumber";
            this.colToolNumber.Visible = true;
            this.colToolNumber.Width = 80;
            // 
            // colMWONumber
            // 
            this.colMWONumber.FieldName = "MWONumber";
            this.colMWONumber.Name = "colMWONumber";
            this.colMWONumber.Visible = true;
            this.colMWONumber.Width = 81;
            // 
            // colProjectNumber
            // 
            this.colProjectNumber.FieldName = "ProjectNumber";
            this.colProjectNumber.Name = "colProjectNumber";
            this.colProjectNumber.Visible = true;
            this.colProjectNumber.Width = 59;
            // 
            // colStage
            // 
            this.colStage.FieldName = "Stage";
            this.colStage.Name = "colStage";
            this.colStage.Visible = true;
            this.colStage.Width = 53;
            // 
            // colCustomer
            // 
            this.colCustomer.FieldName = "Customer";
            this.colCustomer.Name = "colCustomer";
            this.colCustomer.Visible = true;
            this.colCustomer.Width = 73;
            // 
            // colPartName
            // 
            this.colPartName.FieldName = "PartName";
            this.colPartName.Name = "colPartName";
            this.colPartName.Visible = true;
            this.colPartName.Width = 149;
            // 
            // colMoldCost
            // 
            this.colMoldCost.FieldName = "MoldCost";
            this.colMoldCost.Name = "colMoldCost";
            this.colMoldCost.Visible = true;
            this.colMoldCost.Width = 65;
            // 
            // colDeliveryInWeeks
            // 
            this.colDeliveryInWeeks.FieldName = "DeliveryInWeeks";
            this.colDeliveryInWeeks.Name = "colDeliveryInWeeks";
            this.colDeliveryInWeeks.Visible = true;
            this.colDeliveryInWeeks.Width = 66;
            // 
            // colStartDate
            // 
            this.colStartDate.FieldName = "StartDate";
            this.colStartDate.Name = "colStartDate";
            this.colStartDate.Visible = true;
            this.colStartDate.Width = 70;
            // 
            // colID
            // 
            this.colID.FieldName = "ID";
            this.colID.Name = "colID";
            // 
            // colFinishDate
            // 
            this.colFinishDate.FieldName = "FinishDate";
            this.colFinishDate.Name = "colFinishDate";
            this.colFinishDate.Visible = true;
            this.colFinishDate.Width = 70;
            // 
            // colAdjustedDeliveryDate
            // 
            this.colAdjustedDeliveryDate.FieldName = "AdjustedDeliveryDate";
            this.colAdjustedDeliveryDate.Name = "colAdjustedDeliveryDate";
            this.colAdjustedDeliveryDate.Visible = true;
            this.colAdjustedDeliveryDate.Width = 63;
            // 
            // colEngineer
            // 
            this.colEngineer.FieldName = "Engineer";
            this.colEngineer.Name = "colEngineer";
            this.colEngineer.Visible = true;
            this.colEngineer.Width = 55;
            // 
            // colDesigner
            // 
            this.colDesigner.FieldName = "Designer";
            this.colDesigner.Name = "colDesigner";
            this.colDesigner.Visible = true;
            this.colDesigner.Width = 65;
            // 
            // colToolMaker
            // 
            this.colToolMaker.FieldName = "ToolMaker";
            this.colToolMaker.Name = "colToolMaker";
            this.colToolMaker.Visible = true;
            this.colToolMaker.Width = 65;
            // 
            // colRoughProgrammer
            // 
            this.colRoughProgrammer.FieldName = "RoughProgrammer";
            this.colRoughProgrammer.Name = "colRoughProgrammer";
            this.colRoughProgrammer.Visible = true;
            this.colRoughProgrammer.Width = 65;
            // 
            // colElectrodeProgrammer
            // 
            this.colElectrodeProgrammer.FieldName = "ElectrodeProgrammer";
            this.colElectrodeProgrammer.Name = "colElectrodeProgrammer";
            this.colElectrodeProgrammer.Visible = true;
            this.colElectrodeProgrammer.Width = 65;
            // 
            // colFinishProgrammer
            // 
            this.colFinishProgrammer.FieldName = "FinishProgrammer";
            this.colFinishProgrammer.Name = "colFinishProgrammer";
            this.colFinishProgrammer.Visible = true;
            this.colFinishProgrammer.Width = 65;
            // 
            // colManifold
            // 
            this.colManifold.FieldName = "Manifold";
            this.colManifold.Name = "colManifold";
            this.colManifold.Visible = true;
            this.colManifold.Width = 54;
            // 
            // colMoldBase
            // 
            this.colMoldBase.FieldName = "MoldBase";
            this.colMoldBase.Name = "colMoldBase";
            this.colMoldBase.Visible = true;
            this.colMoldBase.Width = 54;
            // 
            // colGeneralNotes
            // 
            this.colGeneralNotes.FieldName = "GeneralNotes";
            this.colGeneralNotes.Name = "colGeneralNotes";
            this.colGeneralNotes.Visible = true;
            this.colGeneralNotes.Width = 54;
            // 
            // colGeneralNotesRTF
            // 
            this.colGeneralNotesRTF.FieldName = "GeneralNotesRTF";
            this.colGeneralNotesRTF.Name = "colGeneralNotesRTF";
            this.colGeneralNotesRTF.Visible = true;
            this.colGeneralNotesRTF.Width = 83;
            // 
            // UseSelectedProjectButton
            // 
            this.UseSelectedProjectButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.UseSelectedProjectButton.Location = new System.Drawing.Point(12, 12);
            this.UseSelectedProjectButton.Name = "UseSelectedProjectButton";
            this.UseSelectedProjectButton.Size = new System.Drawing.Size(131, 25);
            this.UseSelectedProjectButton.TabIndex = 1;
            this.UseSelectedProjectButton.Text = "Use Selected Project";
            this.UseSelectedProjectButton.Click += new System.EventHandler(this.UseSelectedProjectButton_Click);
            // 
            // workLoadTableAdapter
            // 
            this.workLoadTableAdapter.ClearBeforeFill = true;
            // 
            // colApprentice
            // 
            this.colApprentice.Caption = "Apprentice";
            this.colApprentice.FieldName = "Apprentice";
            this.colApprentice.Name = "colApprentice";
            this.colApprentice.Visible = true;
            // 
            // gridBand1
            // 
            this.gridBand1.Caption = "Project";
            this.gridBand1.Columns.Add(this.colToolNumber);
            this.gridBand1.Columns.Add(this.colMWONumber);
            this.gridBand1.Columns.Add(this.colProjectNumber);
            this.gridBand1.Columns.Add(this.colStage);
            this.gridBand1.Columns.Add(this.colCustomer);
            this.gridBand1.Columns.Add(this.colPartName);
            this.gridBand1.Columns.Add(this.colMoldCost);
            this.gridBand1.Columns.Add(this.colDeliveryInWeeks);
            this.gridBand1.Name = "gridBand1";
            this.gridBand1.VisibleIndex = 0;
            this.gridBand1.Width = 626;
            // 
            // gridBand2
            // 
            this.gridBand2.Caption = "Milestones";
            this.gridBand2.Columns.Add(this.colStartDate);
            this.gridBand2.Columns.Add(this.colID);
            this.gridBand2.Columns.Add(this.colFinishDate);
            this.gridBand2.Columns.Add(this.colAdjustedDeliveryDate);
            this.gridBand2.Name = "gridBand2";
            this.gridBand2.VisibleIndex = 1;
            this.gridBand2.Width = 203;
            // 
            // gridBand3
            // 
            this.gridBand3.Caption = "Personnel";
            this.gridBand3.Columns.Add(this.colEngineer);
            this.gridBand3.Columns.Add(this.colDesigner);
            this.gridBand3.Columns.Add(this.colToolMaker);
            this.gridBand3.Columns.Add(this.colRoughProgrammer);
            this.gridBand3.Columns.Add(this.colElectrodeProgrammer);
            this.gridBand3.Columns.Add(this.colFinishProgrammer);
            this.gridBand3.Columns.Add(this.colApprentice);
            this.gridBand3.Name = "gridBand3";
            this.gridBand3.VisibleIndex = 2;
            this.gridBand3.Width = 455;
            // 
            // gridBand4
            // 
            this.gridBand4.Caption = "General Info";
            this.gridBand4.Columns.Add(this.colManifold);
            this.gridBand4.Columns.Add(this.colMoldBase);
            this.gridBand4.Columns.Add(this.colGeneralNotes);
            this.gridBand4.Columns.Add(this.colGeneralNotesRTF);
            this.gridBand4.Name = "gridBand4";
            this.gridBand4.VisibleIndex = 3;
            this.gridBand4.Width = 245;
            // 
            // ProjectSelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1143, 627);
            this.Controls.Add(this.UseSelectedProjectButton);
            this.Controls.Add(this.gridControl1);
            this.Name = "ProjectSelectionForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Project Selection Form";
            this.Load += new System.EventHandler(this.ProjectSelectionForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.workLoadBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.workload_Tracking_System_DBDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bandedGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraEditors.SimpleButton UseSelectedProjectButton;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridView bandedGridView1;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colID;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colToolNumber;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colMWONumber;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colProjectNumber;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colStage;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colCustomer;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colPartName;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colEngineer;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colDeliveryInWeeks;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colStartDate;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colFinishDate;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colAdjustedDeliveryDate;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colMoldCost;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colDesigner;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colToolMaker;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colRoughProgrammer;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colElectrodeProgrammer;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colFinishProgrammer;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colManifold;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colMoldBase;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colGeneralNotes;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colGeneralNotesRTF;
        private Workload_Tracking_System_DBDataSet workload_Tracking_System_DBDataSet;
        private System.Windows.Forms.BindingSource workLoadBindingSource;
        private Workload_Tracking_System_DBDataSetTableAdapters.WorkLoadTableAdapter workLoadTableAdapter;
        private DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand1;
        private DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand2;
        private DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand3;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colApprentice;
        private DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand4;
    }
}