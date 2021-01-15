namespace Toolroom_Project_Viewer
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
            this.components = new System.ComponentModel.Container();
            DevExpress.XtraEditors.RangeControlRange rangeControlRange1 = new DevExpress.XtraEditors.RangeControlRange();
            DevExpress.XtraCharts.XYDiagram xyDiagram1 = new DevExpress.XtraCharts.XYDiagram();
            DevExpress.XtraCharts.Series series1 = new DevExpress.XtraCharts.Series();
            DevExpress.XtraCharts.SideBySideBarSeriesLabel sideBySideBarSeriesLabel1 = new DevExpress.XtraCharts.SideBySideBarSeriesLabel();
            DevExpress.XtraSplashScreen.SplashScreenManager splashScreenManager1 = new DevExpress.XtraSplashScreen.SplashScreenManager(this, typeof(global::Toolroom_Project_Viewer.MainSplashScreen), true, true);
            DevExpress.XtraEditors.RangeControlRange rangeControlRange2 = new DevExpress.XtraEditors.RangeControlRange();
            DevExpress.XtraGrid.GridFormatRule gridFormatRule1 = new DevExpress.XtraGrid.GridFormatRule();
            DevExpress.XtraEditors.FormatConditionRuleDataBar formatConditionRuleDataBar1 = new DevExpress.XtraEditors.FormatConditionRuleDataBar();
            DevExpress.XtraGrid.GridLevelNode gridLevelNode1 = new DevExpress.XtraGrid.GridLevelNode();
            DevExpress.XtraGrid.GridLevelNode gridLevelNode2 = new DevExpress.XtraGrid.GridLevelNode();
            DevExpress.XtraGrid.GridFormatRule gridFormatRule2 = new DevExpress.XtraGrid.GridFormatRule();
            DevExpress.XtraEditors.FormatConditionRuleDataBar formatConditionRuleDataBar2 = new DevExpress.XtraEditors.FormatConditionRuleDataBar();
            DevExpress.XtraGrid.GridFormatRule gridFormatRule3 = new DevExpress.XtraGrid.GridFormatRule();
            DevExpress.XtraEditors.FormatConditionRuleExpression formatConditionRuleExpression1 = new DevExpress.XtraEditors.FormatConditionRuleExpression();
            DevExpress.XtraGrid.GridFormatRule gridFormatRule4 = new DevExpress.XtraGrid.GridFormatRule();
            DevExpress.XtraEditors.FormatConditionRuleDataBar formatConditionRuleDataBar3 = new DevExpress.XtraEditors.FormatConditionRuleDataBar();
            DevExpress.XtraScheduler.TimeRuler timeRuler1 = new DevExpress.XtraScheduler.TimeRuler();
            DevExpress.XtraScheduler.TimeRuler timeRuler2 = new DevExpress.XtraScheduler.TimeRuler();
            DevExpress.XtraScheduler.TimeRuler timeRuler3 = new DevExpress.XtraScheduler.TimeRuler();
            DevExpress.XtraScheduler.TimeRuler timeRuler4 = new DevExpress.XtraScheduler.TimeRuler();
            DevExpress.XtraScheduler.TimeRuler timeRuler5 = new DevExpress.XtraScheduler.TimeRuler();
            DevExpress.XtraScheduler.TimeRuler timeRuler6 = new DevExpress.XtraScheduler.TimeRuler();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
            this.rangeControl1 = new DevExpress.XtraEditors.RangeControl();
            this.chartControl1 = new DevExpress.XtraCharts.ChartControl();
            this.workload_Tracking_System_DBDataSet = new Toolroom_Project_Viewer.Workload_Tracking_System_DBDataSet();
            this.rangeControl2 = new DevExpress.XtraEditors.RangeControl();
            this.colPercentComplete = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridView4 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colID2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colProjectNumber3 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colComponent1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colPictures = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colMaterial = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemComboBox3 = new DevExpress.XtraEditors.Repository.RepositoryItemComboBox();
            this.colFinish = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colNotes = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colPosition = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colPriority1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemSpinEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
            this.colQuantity = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colSpares = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colStatus2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridControl3 = new DevExpress.XtraGrid.GridControl();
            this.projectsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gridView3 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colID1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colJobNumber1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colProjectNumber2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colStagePV = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colCustomer1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colProject = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDueDate1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colPriority = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colStatus1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDesigner1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.personnelComboBoxEdit = new DevExpress.XtraEditors.Repository.RepositoryItemComboBox();
            this.colToolMaker2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colRoughProgrammer1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colElectrodeProgrammer1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colFinishProgrammer1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colApprentice = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colEngineer1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colOverlapAllowed = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colIncludeHours = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colKanBanWorkbookPath = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemHyperLinkEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemHyperLinkEdit();
            this.colPercentComplete1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDateModified = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colLastKanBanGenerationDate = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colLatestFinishDate = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemImageEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemImageEdit();
            this.repositoryItemPictureEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemPictureEdit();
            this.repositoryItemImageComboBox1 = new DevExpress.XtraEditors.Repository.RepositoryItemImageComboBox();
            this.repositoryItemTextEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
            this.stageComboBoxEdit = new DevExpress.XtraEditors.Repository.RepositoryItemComboBox();
            this.genericDateEdit = new DevExpress.XtraEditors.Repository.RepositoryItemDateEdit();
            this.projectBandedGridView = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridView();
            this.SegoeUI = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            this.colJobNumberBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colProjectNumberBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colStageBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colCustomerBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colProjectBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colMoldCostBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colDeliveryInWeeksBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.milestonesGridBand = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            this.colStatusBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colStartDateBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colDueDateBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colAdjustedDeliveryDateBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.personnelGridBand = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            this.colEngineerBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colDesignerBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colToolMakerBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colRoughProgrammerBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colElectrodeProgrammerBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colFinishProgrammerBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colApprenticeBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.generalInfoGridBand = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            this.colManifoldBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colMoldBaseBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colGeneralNotesBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.colIDBGV = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            this.gridView5 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colID4 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colTaskName1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colResource1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colMachine = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colHours = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDuration1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colStartDate2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colFinishDate2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colTaskID1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colPredecessors1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colNotes1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colStatus3 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colInitials = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDateCompleted = new DevExpress.XtraGrid.Columns.GridColumn();
            this.DeptProgressGridView = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.DepartmentColDPV = new DevExpress.XtraGrid.Columns.GridColumn();
            this.PercentCompleteColDPV = new DevExpress.XtraGrid.Columns.GridColumn();
            this.projectsTableAdapter = new Toolroom_Project_Viewer.Workload_Tracking_System_DBDataSetTableAdapters.ProjectsTableAdapter();
            this.tasksTableAdapter = new Toolroom_Project_Viewer.Workload_Tracking_System_DBDataSetTableAdapters.TasksTableAdapter();
            this.schedulerStorage1 = new DevExpress.XtraScheduler.SchedulerStorage(this.components);
            this.xtraTabControl1 = new DevExpress.XtraTab.XtraTabControl();
            this.xtraTabPage1 = new DevExpress.XtraTab.XtraTabPage();
            this.includeCompletesCheckEdit = new DevExpress.XtraEditors.CheckEdit();
            this.includeQuotesCheckEdit = new DevExpress.XtraEditors.CheckEdit();
            this.labelControl7 = new DevExpress.XtraEditors.LabelControl();
            this.projectCheckedComboBoxEdit = new DevExpress.XtraEditors.CheckedComboBoxEdit();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.GroupByRadioGroup = new DevExpress.XtraEditors.RadioGroup();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.refreshButton = new DevExpress.XtraEditors.SimpleButton();
            this.schedulerControl1 = new DevExpress.XtraScheduler.SchedulerControl();
            this.departmentComboBox = new DevExpress.XtraEditors.ComboBoxEdit();
            this.xtraTabPage2 = new DevExpress.XtraTab.XtraTabPage();
            this.PrintEmployeeWorkCheckedComboBoxEdit = new DevExpress.XtraEditors.CheckedComboBoxEdit();
            this.labelControl8 = new DevExpress.XtraEditors.LabelControl();
            this.daysAheadSpinEdit = new DevExpress.XtraEditors.SpinEdit();
            this.filterTasksByDatesCheckEdit = new DevExpress.XtraEditors.CheckEdit();
            this.printEmployeeWorkButton = new DevExpress.XtraEditors.SimpleButton();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.PrintDeptsCheckedComboBoxEdit = new DevExpress.XtraEditors.CheckedComboBoxEdit();
            this.printTaskViewButton = new DevExpress.XtraEditors.SimpleButton();
            this.RefreshTasksButton = new DevExpress.XtraEditors.SimpleButton();
            this.departmentComboBox2 = new DevExpress.XtraEditors.ComboBoxEdit();
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colID3 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colProjectStatus = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colJobNumber = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colProjectNumber = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colComponent = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemHyperLinkEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemHyperLinkEdit();
            this.colTaskID = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colTaskName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colNotes2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colToolMaker = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colHours1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemSpinEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
            this.colDuration = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colStartDate = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colFinishDate = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colPredecessors = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDueDate = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colMachine1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemCheckedComboBoxEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckedComboBoxEdit();
            this.colPersonnel = new DevExpress.XtraGrid.Columns.GridColumn();
            this.resourceRepositoryItemComboBox = new DevExpress.XtraEditors.Repository.RepositoryItemComboBox();
            this.colStatus = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemDateEdit4 = new DevExpress.XtraEditors.Repository.RepositoryItemDateEdit();
            this.repositoryItemDateEdit5 = new DevExpress.XtraEditors.Repository.RepositoryItemDateEdit();
            this.xtraTabPage7 = new DevExpress.XtraTab.XtraTabPage();
            this.restoreProjectButton = new DevExpress.XtraEditors.SimpleButton();
            this.workLoadViewPrintPreviewButton = new DevExpress.XtraEditors.SimpleButton();
            this.workLoadViewPrint2Button = new DevExpress.XtraEditors.SimpleButton();
            this.workLoadViewPrintButton = new DevExpress.XtraEditors.SimpleButton();
            this.changeViewRadioGroup = new DevExpress.XtraEditors.RadioGroup();
            this.refreshLabelControl = new DevExpress.XtraEditors.LabelControl();
            this.resourceButton = new DevExpress.XtraEditors.SimpleButton();
            this.editProjectButton = new DevExpress.XtraEditors.SimpleButton();
            this.createProjectButton = new DevExpress.XtraEditors.SimpleButton();
            this.backDateButton = new DevExpress.XtraEditors.SimpleButton();
            this.forwardDateButton = new DevExpress.XtraEditors.SimpleButton();
            this.kanBanButton = new DevExpress.XtraEditors.SimpleButton();
            this.copyButton = new DevExpress.XtraEditors.SimpleButton();
            this.RefreshProjectsButton = new DevExpress.XtraEditors.SimpleButton();
            this.xtraTabPage3 = new DevExpress.XtraTab.XtraTabPage();
            this.chartRadioGroup = new DevExpress.XtraEditors.RadioGroup();
            this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
            this.timeFrameComboBoxEdit = new DevExpress.XtraEditors.ComboBoxEdit();
            this.TimeUnitsComboBox = new DevExpress.XtraEditors.ComboBoxEdit();
            this.RefreshChartButton = new DevExpress.XtraEditors.SimpleButton();
            this.xtraTabPage4 = new DevExpress.XtraTab.XtraTabPage();
            this.labelControl6 = new DevExpress.XtraEditors.LabelControl();
            this.panel1 = new System.Windows.Forms.Panel();
            this.splitContainerControl1 = new DevExpress.XtraEditors.SplitContainerControl();
            this.resourcesTree1 = new DevExpress.XtraScheduler.UI.ResourcesTree();
            this.colCaption = new DevExpress.XtraScheduler.Native.ResourceTreeColumn();
            this.schedulerControl2 = new DevExpress.XtraScheduler.SchedulerControl();
            this.schedulerStorage2 = new DevExpress.XtraScheduler.SchedulerStorage(this.components);
            this.projectComboBox = new DevExpress.XtraEditors.ComboBoxEdit();
            this.RefreshGanttButton = new DevExpress.XtraEditors.SimpleButton();
            this.tasksBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.defaultLookAndFeel1 = new DevExpress.LookAndFeel.DefaultLookAndFeel(this.components);
            this.resourcesBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.resourcesTableAdapter = new Toolroom_Project_Viewer.Workload_Tracking_System_DBDataSetTableAdapters.ResourcesTableAdapter();
            this.behaviorManager1 = new DevExpress.Utils.Behaviors.BehaviorManager(this.components);
            this.componentsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.componentsTableAdapter = new Toolroom_Project_Viewer.Workload_Tracking_System_DBDataSetTableAdapters.ComponentsTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.rangeControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chartControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(xyDiagram1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(series1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(sideBySideBarSeriesLabel1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.workload_Tracking_System_DBDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rangeControl2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemComboBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.projectsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.personnelComboBoxEdit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemHyperLinkEdit2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemImageEdit2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemPictureEdit1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemImageComboBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.stageComboBoxEdit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.genericDateEdit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.genericDateEdit.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.projectBandedGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DeptProgressGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerStorage1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).BeginInit();
            this.xtraTabControl1.SuspendLayout();
            this.xtraTabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.includeCompletesCheckEdit.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.includeQuotesCheckEdit.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.projectCheckedComboBoxEdit.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GroupByRadioGroup.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.departmentComboBox.Properties)).BeginInit();
            this.xtraTabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PrintEmployeeWorkCheckedComboBoxEdit.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.daysAheadSpinEdit.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.filterTasksByDatesCheckEdit.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PrintDeptsCheckedComboBoxEdit.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.departmentComboBox2.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemHyperLinkEdit1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckedComboBoxEdit1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.resourceRepositoryItemComboBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit4.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit5.CalendarTimeProperties)).BeginInit();
            this.xtraTabPage7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.changeViewRadioGroup.Properties)).BeginInit();
            this.xtraTabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartRadioGroup.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.timeFrameComboBoxEdit.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TimeUnitsComboBox.Properties)).BeginInit();
            this.xtraTabPage4.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).BeginInit();
            this.splitContainerControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.resourcesTree1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerControl2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerStorage2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.projectComboBox.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tasksBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.resourcesBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.behaviorManager1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.componentsBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // rangeControl1
            // 
            this.rangeControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rangeControl1.Client = this.chartControl1;
            this.rangeControl1.Location = new System.Drawing.Point(12, 849);
            this.rangeControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.rangeControl1.Name = "rangeControl1";
            rangeControlRange1.Maximum = 9.6D;
            rangeControlRange1.Minimum = -0.6D;
            rangeControlRange1.Owner = this.rangeControl1;
            this.rangeControl1.SelectedRange = rangeControlRange1;
            this.rangeControl1.Size = new System.Drawing.Size(1492, 52);
            this.rangeControl1.TabIndex = 16;
            this.rangeControl1.Text = "rangeControl1";
            this.rangeControl1.VisibleRangeMaximumScaleFactor = double.PositiveInfinity;
            // 
            // chartControl1
            // 
            this.chartControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.chartControl1.CrosshairOptions.CrosshairLabelMode = DevExpress.XtraCharts.CrosshairLabelMode.ShowForNearestSeries;
            this.chartControl1.CrosshairOptions.GroupHeaderPattern = "{A}";
            this.chartControl1.CrosshairOptions.HighlightPoints = false;
            this.chartControl1.DataSource = this.workload_Tracking_System_DBDataSet.Machines;
            xyDiagram1.AxisX.Tickmarks.MinorVisible = false;
            xyDiagram1.AxisX.VisibleInPanesSerializable = "-1";
            xyDiagram1.AxisY.VisibleInPanesSerializable = "-1";
            this.chartControl1.Diagram = xyDiagram1;
            this.chartControl1.Legend.AlignmentHorizontal = DevExpress.XtraCharts.LegendAlignmentHorizontal.Right;
            this.chartControl1.Legend.Name = "Default Legend";
            this.chartControl1.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True;
            this.chartControl1.Location = new System.Drawing.Point(12, 69);
            this.chartControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chartControl1.Name = "chartControl1";
            series1.Name = "Hours";
            this.chartControl1.SeriesSerializable = new DevExpress.XtraCharts.Series[] {
        series1};
            this.chartControl1.SeriesTemplate.CrosshairLabelPattern = "{A}:{V}";
            sideBySideBarSeriesLabel1.ResolveOverlappingMode = DevExpress.XtraCharts.ResolveOverlappingMode.Default;
            sideBySideBarSeriesLabel1.TextPattern = "{A}-{V}";
            this.chartControl1.SeriesTemplate.Label = sideBySideBarSeriesLabel1;
            this.chartControl1.SeriesTemplate.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
            this.chartControl1.SeriesTemplate.SeriesColorizer = null;
            this.chartControl1.SeriesTemplate.ToolTipSeriesPattern = "{S}-";
            this.chartControl1.Size = new System.Drawing.Size(1572, 665);
            this.chartControl1.TabIndex = 0;
            // 
            // workload_Tracking_System_DBDataSet
            // 
            this.workload_Tracking_System_DBDataSet.DataSetName = "Workload_Tracking_System_DBDataSet";
            this.workload_Tracking_System_DBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // splashScreenManager1
            // 
            splashScreenManager1.ClosingDelay = 250;
            // 
            // rangeControl2
            // 
            this.rangeControl2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rangeControl2.Client = this.chartControl1;
            this.rangeControl2.Location = new System.Drawing.Point(12, 742);
            this.rangeControl2.Name = "rangeControl2";
            rangeControlRange2.Maximum = 9.6D;
            rangeControlRange2.Minimum = -0.6D;
            rangeControlRange2.Owner = this.rangeControl2;
            this.rangeControl2.SelectedRange = rangeControlRange2;
            this.rangeControl2.Size = new System.Drawing.Size(1572, 46);
            this.rangeControl2.TabIndex = 19;
            this.rangeControl2.Text = "rangeControl2";
            // 
            // colPercentComplete
            // 
            this.colPercentComplete.DisplayFormat.FormatString = "P0";
            this.colPercentComplete.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.colPercentComplete.FieldName = "PercentComplete";
            this.colPercentComplete.Name = "colPercentComplete";
            this.colPercentComplete.OptionsColumn.AllowEdit = false;
            this.colPercentComplete.Visible = true;
            this.colPercentComplete.VisibleIndex = 9;
            this.colPercentComplete.Width = 116;
            // 
            // gridView4
            // 
            this.gridView4.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.gridView4.Appearance.SelectedRow.Options.UseBackColor = true;
            this.gridView4.AppearancePrint.HeaderPanel.Options.UseTextOptions = true;
            this.gridView4.AppearancePrint.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.gridView4.ChildGridLevelName = "Tasks";
            this.gridView4.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID2,
            this.colProjectNumber3,
            this.colComponent1,
            this.colPictures,
            this.colMaterial,
            this.colFinish,
            this.colNotes,
            this.colPosition,
            this.colPriority1,
            this.colQuantity,
            this.colSpares,
            this.colStatus2,
            this.colPercentComplete});
            gridFormatRule1.Column = this.colPercentComplete;
            gridFormatRule1.Name = "Format0";
            formatConditionRuleDataBar1.Maximum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            formatConditionRuleDataBar1.MaximumType = DevExpress.XtraEditors.FormatConditionValueType.Number;
            formatConditionRuleDataBar1.MinimumType = DevExpress.XtraEditors.FormatConditionValueType.Number;
            formatConditionRuleDataBar1.PredefinedName = "Blue";
            gridFormatRule1.Rule = formatConditionRuleDataBar1;
            this.gridView4.FormatRules.Add(gridFormatRule1);
            this.gridView4.GridControl = this.gridControl3;
            this.gridView4.Name = "gridView4";
            this.gridView4.OptionsBehavior.EditingMode = DevExpress.XtraGrid.Views.Grid.GridEditingMode.Inplace;
            this.gridView4.OptionsPrint.AllowMultilineHeaders = true;
            this.gridView4.OptionsPrint.AutoWidth = false;
            this.gridView4.OptionsSelection.MultiSelect = true;
            this.gridView4.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
            this.gridView4.OptionsView.ColumnAutoWidth = false;
            this.gridView4.MasterRowExpanded += new DevExpress.XtraGrid.Views.Grid.CustomMasterRowEventHandler(this.gridView_MasterRowExpanded);
            this.gridView4.MasterRowCollapsed += new DevExpress.XtraGrid.Views.Grid.CustomMasterRowEventHandler(this.gridView_MasterRowCollapsed);
            this.gridView4.CustomRowCellEditForEditing += new DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventHandler(this.gridView4_CustomRowCellEditForEditing);
            this.gridView4.CellValueChanged += new DevExpress.XtraGrid.Views.Base.CellValueChangedEventHandler(this.gridView4_CellValueChanged);
            this.gridView4.ValidatingEditor += new DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventHandler(this.GridView4_ValidatingEditor);
            // 
            // colID2
            // 
            this.colID2.FieldName = "ID";
            this.colID2.Name = "colID2";
            // 
            // colProjectNumber3
            // 
            this.colProjectNumber3.FieldName = "ProjectNumber";
            this.colProjectNumber3.Name = "colProjectNumber3";
            // 
            // colComponent1
            // 
            this.colComponent1.FieldName = "Component";
            this.colComponent1.Name = "colComponent1";
            this.colComponent1.Visible = true;
            this.colComponent1.VisibleIndex = 0;
            this.colComponent1.Width = 157;
            // 
            // colPictures
            // 
            this.colPictures.FieldName = "Picture";
            this.colPictures.Name = "colPictures";
            this.colPictures.Visible = true;
            this.colPictures.VisibleIndex = 1;
            this.colPictures.Width = 91;
            // 
            // colMaterial
            // 
            this.colMaterial.ColumnEdit = this.repositoryItemComboBox3;
            this.colMaterial.FieldName = "Material";
            this.colMaterial.Name = "colMaterial";
            this.colMaterial.Visible = true;
            this.colMaterial.VisibleIndex = 2;
            this.colMaterial.Width = 143;
            // 
            // repositoryItemComboBox3
            // 
            this.repositoryItemComboBox3.AutoHeight = false;
            this.repositoryItemComboBox3.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemComboBox3.Items.AddRange(new object[] {
            "420 SS",
            "Aluminum",
            "Caldie",
            "Copper",
            "H13",
            "HRC60",
            "Moldmax",
            "Moldstar 90",
            "Moldstar 97",
            "P20",
            "S7",
            "W360"});
            this.repositoryItemComboBox3.Name = "repositoryItemComboBox3";
            // 
            // colFinish
            // 
            this.colFinish.FieldName = "Finish";
            this.colFinish.Name = "colFinish";
            this.colFinish.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colFinish.Visible = true;
            this.colFinish.VisibleIndex = 3;
            this.colFinish.Width = 143;
            // 
            // colNotes
            // 
            this.colNotes.FieldName = "Notes";
            this.colNotes.Name = "colNotes";
            this.colNotes.Visible = true;
            this.colNotes.VisibleIndex = 4;
            this.colNotes.Width = 313;
            // 
            // colPosition
            // 
            this.colPosition.FieldName = "Position";
            this.colPosition.Name = "colPosition";
            this.colPosition.Width = 52;
            // 
            // colPriority1
            // 
            this.colPriority1.AppearanceCell.Options.UseTextOptions = true;
            this.colPriority1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colPriority1.ColumnEdit = this.repositoryItemSpinEdit2;
            this.colPriority1.FieldName = "Priority";
            this.colPriority1.Name = "colPriority1";
            this.colPriority1.Visible = true;
            this.colPriority1.VisibleIndex = 5;
            this.colPriority1.Width = 56;
            // 
            // repositoryItemSpinEdit2
            // 
            this.repositoryItemSpinEdit2.AutoHeight = false;
            this.repositoryItemSpinEdit2.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemSpinEdit2.EditFormat.FormatString = "d0";
            this.repositoryItemSpinEdit2.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.repositoryItemSpinEdit2.Name = "repositoryItemSpinEdit2";
            // 
            // colQuantity
            // 
            this.colQuantity.AppearanceCell.Options.UseTextOptions = true;
            this.colQuantity.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colQuantity.ColumnEdit = this.repositoryItemSpinEdit2;
            this.colQuantity.FieldName = "Quantity";
            this.colQuantity.Name = "colQuantity";
            this.colQuantity.Visible = true;
            this.colQuantity.VisibleIndex = 6;
            this.colQuantity.Width = 65;
            // 
            // colSpares
            // 
            this.colSpares.AppearanceCell.Options.UseTextOptions = true;
            this.colSpares.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colSpares.ColumnEdit = this.repositoryItemSpinEdit2;
            this.colSpares.FieldName = "Spares";
            this.colSpares.Name = "colSpares";
            this.colSpares.Visible = true;
            this.colSpares.VisibleIndex = 7;
            this.colSpares.Width = 56;
            // 
            // colStatus2
            // 
            this.colStatus2.FieldName = "Status";
            this.colStatus2.Name = "colStatus2";
            this.colStatus2.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colStatus2.Visible = true;
            this.colStatus2.VisibleIndex = 8;
            this.colStatus2.Width = 151;
            // 
            // gridControl3
            // 
            this.gridControl3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridControl3.DataSource = this.projectsBindingSource;
            this.gridControl3.EmbeddedNavigator.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.gridControl3.Font = new System.Drawing.Font("Tahoma", 8.25F);
            gridLevelNode1.LevelTemplate = this.gridView4;
            gridLevelNode2.LevelTemplate = this.gridView5;
            gridLevelNode2.RelationName = "Tasks";
            gridLevelNode1.Nodes.AddRange(new DevExpress.XtraGrid.GridLevelNode[] {
            gridLevelNode2});
            gridLevelNode1.RelationName = "Components";
            this.gridControl3.LevelTree.Nodes.AddRange(new DevExpress.XtraGrid.GridLevelNode[] {
            gridLevelNode1});
            this.gridControl3.Location = new System.Drawing.Point(13, 41);
            this.gridControl3.MainView = this.projectBandedGridView;
            this.gridControl3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.gridControl3.Name = "gridControl3";
            this.gridControl3.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemSpinEdit2,
            this.repositoryItemComboBox3,
            this.repositoryItemImageEdit2,
            this.repositoryItemPictureEdit1,
            this.repositoryItemImageComboBox1,
            this.repositoryItemHyperLinkEdit2,
            this.repositoryItemTextEdit2,
            this.stageComboBoxEdit,
            this.genericDateEdit,
            this.personnelComboBoxEdit});
            this.gridControl3.Size = new System.Drawing.Size(1573, 746);
            this.gridControl3.TabIndex = 0;
            this.gridControl3.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView3,
            this.projectBandedGridView,
            this.gridView5,
            this.DeptProgressGridView,
            this.gridView4});
            this.gridControl3.Load += new System.EventHandler(this.gridControl3_Load);
            // 
            // projectsBindingSource
            // 
            this.projectsBindingSource.DataMember = "Projects";
            this.projectsBindingSource.DataSource = this.workload_Tracking_System_DBDataSet;
            // 
            // gridView3
            // 
            this.gridView3.Appearance.HeaderPanel.Options.UseTextOptions = true;
            this.gridView3.Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.gridView3.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.gridView3.Appearance.SelectedRow.Options.UseBackColor = true;
            this.gridView3.AppearancePrint.HeaderPanel.Options.UseTextOptions = true;
            this.gridView3.AppearancePrint.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.gridView3.ChildGridLevelName = "Components";
            this.gridView3.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID1,
            this.colJobNumber1,
            this.colProjectNumber2,
            this.colStagePV,
            this.colCustomer1,
            this.colProject,
            this.colDueDate1,
            this.colPriority,
            this.colStatus1,
            this.colDesigner1,
            this.colToolMaker2,
            this.colRoughProgrammer1,
            this.colElectrodeProgrammer1,
            this.colFinishProgrammer1,
            this.colApprentice,
            this.colEngineer1,
            this.colOverlapAllowed,
            this.colIncludeHours,
            this.colKanBanWorkbookPath,
            this.colPercentComplete1,
            this.colDateModified,
            this.colLastKanBanGenerationDate,
            this.colLatestFinishDate});
            gridFormatRule2.Column = this.colPercentComplete1;
            gridFormatRule2.ColumnApplyTo = this.colPercentComplete1;
            gridFormatRule2.Name = "PercentCompleteFormat";
            formatConditionRuleDataBar2.Maximum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            formatConditionRuleDataBar2.MaximumType = DevExpress.XtraEditors.FormatConditionValueType.Number;
            formatConditionRuleDataBar2.MinimumType = DevExpress.XtraEditors.FormatConditionValueType.Number;
            formatConditionRuleDataBar2.PredefinedName = "Blue";
            gridFormatRule2.Rule = formatConditionRuleDataBar2;
            gridFormatRule3.Column = this.colDueDate1;
            gridFormatRule3.ColumnApplyTo = this.colDueDate1;
            gridFormatRule3.Name = "DueDateViolated";
            formatConditionRuleExpression1.Appearance.BackColor = System.Drawing.Color.Red;
            formatConditionRuleExpression1.Appearance.ForeColor = System.Drawing.Color.White;
            formatConditionRuleExpression1.Appearance.Options.UseBackColor = true;
            formatConditionRuleExpression1.Appearance.Options.UseForeColor = true;
            gridFormatRule3.Rule = formatConditionRuleExpression1;
            this.gridView3.FormatRules.Add(gridFormatRule2);
            this.gridView3.FormatRules.Add(gridFormatRule3);
            this.gridView3.GridControl = this.gridControl3;
            this.gridView3.Name = "gridView3";
            this.gridView3.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.True;
            this.gridView3.OptionsBehavior.EditingMode = DevExpress.XtraGrid.Views.Grid.GridEditingMode.Inplace;
            this.gridView3.OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.MouseDownFocused;
            this.gridView3.OptionsPrint.AllowMultilineHeaders = true;
            this.gridView3.OptionsPrint.AutoWidth = false;
            this.gridView3.OptionsPrint.PrintSelectedRowsOnly = true;
            this.gridView3.OptionsSelection.MultiSelect = true;
            this.gridView3.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
            this.gridView3.OptionsView.AllowHtmlDrawHeaders = true;
            this.gridView3.OptionsView.ColumnAutoWidth = false;
            this.gridView3.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.True;
            this.gridView3.OptionsView.RowAutoHeight = true;
            this.gridView3.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.colDueDate1, DevExpress.Data.ColumnSortOrder.Ascending)});
            this.gridView3.RowCellStyle += new DevExpress.XtraGrid.Views.Grid.RowCellStyleEventHandler(this.gridView3_RowCellStyle);
            this.gridView3.RowStyle += new DevExpress.XtraGrid.Views.Grid.RowStyleEventHandler(this.gridView3_RowStyle);
            this.gridView3.MasterRowExpanded += new DevExpress.XtraGrid.Views.Grid.CustomMasterRowEventHandler(this.gridView_MasterRowExpanded);
            this.gridView3.MasterRowCollapsed += new DevExpress.XtraGrid.Views.Grid.CustomMasterRowEventHandler(this.gridView_MasterRowCollapsed);
            this.gridView3.ShownEditor += new System.EventHandler(this.gridView3_ShownEditor);
            this.gridView3.CellValueChanged += new DevExpress.XtraGrid.Views.Base.CellValueChangedEventHandler(this.gridView3_CellValueChanged);
            this.gridView3.CustomRowFilter += new DevExpress.XtraGrid.Views.Base.RowFilterEventHandler(this.gridView3_CustomRowFilter);
            this.gridView3.PrintInitialize += new DevExpress.XtraGrid.Views.Base.PrintInitializeEventHandler(this.gridView3_PrintInitialize);
            this.gridView3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.gridView3_KeyDown);
            this.gridView3.ValidatingEditor += new DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventHandler(this.GridView3_ValidatingEditor);
            this.gridView3.InvalidValueException += new DevExpress.XtraEditors.Controls.InvalidValueExceptionEventHandler(this.GridView3_InvalidValueException);
            // 
            // colID1
            // 
            this.colID1.FieldName = "ID";
            this.colID1.Name = "colID1";
            // 
            // colJobNumber1
            // 
            this.colJobNumber1.FieldName = "JobNumber";
            this.colJobNumber1.Name = "colJobNumber1";
            this.colJobNumber1.Visible = true;
            this.colJobNumber1.VisibleIndex = 0;
            this.colJobNumber1.Width = 120;
            // 
            // colProjectNumber2
            // 
            this.colProjectNumber2.FieldName = "ProjectNumber";
            this.colProjectNumber2.Name = "colProjectNumber2";
            this.colProjectNumber2.Visible = true;
            this.colProjectNumber2.VisibleIndex = 1;
            this.colProjectNumber2.Width = 60;
            // 
            // colStagePV
            // 
            this.colStagePV.Caption = "Stage";
            this.colStagePV.FieldName = "Stage";
            this.colStagePV.Name = "colStagePV";
            // 
            // colCustomer1
            // 
            this.colCustomer1.FieldName = "Customer";
            this.colCustomer1.Name = "colCustomer1";
            this.colCustomer1.Visible = true;
            this.colCustomer1.VisibleIndex = 2;
            this.colCustomer1.Width = 123;
            // 
            // colProject
            // 
            this.colProject.Caption = "Part Name / Project";
            this.colProject.FieldName = "Project";
            this.colProject.Name = "colProject";
            this.colProject.Visible = true;
            this.colProject.VisibleIndex = 3;
            this.colProject.Width = 140;
            // 
            // colDueDate1
            // 
            this.colDueDate1.AppearanceCell.Options.UseTextOptions = true;
            this.colDueDate1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colDueDate1.FieldName = "DueDate";
            this.colDueDate1.Name = "colDueDate1";
            this.colDueDate1.Visible = true;
            this.colDueDate1.VisibleIndex = 4;
            this.colDueDate1.Width = 65;
            // 
            // colPriority
            // 
            this.colPriority.FieldName = "Priority";
            this.colPriority.Name = "colPriority";
            this.colPriority.Width = 105;
            // 
            // colStatus1
            // 
            this.colStatus1.FieldName = "Status";
            this.colStatus1.Name = "colStatus1";
            this.colStatus1.Visible = true;
            this.colStatus1.VisibleIndex = 5;
            this.colStatus1.Width = 93;
            // 
            // colDesigner1
            // 
            this.colDesigner1.ColumnEdit = this.personnelComboBoxEdit;
            this.colDesigner1.FieldName = "Designer";
            this.colDesigner1.Name = "colDesigner1";
            this.colDesigner1.Visible = true;
            this.colDesigner1.VisibleIndex = 7;
            this.colDesigner1.Width = 105;
            // 
            // personnelComboBoxEdit
            // 
            this.personnelComboBoxEdit.AutoHeight = false;
            this.personnelComboBoxEdit.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.personnelComboBoxEdit.Name = "personnelComboBoxEdit";
            // 
            // colToolMaker2
            // 
            this.colToolMaker2.ColumnEdit = this.personnelComboBoxEdit;
            this.colToolMaker2.FieldName = "ToolMaker";
            this.colToolMaker2.Name = "colToolMaker2";
            this.colToolMaker2.Visible = true;
            this.colToolMaker2.VisibleIndex = 8;
            this.colToolMaker2.Width = 105;
            // 
            // colRoughProgrammer1
            // 
            this.colRoughProgrammer1.ColumnEdit = this.personnelComboBoxEdit;
            this.colRoughProgrammer1.FieldName = "RoughProgrammer";
            this.colRoughProgrammer1.Name = "colRoughProgrammer1";
            this.colRoughProgrammer1.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colRoughProgrammer1.Visible = true;
            this.colRoughProgrammer1.VisibleIndex = 9;
            this.colRoughProgrammer1.Width = 105;
            // 
            // colElectrodeProgrammer1
            // 
            this.colElectrodeProgrammer1.ColumnEdit = this.personnelComboBoxEdit;
            this.colElectrodeProgrammer1.FieldName = "ElectrodeProgrammer";
            this.colElectrodeProgrammer1.Name = "colElectrodeProgrammer1";
            this.colElectrodeProgrammer1.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colElectrodeProgrammer1.Visible = true;
            this.colElectrodeProgrammer1.VisibleIndex = 10;
            this.colElectrodeProgrammer1.Width = 105;
            // 
            // colFinishProgrammer1
            // 
            this.colFinishProgrammer1.ColumnEdit = this.personnelComboBoxEdit;
            this.colFinishProgrammer1.FieldName = "FinishProgrammer";
            this.colFinishProgrammer1.Name = "colFinishProgrammer1";
            this.colFinishProgrammer1.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colFinishProgrammer1.Visible = true;
            this.colFinishProgrammer1.VisibleIndex = 11;
            this.colFinishProgrammer1.Width = 105;
            // 
            // colApprentice
            // 
            this.colApprentice.ColumnEdit = this.personnelComboBoxEdit;
            this.colApprentice.FieldName = "Apprentice";
            this.colApprentice.Name = "colApprentice";
            this.colApprentice.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colApprentice.Visible = true;
            this.colApprentice.VisibleIndex = 12;
            this.colApprentice.Width = 116;
            // 
            // colEngineer1
            // 
            this.colEngineer1.FieldName = "Engineer";
            this.colEngineer1.Name = "colEngineer1";
            this.colEngineer1.Width = 105;
            // 
            // colOverlapAllowed
            // 
            this.colOverlapAllowed.FieldName = "OverlapAllowed";
            this.colOverlapAllowed.Name = "colOverlapAllowed";
            this.colOverlapAllowed.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colOverlapAllowed.Visible = true;
            this.colOverlapAllowed.VisibleIndex = 13;
            this.colOverlapAllowed.Width = 61;
            // 
            // colIncludeHours
            // 
            this.colIncludeHours.FieldName = "IncludeHours";
            this.colIncludeHours.Name = "colIncludeHours";
            this.colIncludeHours.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colIncludeHours.Visible = true;
            this.colIncludeHours.VisibleIndex = 14;
            // 
            // colKanBanWorkbookPath
            // 
            this.colKanBanWorkbookPath.ColumnEdit = this.repositoryItemHyperLinkEdit2;
            this.colKanBanWorkbookPath.FieldName = "KanBanWorkbookPath";
            this.colKanBanWorkbookPath.Name = "colKanBanWorkbookPath";
            this.colKanBanWorkbookPath.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colKanBanWorkbookPath.Visible = true;
            this.colKanBanWorkbookPath.VisibleIndex = 15;
            this.colKanBanWorkbookPath.Width = 357;
            // 
            // repositoryItemHyperLinkEdit2
            // 
            this.repositoryItemHyperLinkEdit2.AutoHeight = false;
            this.repositoryItemHyperLinkEdit2.LinkColor = System.Drawing.Color.Blue;
            this.repositoryItemHyperLinkEdit2.Name = "repositoryItemHyperLinkEdit2";
            this.repositoryItemHyperLinkEdit2.SingleClick = true;
            // 
            // colPercentComplete1
            // 
            this.colPercentComplete1.DisplayFormat.FormatString = "P0";
            this.colPercentComplete1.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.colPercentComplete1.FieldName = "PercentComplete";
            this.colPercentComplete1.Name = "colPercentComplete1";
            this.colPercentComplete1.OptionsColumn.AllowEdit = false;
            this.colPercentComplete1.Visible = true;
            this.colPercentComplete1.VisibleIndex = 6;
            this.colPercentComplete1.Width = 108;
            // 
            // colDateModified
            // 
            this.colDateModified.FieldName = "DateModified";
            this.colDateModified.Name = "colDateModified";
            // 
            // colLastKanBanGenerationDate
            // 
            this.colLastKanBanGenerationDate.FieldName = "LastKanBanGenerationDate";
            this.colLastKanBanGenerationDate.Name = "colLastKanBanGenerationDate";
            // 
            // colLatestFinishDate
            // 
            this.colLatestFinishDate.Caption = "Latest Finish Date";
            this.colLatestFinishDate.FieldName = "LatestFinishDate";
            this.colLatestFinishDate.Name = "colLatestFinishDate";
            this.colLatestFinishDate.Visible = true;
            this.colLatestFinishDate.VisibleIndex = 16;
            // 
            // repositoryItemImageEdit2
            // 
            this.repositoryItemImageEdit2.AutoHeight = false;
            this.repositoryItemImageEdit2.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemImageEdit2.Name = "repositoryItemImageEdit2";
            this.repositoryItemImageEdit2.PictureStoreMode = DevExpress.XtraEditors.Controls.PictureStoreMode.ByteArray;
            this.repositoryItemImageEdit2.PopupFormSize = new System.Drawing.Size(600, 599);
            this.repositoryItemImageEdit2.Validating += new System.ComponentModel.CancelEventHandler(this.RepositoryItemImageEdit2_Validating);
            // 
            // repositoryItemPictureEdit1
            // 
            this.repositoryItemPictureEdit1.Name = "repositoryItemPictureEdit1";
            this.repositoryItemPictureEdit1.PictureStoreMode = DevExpress.XtraEditors.Controls.PictureStoreMode.ByteArray;
            // 
            // repositoryItemImageComboBox1
            // 
            this.repositoryItemImageComboBox1.AutoHeight = false;
            this.repositoryItemImageComboBox1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemImageComboBox1.Name = "repositoryItemImageComboBox1";
            // 
            // repositoryItemTextEdit2
            // 
            this.repositoryItemTextEdit2.AutoHeight = false;
            this.repositoryItemTextEdit2.Name = "repositoryItemTextEdit2";
            // 
            // stageComboBoxEdit
            // 
            this.stageComboBoxEdit.AutoHeight = false;
            this.stageComboBoxEdit.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.stageComboBoxEdit.Items.AddRange(new object[] {
            "1 - In-Design",
            "2 - In-Programming",
            "3 - In-Shop",
            "4 - In-Mold Check-In or Outside Vendors",
            "5 - Rework",
            "6 - In-Repair / Development",
            "7 - Completed",
            "8 - Quoted / Forecasted"});
            this.stageComboBoxEdit.Name = "stageComboBoxEdit";
            // 
            // genericDateEdit
            // 
            this.genericDateEdit.AutoHeight = false;
            this.genericDateEdit.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.genericDateEdit.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.genericDateEdit.Name = "genericDateEdit";
            // 
            // projectBandedGridView
            // 
            this.projectBandedGridView.Appearance.GroupRow.BackColor = System.Drawing.Color.LightBlue;
            this.projectBandedGridView.Appearance.GroupRow.ForeColor = System.Drawing.Color.Black;
            this.projectBandedGridView.Appearance.GroupRow.Options.UseBackColor = true;
            this.projectBandedGridView.Appearance.GroupRow.Options.UseForeColor = true;
            this.projectBandedGridView.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.projectBandedGridView.Appearance.SelectedRow.Options.UseBackColor = true;
            this.projectBandedGridView.AppearancePrint.BandPanel.BackColor = System.Drawing.Color.LightBlue;
            this.projectBandedGridView.AppearancePrint.BandPanel.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold);
            this.projectBandedGridView.AppearancePrint.BandPanel.Options.UseBackColor = true;
            this.projectBandedGridView.AppearancePrint.BandPanel.Options.UseFont = true;
            this.projectBandedGridView.AppearancePrint.EvenRow.BackColor = System.Drawing.Color.Gainsboro;
            this.projectBandedGridView.AppearancePrint.EvenRow.Options.UseBackColor = true;
            this.projectBandedGridView.AppearancePrint.GroupRow.BackColor = System.Drawing.Color.LightBlue;
            this.projectBandedGridView.AppearancePrint.GroupRow.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold);
            this.projectBandedGridView.AppearancePrint.GroupRow.Options.UseBackColor = true;
            this.projectBandedGridView.AppearancePrint.GroupRow.Options.UseFont = true;
            this.projectBandedGridView.AppearancePrint.HeaderPanel.BackColor = System.Drawing.Color.LightBlue;
            this.projectBandedGridView.AppearancePrint.HeaderPanel.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold);
            this.projectBandedGridView.AppearancePrint.HeaderPanel.ForeColor = System.Drawing.Color.Black;
            this.projectBandedGridView.AppearancePrint.HeaderPanel.Options.UseBackColor = true;
            this.projectBandedGridView.AppearancePrint.HeaderPanel.Options.UseFont = true;
            this.projectBandedGridView.AppearancePrint.HeaderPanel.Options.UseForeColor = true;
            this.projectBandedGridView.AppearancePrint.HeaderPanel.Options.UseTextOptions = true;
            this.projectBandedGridView.AppearancePrint.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.projectBandedGridView.Bands.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] {
            this.SegoeUI,
            this.milestonesGridBand,
            this.personnelGridBand,
            this.generalInfoGridBand});
            this.projectBandedGridView.Columns.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn[] {
            this.colIDBGV,
            this.colJobNumberBGV,
            this.colProjectNumberBGV,
            this.colStageBGV,
            this.colCustomerBGV,
            this.colProjectBGV,
            this.colMoldCostBGV,
            this.colStatusBGV,
            this.colDeliveryInWeeksBGV,
            this.colStartDateBGV,
            this.colDueDateBGV,
            this.colAdjustedDeliveryDateBGV,
            this.colEngineerBGV,
            this.colDesignerBGV,
            this.colToolMakerBGV,
            this.colRoughProgrammerBGV,
            this.colElectrodeProgrammerBGV,
            this.colFinishProgrammerBGV,
            this.colApprenticeBGV,
            this.colManifoldBGV,
            this.colMoldBaseBGV,
            this.colGeneralNotesBGV});
            this.projectBandedGridView.GridControl = this.gridControl3;
            this.projectBandedGridView.GroupCount = 1;
            this.projectBandedGridView.GroupSummary.AddRange(new DevExpress.XtraGrid.GridSummaryItem[] {
            new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Count, "Project", null, ", Count = {0}"),
            new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "MoldCost", this.colMoldCostBGV, "{0:c0}")});
            this.projectBandedGridView.Name = "projectBandedGridView";
            this.projectBandedGridView.OptionsBehavior.AutoExpandAllGroups = true;
            this.projectBandedGridView.OptionsPrint.AutoWidth = false;
            this.projectBandedGridView.OptionsPrint.EnableAppearanceEvenRow = true;
            this.projectBandedGridView.OptionsPrint.PrintSelectedRowsOnly = true;
            this.projectBandedGridView.OptionsSelection.MultiSelect = true;
            this.projectBandedGridView.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
            this.projectBandedGridView.OptionsSelection.ShowCheckBoxSelectorInGroupRow = DevExpress.Utils.DefaultBoolean.True;
            this.projectBandedGridView.OptionsView.ColumnAutoWidth = false;
            this.projectBandedGridView.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.True;
            this.projectBandedGridView.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Top;
            this.projectBandedGridView.OptionsView.ShowFooter = true;
            this.projectBandedGridView.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.colStageBGV, DevExpress.Data.ColumnSortOrder.Ascending)});
            this.projectBandedGridView.RowCellStyle += new DevExpress.XtraGrid.Views.Grid.RowCellStyleEventHandler(this.projectBandedGridView_RowCellStyle);
            this.projectBandedGridView.CustomRowCellEditForEditing += new DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventHandler(this.projectBandedGridView_CustomRowCellEditForEditing);
            this.projectBandedGridView.ShownEditor += new System.EventHandler(this.projectBandedGridView_ShownEditor);
            this.projectBandedGridView.CellValueChanged += new DevExpress.XtraGrid.Views.Base.CellValueChangedEventHandler(this.projectBandedGridView_CellValueChanged);
            this.projectBandedGridView.InvalidRowException += new DevExpress.XtraGrid.Views.Base.InvalidRowExceptionEventHandler(this.projectBandedGridView_InvalidRowException);
            this.projectBandedGridView.ValidateRow += new DevExpress.XtraGrid.Views.Base.ValidateRowEventHandler(this.projectBandedGridView_ValidateRow);
            this.projectBandedGridView.RowUpdated += new DevExpress.XtraGrid.Views.Base.RowObjectEventHandler(this.projectBandedGridView_RowUpdated);
            this.projectBandedGridView.PrintInitialize += new DevExpress.XtraGrid.Views.Base.PrintInitializeEventHandler(this.projectBandedGridView_PrintInitialize);
            this.projectBandedGridView.KeyDown += new System.Windows.Forms.KeyEventHandler(this.projectBandedGridView_KeyDown);
            this.projectBandedGridView.MouseDown += new System.Windows.Forms.MouseEventHandler(this.projectBandedGridView_MouseDown);
            this.projectBandedGridView.ValidatingEditor += new DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventHandler(this.projectBandedGridView_ValidatingEditor);
            // 
            // SegoeUI
            // 
            this.SegoeUI.Caption = "Project";
            this.SegoeUI.Columns.Add(this.colJobNumberBGV);
            this.SegoeUI.Columns.Add(this.colProjectNumberBGV);
            this.SegoeUI.Columns.Add(this.colStageBGV);
            this.SegoeUI.Columns.Add(this.colCustomerBGV);
            this.SegoeUI.Columns.Add(this.colProjectBGV);
            this.SegoeUI.Columns.Add(this.colMoldCostBGV);
            this.SegoeUI.Columns.Add(this.colDeliveryInWeeksBGV);
            this.SegoeUI.Name = "SegoeUI";
            this.SegoeUI.VisibleIndex = 0;
            this.SegoeUI.Width = 565;
            // 
            // colJobNumberBGV
            // 
            this.colJobNumberBGV.Caption = "Job #";
            this.colJobNumberBGV.FieldName = "JobNumber";
            this.colJobNumberBGV.Name = "colJobNumberBGV";
            this.colJobNumberBGV.Visible = true;
            this.colJobNumberBGV.Width = 99;
            // 
            // colProjectNumberBGV
            // 
            this.colProjectNumberBGV.Caption = "Project #";
            this.colProjectNumberBGV.FieldName = "ProjectNumber";
            this.colProjectNumberBGV.Name = "colProjectNumberBGV";
            this.colProjectNumberBGV.Visible = true;
            this.colProjectNumberBGV.Width = 69;
            // 
            // colStageBGV
            // 
            this.colStageBGV.Caption = "Stage";
            this.colStageBGV.ColumnEdit = this.stageComboBoxEdit;
            this.colStageBGV.FieldName = "Stage";
            this.colStageBGV.Name = "colStageBGV";
            this.colStageBGV.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colStageBGV.Visible = true;
            this.colStageBGV.Width = 65;
            // 
            // colCustomerBGV
            // 
            this.colCustomerBGV.Caption = "Customer";
            this.colCustomerBGV.FieldName = "Customer";
            this.colCustomerBGV.Name = "colCustomerBGV";
            this.colCustomerBGV.Visible = true;
            this.colCustomerBGV.Width = 99;
            // 
            // colProjectBGV
            // 
            this.colProjectBGV.Caption = "Part Name / Project";
            this.colProjectBGV.FieldName = "Project";
            this.colProjectBGV.Name = "colProjectBGV";
            this.colProjectBGV.Visible = true;
            this.colProjectBGV.Width = 99;
            // 
            // colMoldCostBGV
            // 
            this.colMoldCostBGV.Caption = "Mold Cost";
            this.colMoldCostBGV.DisplayFormat.FormatString = "c0";
            this.colMoldCostBGV.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.colMoldCostBGV.FieldName = "MoldCost";
            this.colMoldCostBGV.Name = "colMoldCostBGV";
            this.colMoldCostBGV.Visible = true;
            this.colMoldCostBGV.Width = 69;
            // 
            // colDeliveryInWeeksBGV
            // 
            this.colDeliveryInWeeksBGV.Caption = "Delivery In Weeks";
            this.colDeliveryInWeeksBGV.FieldName = "DeliveryInWeeks";
            this.colDeliveryInWeeksBGV.Name = "colDeliveryInWeeksBGV";
            this.colDeliveryInWeeksBGV.Visible = true;
            this.colDeliveryInWeeksBGV.Width = 65;
            // 
            // milestonesGridBand
            // 
            this.milestonesGridBand.Caption = "Milestones";
            this.milestonesGridBand.Columns.Add(this.colStatusBGV);
            this.milestonesGridBand.Columns.Add(this.colStartDateBGV);
            this.milestonesGridBand.Columns.Add(this.colDueDateBGV);
            this.milestonesGridBand.Columns.Add(this.colAdjustedDeliveryDateBGV);
            this.milestonesGridBand.Name = "milestonesGridBand";
            this.milestonesGridBand.VisibleIndex = 1;
            this.milestonesGridBand.Width = 207;
            // 
            // colStatusBGV
            // 
            this.colStatusBGV.Caption = "Status";
            this.colStatusBGV.FieldName = "Status";
            this.colStatusBGV.Name = "colStatusBGV";
            // 
            // colStartDateBGV
            // 
            this.colStartDateBGV.Caption = "Start Date";
            this.colStartDateBGV.ColumnEdit = this.genericDateEdit;
            this.colStartDateBGV.FieldName = "StartDate";
            this.colStartDateBGV.Name = "colStartDateBGV";
            this.colStartDateBGV.Visible = true;
            this.colStartDateBGV.Width = 69;
            // 
            // colDueDateBGV
            // 
            this.colDueDateBGV.Caption = "Due Date";
            this.colDueDateBGV.ColumnEdit = this.genericDateEdit;
            this.colDueDateBGV.FieldName = "DueDate";
            this.colDueDateBGV.Name = "colDueDateBGV";
            this.colDueDateBGV.Visible = true;
            this.colDueDateBGV.Width = 69;
            // 
            // colAdjustedDeliveryDateBGV
            // 
            this.colAdjustedDeliveryDateBGV.Caption = "Adj. Delivery Date";
            this.colAdjustedDeliveryDateBGV.ColumnEdit = this.genericDateEdit;
            this.colAdjustedDeliveryDateBGV.FieldName = "AdjustedDeliveryDate";
            this.colAdjustedDeliveryDateBGV.Name = "colAdjustedDeliveryDateBGV";
            this.colAdjustedDeliveryDateBGV.Visible = true;
            this.colAdjustedDeliveryDateBGV.Width = 69;
            // 
            // personnelGridBand
            // 
            this.personnelGridBand.Caption = "Personnel";
            this.personnelGridBand.Columns.Add(this.colEngineerBGV);
            this.personnelGridBand.Columns.Add(this.colDesignerBGV);
            this.personnelGridBand.Columns.Add(this.colToolMakerBGV);
            this.personnelGridBand.Columns.Add(this.colRoughProgrammerBGV);
            this.personnelGridBand.Columns.Add(this.colElectrodeProgrammerBGV);
            this.personnelGridBand.Columns.Add(this.colFinishProgrammerBGV);
            this.personnelGridBand.Columns.Add(this.colApprenticeBGV);
            this.personnelGridBand.Name = "personnelGridBand";
            this.personnelGridBand.VisibleIndex = 2;
            this.personnelGridBand.Width = 454;
            // 
            // colEngineerBGV
            // 
            this.colEngineerBGV.Caption = "Engineer";
            this.colEngineerBGV.ColumnEdit = this.personnelComboBoxEdit;
            this.colEngineerBGV.FieldName = "Engineer";
            this.colEngineerBGV.Name = "colEngineerBGV";
            this.colEngineerBGV.Visible = true;
            this.colEngineerBGV.Width = 64;
            // 
            // colDesignerBGV
            // 
            this.colDesignerBGV.Caption = "Designer";
            this.colDesignerBGV.ColumnEdit = this.personnelComboBoxEdit;
            this.colDesignerBGV.FieldName = "Designer";
            this.colDesignerBGV.Name = "colDesignerBGV";
            this.colDesignerBGV.Visible = true;
            this.colDesignerBGV.Width = 65;
            // 
            // colToolMakerBGV
            // 
            this.colToolMakerBGV.Caption = "Tool Maker";
            this.colToolMakerBGV.ColumnEdit = this.personnelComboBoxEdit;
            this.colToolMakerBGV.FieldName = "ToolMaker";
            this.colToolMakerBGV.Name = "colToolMakerBGV";
            this.colToolMakerBGV.Visible = true;
            this.colToolMakerBGV.Width = 65;
            // 
            // colRoughProgrammerBGV
            // 
            this.colRoughProgrammerBGV.Caption = "Rough Programmer";
            this.colRoughProgrammerBGV.ColumnEdit = this.personnelComboBoxEdit;
            this.colRoughProgrammerBGV.FieldName = "RoughProgrammer";
            this.colRoughProgrammerBGV.Name = "colRoughProgrammerBGV";
            this.colRoughProgrammerBGV.Visible = true;
            this.colRoughProgrammerBGV.Width = 65;
            // 
            // colElectrodeProgrammerBGV
            // 
            this.colElectrodeProgrammerBGV.Caption = "Electrode Programmer";
            this.colElectrodeProgrammerBGV.ColumnEdit = this.personnelComboBoxEdit;
            this.colElectrodeProgrammerBGV.FieldName = "ElectrodeProgrammer";
            this.colElectrodeProgrammerBGV.Name = "colElectrodeProgrammerBGV";
            this.colElectrodeProgrammerBGV.Visible = true;
            this.colElectrodeProgrammerBGV.Width = 65;
            // 
            // colFinishProgrammerBGV
            // 
            this.colFinishProgrammerBGV.Caption = "Finish Programmer";
            this.colFinishProgrammerBGV.ColumnEdit = this.personnelComboBoxEdit;
            this.colFinishProgrammerBGV.FieldName = "FinishProgrammer";
            this.colFinishProgrammerBGV.Name = "colFinishProgrammerBGV";
            this.colFinishProgrammerBGV.Visible = true;
            this.colFinishProgrammerBGV.Width = 65;
            // 
            // colApprenticeBGV
            // 
            this.colApprenticeBGV.Caption = "Apprentice";
            this.colApprenticeBGV.ColumnEdit = this.personnelComboBoxEdit;
            this.colApprenticeBGV.FieldName = "Apprentice";
            this.colApprenticeBGV.Name = "colApprenticeBGV";
            this.colApprenticeBGV.Visible = true;
            this.colApprenticeBGV.Width = 65;
            // 
            // generalInfoGridBand
            // 
            this.generalInfoGridBand.Caption = "General Info";
            this.generalInfoGridBand.Columns.Add(this.colManifoldBGV);
            this.generalInfoGridBand.Columns.Add(this.colMoldBaseBGV);
            this.generalInfoGridBand.Columns.Add(this.colGeneralNotesBGV);
            this.generalInfoGridBand.Name = "generalInfoGridBand";
            this.generalInfoGridBand.VisibleIndex = 3;
            this.generalInfoGridBand.Width = 871;
            // 
            // colManifoldBGV
            // 
            this.colManifoldBGV.Caption = "Manifold";
            this.colManifoldBGV.FieldName = "Manifold";
            this.colManifoldBGV.Name = "colManifoldBGV";
            this.colManifoldBGV.Visible = true;
            // 
            // colMoldBaseBGV
            // 
            this.colMoldBaseBGV.Caption = "Mold Base";
            this.colMoldBaseBGV.FieldName = "Moldbase";
            this.colMoldBaseBGV.Name = "colMoldBaseBGV";
            this.colMoldBaseBGV.Visible = true;
            this.colMoldBaseBGV.Width = 71;
            // 
            // colGeneralNotesBGV
            // 
            this.colGeneralNotesBGV.Caption = "General Notes";
            this.colGeneralNotesBGV.FieldName = "GeneralNotes";
            this.colGeneralNotesBGV.Name = "colGeneralNotesBGV";
            this.colGeneralNotesBGV.Visible = true;
            this.colGeneralNotesBGV.Width = 725;
            // 
            // colIDBGV
            // 
            this.colIDBGV.FieldName = "ID";
            this.colIDBGV.Name = "colIDBGV";
            this.colIDBGV.Visible = true;
            // 
            // gridView5
            // 
            this.gridView5.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.gridView5.Appearance.SelectedRow.Options.UseBackColor = true;
            this.gridView5.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID4,
            this.colTaskName1,
            this.colResource1,
            this.colMachine,
            this.colHours,
            this.colDuration1,
            this.colStartDate2,
            this.colFinishDate2,
            this.colTaskID1,
            this.colPredecessors1,
            this.colNotes1,
            this.colStatus3,
            this.colInitials,
            this.colDateCompleted});
            this.gridView5.GridControl = this.gridControl3;
            this.gridView5.Name = "gridView5";
            this.gridView5.OptionsView.ColumnAutoWidth = false;
            this.gridView5.MasterRowExpanded += new DevExpress.XtraGrid.Views.Grid.CustomMasterRowEventHandler(this.gridView_MasterRowExpanded);
            this.gridView5.MasterRowCollapsed += new DevExpress.XtraGrid.Views.Grid.CustomMasterRowEventHandler(this.gridView_MasterRowCollapsed);
            this.gridView5.CellValueChanged += new DevExpress.XtraGrid.Views.Base.CellValueChangedEventHandler(this.gridView5_CellValueChanged);
            this.gridView5.ValidatingEditor += new DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventHandler(this.GridView5_ValidatingEditor);
            this.gridView5.InvalidValueException += new DevExpress.XtraEditors.Controls.InvalidValueExceptionEventHandler(this.GridView5_InvalidValueException);
            // 
            // colID4
            // 
            this.colID4.FieldName = "ID";
            this.colID4.Name = "colID4";
            // 
            // colTaskName1
            // 
            this.colTaskName1.FieldName = "TaskName";
            this.colTaskName1.Name = "colTaskName1";
            this.colTaskName1.Visible = true;
            this.colTaskName1.VisibleIndex = 0;
            this.colTaskName1.Width = 155;
            // 
            // colResource1
            // 
            this.colResource1.FieldName = "Personnel";
            this.colResource1.Name = "colResource1";
            this.colResource1.Visible = true;
            this.colResource1.VisibleIndex = 9;
            this.colResource1.Width = 87;
            // 
            // colMachine
            // 
            this.colMachine.FieldName = "Machine";
            this.colMachine.Name = "colMachine";
            this.colMachine.Visible = true;
            this.colMachine.VisibleIndex = 8;
            this.colMachine.Width = 96;
            // 
            // colHours
            // 
            this.colHours.ColumnEdit = this.repositoryItemSpinEdit2;
            this.colHours.FieldName = "Hours";
            this.colHours.Name = "colHours";
            this.colHours.Visible = true;
            this.colHours.VisibleIndex = 1;
            this.colHours.Width = 41;
            // 
            // colDuration1
            // 
            this.colDuration1.FieldName = "Duration";
            this.colDuration1.Name = "colDuration1";
            this.colDuration1.Visible = true;
            this.colDuration1.VisibleIndex = 2;
            this.colDuration1.Width = 69;
            // 
            // colStartDate2
            // 
            this.colStartDate2.FieldName = "StartDate";
            this.colStartDate2.Name = "colStartDate2";
            this.colStartDate2.Visible = true;
            this.colStartDate2.VisibleIndex = 3;
            this.colStartDate2.Width = 73;
            // 
            // colFinishDate2
            // 
            this.colFinishDate2.FieldName = "FinishDate";
            this.colFinishDate2.Name = "colFinishDate2";
            this.colFinishDate2.Visible = true;
            this.colFinishDate2.VisibleIndex = 4;
            this.colFinishDate2.Width = 69;
            // 
            // colTaskID1
            // 
            this.colTaskID1.FieldName = "TaskID";
            this.colTaskID1.Name = "colTaskID1";
            this.colTaskID1.OptionsColumn.AllowEdit = false;
            this.colTaskID1.Visible = true;
            this.colTaskID1.VisibleIndex = 5;
            this.colTaskID1.Width = 45;
            // 
            // colPredecessors1
            // 
            this.colPredecessors1.FieldName = "Predecessors";
            this.colPredecessors1.Name = "colPredecessors1";
            this.colPredecessors1.Visible = true;
            this.colPredecessors1.VisibleIndex = 6;
            this.colPredecessors1.Width = 96;
            // 
            // colNotes1
            // 
            this.colNotes1.FieldName = "Notes";
            this.colNotes1.Name = "colNotes1";
            this.colNotes1.Visible = true;
            this.colNotes1.VisibleIndex = 7;
            this.colNotes1.Width = 273;
            // 
            // colStatus3
            // 
            this.colStatus3.FieldName = "Status";
            this.colStatus3.Name = "colStatus3";
            this.colStatus3.OptionsColumn.AllowEdit = false;
            this.colStatus3.Visible = true;
            this.colStatus3.VisibleIndex = 10;
            this.colStatus3.Width = 89;
            // 
            // colInitials
            // 
            this.colInitials.FieldName = "Initials";
            this.colInitials.Name = "colInitials";
            this.colInitials.OptionsColumn.AllowEdit = false;
            this.colInitials.Visible = true;
            this.colInitials.VisibleIndex = 11;
            this.colInitials.Width = 61;
            // 
            // colDateCompleted
            // 
            this.colDateCompleted.FieldName = "DateCompleted";
            this.colDateCompleted.Name = "colDateCompleted";
            this.colDateCompleted.OptionsColumn.AllowEdit = false;
            this.colDateCompleted.Visible = true;
            this.colDateCompleted.VisibleIndex = 12;
            this.colDateCompleted.Width = 95;
            // 
            // DeptProgressGridView
            // 
            this.DeptProgressGridView.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.DepartmentColDPV,
            this.PercentCompleteColDPV});
            gridFormatRule4.Column = this.PercentCompleteColDPV;
            gridFormatRule4.ColumnApplyTo = this.PercentCompleteColDPV;
            gridFormatRule4.Name = "PercentCompleteFormat";
            formatConditionRuleDataBar3.Maximum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            formatConditionRuleDataBar3.MaximumType = DevExpress.XtraEditors.FormatConditionValueType.Number;
            formatConditionRuleDataBar3.MinimumType = DevExpress.XtraEditors.FormatConditionValueType.Number;
            formatConditionRuleDataBar3.PredefinedName = "Blue";
            gridFormatRule4.Rule = formatConditionRuleDataBar3;
            this.DeptProgressGridView.FormatRules.Add(gridFormatRule4);
            this.DeptProgressGridView.GridControl = this.gridControl3;
            this.DeptProgressGridView.Name = "DeptProgressGridView";
            this.DeptProgressGridView.OptionsView.ColumnAutoWidth = false;
            // 
            // DepartmentColDPV
            // 
            this.DepartmentColDPV.Caption = "Department";
            this.DepartmentColDPV.FieldName = "Department";
            this.DepartmentColDPV.Name = "DepartmentColDPV";
            this.DepartmentColDPV.Visible = true;
            this.DepartmentColDPV.VisibleIndex = 0;
            this.DepartmentColDPV.Width = 125;
            // 
            // PercentCompleteColDPV
            // 
            this.PercentCompleteColDPV.Caption = "Percent Complete";
            this.PercentCompleteColDPV.DisplayFormat.FormatString = "p0";
            this.PercentCompleteColDPV.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.PercentCompleteColDPV.FieldName = "PercentComplete";
            this.PercentCompleteColDPV.Name = "PercentCompleteColDPV";
            this.PercentCompleteColDPV.OptionsColumn.AllowEdit = false;
            this.PercentCompleteColDPV.Visible = true;
            this.PercentCompleteColDPV.VisibleIndex = 1;
            this.PercentCompleteColDPV.Width = 125;
            // 
            // projectsTableAdapter
            // 
            this.projectsTableAdapter.ClearBeforeFill = true;
            // 
            // tasksTableAdapter
            // 
            this.tasksTableAdapter.ClearBeforeFill = true;
            // 
            // schedulerStorage1
            // 
            this.schedulerStorage1.Appointments.ResourceSharing = true;
            this.schedulerStorage1.AppointmentChanging += new DevExpress.XtraScheduler.PersistentObjectCancelEventHandler(this.schedulerStorage1_AppointmentChanging);
            this.schedulerStorage1.AppointmentsChanged += new DevExpress.XtraScheduler.PersistentObjectsEventHandler(this.schedulerStorage1_AppointmentsChanged);
            this.schedulerStorage1.FilterAppointment += new DevExpress.XtraScheduler.PersistentObjectCancelEventHandler(this.schedulerStorage1_FilterAppointment);
            this.schedulerStorage1.FilterResource += new DevExpress.XtraScheduler.PersistentObjectCancelEventHandler(this.schedulerStorage1_FilterResource);
            // 
            // xtraTabControl1
            // 
            this.xtraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.xtraTabControl1.Appearance.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.xtraTabControl1.Appearance.Options.UseFont = true;
            this.xtraTabControl1.Location = new System.Drawing.Point(13, 12);
            this.xtraTabControl1.Margin = new System.Windows.Forms.Padding(4);
            this.xtraTabControl1.Name = "xtraTabControl1";
            this.xtraTabControl1.SelectedTabPage = this.xtraTabPage1;
            this.xtraTabControl1.Size = new System.Drawing.Size(1602, 827);
            this.xtraTabControl1.TabIndex = 0;
            this.xtraTabControl1.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
            this.xtraTabPage1,
            this.xtraTabPage2,
            this.xtraTabPage7,
            this.xtraTabPage3,
            this.xtraTabPage4});
            // 
            // xtraTabPage1
            // 
            this.xtraTabPage1.Controls.Add(this.includeCompletesCheckEdit);
            this.xtraTabPage1.Controls.Add(this.includeQuotesCheckEdit);
            this.xtraTabPage1.Controls.Add(this.labelControl7);
            this.xtraTabPage1.Controls.Add(this.projectCheckedComboBoxEdit);
            this.xtraTabPage1.Controls.Add(this.labelControl3);
            this.xtraTabPage1.Controls.Add(this.GroupByRadioGroup);
            this.xtraTabPage1.Controls.Add(this.labelControl1);
            this.xtraTabPage1.Controls.Add(this.refreshButton);
            this.xtraTabPage1.Controls.Add(this.schedulerControl1);
            this.xtraTabPage1.Controls.Add(this.departmentComboBox);
            this.xtraTabPage1.Margin = new System.Windows.Forms.Padding(4);
            this.xtraTabPage1.Name = "xtraTabPage1";
            this.xtraTabPage1.Size = new System.Drawing.Size(1596, 799);
            this.xtraTabPage1.Text = "Department Schedule View";
            // 
            // includeCompletesCheckEdit
            // 
            this.includeCompletesCheckEdit.Location = new System.Drawing.Point(899, 10);
            this.includeCompletesCheckEdit.Name = "includeCompletesCheckEdit";
            this.includeCompletesCheckEdit.Properties.Caption = "Include Completed Tasks";
            this.includeCompletesCheckEdit.Size = new System.Drawing.Size(157, 19);
            this.includeCompletesCheckEdit.TabIndex = 11;
            this.includeCompletesCheckEdit.CheckStateChanged += new System.EventHandler(this.includeCompletesCheckEdit_CheckStateChanged);
            // 
            // includeQuotesCheckEdit
            // 
            this.includeQuotesCheckEdit.Location = new System.Drawing.Point(786, 10);
            this.includeQuotesCheckEdit.Name = "includeQuotesCheckEdit";
            this.includeQuotesCheckEdit.Properties.Caption = "Include Quotes";
            this.includeQuotesCheckEdit.Size = new System.Drawing.Size(107, 19);
            this.includeQuotesCheckEdit.TabIndex = 10;
            this.includeQuotesCheckEdit.CheckedChanged += new System.EventHandler(this.includeQuotesCheckEdit_CheckedChanged);
            // 
            // labelControl7
            // 
            this.labelControl7.Appearance.Options.UseFont = true;
            this.labelControl7.Location = new System.Drawing.Point(232, 10);
            this.labelControl7.Name = "labelControl7";
            this.labelControl7.Size = new System.Drawing.Size(43, 13);
            this.labelControl7.TabIndex = 9;
            this.labelControl7.Text = "Projects:";
            // 
            // projectCheckedComboBoxEdit
            // 
            this.projectCheckedComboBoxEdit.Location = new System.Drawing.Point(292, 7);
            this.projectCheckedComboBoxEdit.Name = "projectCheckedComboBoxEdit";
            this.projectCheckedComboBoxEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.projectCheckedComboBoxEdit.Size = new System.Drawing.Size(120, 20);
            this.projectCheckedComboBoxEdit.TabIndex = 8;
            this.projectCheckedComboBoxEdit.EditValueChanged += new System.EventHandler(this.projectCheckedComboBoxEdit_EditValueChanged);
            // 
            // labelControl3
            // 
            this.labelControl3.Location = new System.Drawing.Point(517, 9);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(50, 13);
            this.labelControl3.TabIndex = 7;
            this.labelControl3.Text = "Group By:";
            // 
            // GroupByRadioGroup
            // 
            this.GroupByRadioGroup.Location = new System.Drawing.Point(581, 4);
            this.GroupByRadioGroup.Name = "GroupByRadioGroup";
            this.GroupByRadioGroup.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.GroupByRadioGroup.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
            new DevExpress.XtraEditors.Controls.RadioGroupItem(true, "Resource"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem(false, "None")});
            this.GroupByRadioGroup.Size = new System.Drawing.Size(167, 28);
            this.GroupByRadioGroup.TabIndex = 6;
            this.GroupByRadioGroup.SelectedIndexChanged += new System.EventHandler(this.GroupByRadioGroup_SelectedIndexChanged);
            // 
            // labelControl1
            // 
            this.labelControl1.Appearance.Options.UseFont = true;
            this.labelControl1.Location = new System.Drawing.Point(12, 9);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(64, 13);
            this.labelControl1.TabIndex = 5;
            this.labelControl1.Text = "Department:";
            // 
            // refreshButton
            // 
            this.refreshButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.refreshButton.Location = new System.Drawing.Point(425, 6);
            this.refreshButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.refreshButton.Name = "refreshButton";
            this.refreshButton.Size = new System.Drawing.Size(69, 21);
            this.refreshButton.TabIndex = 4;
            this.refreshButton.Text = "Refresh";
            this.refreshButton.Click += new System.EventHandler(this.refreshButton_Click);
            // 
            // schedulerControl1
            // 
            this.schedulerControl1.ActiveViewType = DevExpress.XtraScheduler.SchedulerViewType.Gantt;
            this.schedulerControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.schedulerControl1.DataStorage = this.schedulerStorage1;
            this.schedulerControl1.GroupType = DevExpress.XtraScheduler.SchedulerGroupType.Resource;
            this.schedulerControl1.Location = new System.Drawing.Point(12, 36);
            this.schedulerControl1.LookAndFeel.SkinName = "DevExpress Dark Style";
            this.schedulerControl1.Margin = new System.Windows.Forms.Padding(4);
            this.schedulerControl1.Name = "schedulerControl1";
            this.schedulerControl1.Size = new System.Drawing.Size(1573, 750);
            this.schedulerControl1.Start = new System.DateTime(2018, 12, 14, 0, 0, 0, 0);
            this.schedulerControl1.TabIndex = 2;
            this.schedulerControl1.Text = "schedulerControl1";
            this.schedulerControl1.Views.DayView.TimeRulers.Add(timeRuler1);
            this.schedulerControl1.Views.FullWeekView.Enabled = true;
            this.schedulerControl1.Views.FullWeekView.TimeRulers.Add(timeRuler2);
            this.schedulerControl1.Views.TimelineView.ResourcesPerPage = 5;
            this.schedulerControl1.Views.WeekView.Enabled = false;
            this.schedulerControl1.Views.WorkWeekView.TimeRulers.Add(timeRuler3);
            this.schedulerControl1.AppointmentDrop += new DevExpress.XtraScheduler.AppointmentDragEventHandler(this.schedulerControl1_AppointmentDrop);
            this.schedulerControl1.AppointmentResized += new DevExpress.XtraScheduler.AppointmentResizeEventHandler(this.schedulerControl1_AppointmentResized);
            this.schedulerControl1.AppointmentFlyoutShowing += new DevExpress.XtraScheduler.AppointmentFlyoutShowingEventHandler(this.schedulerControl1_AppointmentFlyoutShowing);
            this.schedulerControl1.DragDrop += new System.Windows.Forms.DragEventHandler(this.schedulerControl1_DragDrop);
            // 
            // departmentComboBox
            // 
            this.departmentComboBox.EditValue = "CNC Rough";
            this.departmentComboBox.Location = new System.Drawing.Point(89, 7);
            this.departmentComboBox.Margin = new System.Windows.Forms.Padding(4);
            this.departmentComboBox.Name = "departmentComboBox";
            this.departmentComboBox.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.departmentComboBox.Size = new System.Drawing.Size(127, 20);
            this.departmentComboBox.TabIndex = 2;
            this.departmentComboBox.SelectedIndexChanged += new System.EventHandler(this.departmentComboBox_SelectedIndexChanged);
            // 
            // xtraTabPage2
            // 
            this.xtraTabPage2.Controls.Add(this.PrintEmployeeWorkCheckedComboBoxEdit);
            this.xtraTabPage2.Controls.Add(this.labelControl8);
            this.xtraTabPage2.Controls.Add(this.daysAheadSpinEdit);
            this.xtraTabPage2.Controls.Add(this.filterTasksByDatesCheckEdit);
            this.xtraTabPage2.Controls.Add(this.printEmployeeWorkButton);
            this.xtraTabPage2.Controls.Add(this.labelControl2);
            this.xtraTabPage2.Controls.Add(this.PrintDeptsCheckedComboBoxEdit);
            this.xtraTabPage2.Controls.Add(this.printTaskViewButton);
            this.xtraTabPage2.Controls.Add(this.RefreshTasksButton);
            this.xtraTabPage2.Controls.Add(this.departmentComboBox2);
            this.xtraTabPage2.Controls.Add(this.gridControl1);
            this.xtraTabPage2.Margin = new System.Windows.Forms.Padding(4);
            this.xtraTabPage2.Name = "xtraTabPage2";
            this.xtraTabPage2.Size = new System.Drawing.Size(1596, 799);
            this.xtraTabPage2.Text = "Department Task View";
            // 
            // PrintEmployeeWorkCheckedComboBoxEdit
            // 
            this.PrintEmployeeWorkCheckedComboBoxEdit.EditValue = "";
            this.PrintEmployeeWorkCheckedComboBoxEdit.Location = new System.Drawing.Point(652, 7);
            this.PrintEmployeeWorkCheckedComboBoxEdit.Name = "PrintEmployeeWorkCheckedComboBoxEdit";
            this.PrintEmployeeWorkCheckedComboBoxEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.PrintEmployeeWorkCheckedComboBoxEdit.Size = new System.Drawing.Size(114, 20);
            this.PrintEmployeeWorkCheckedComboBoxEdit.TabIndex = 16;
            // 
            // labelControl8
            // 
            this.labelControl8.Location = new System.Drawing.Point(289, 10);
            this.labelControl8.Name = "labelControl8";
            this.labelControl8.Size = new System.Drawing.Size(60, 13);
            this.labelControl8.TabIndex = 15;
            this.labelControl8.Text = "Days Ahead";
            // 
            // daysAheadSpinEdit
            // 
            this.daysAheadSpinEdit.EditValue = new decimal(new int[] {
            7,
            0,
            0,
            0});
            this.daysAheadSpinEdit.Location = new System.Drawing.Point(244, 7);
            this.daysAheadSpinEdit.Name = "daysAheadSpinEdit";
            this.daysAheadSpinEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.daysAheadSpinEdit.Size = new System.Drawing.Size(39, 20);
            this.daysAheadSpinEdit.TabIndex = 14;
            this.daysAheadSpinEdit.ValueChanged += new System.EventHandler(this.daysAheadSpinEdit_ValueChanged);
            // 
            // filterTasksByDatesCheckEdit
            // 
            this.filterTasksByDatesCheckEdit.EditValue = true;
            this.filterTasksByDatesCheckEdit.Location = new System.Drawing.Point(225, 7);
            this.filterTasksByDatesCheckEdit.Name = "filterTasksByDatesCheckEdit";
            this.filterTasksByDatesCheckEdit.Properties.Caption = "";
            this.filterTasksByDatesCheckEdit.Size = new System.Drawing.Size(20, 19);
            this.filterTasksByDatesCheckEdit.TabIndex = 13;
            this.filterTasksByDatesCheckEdit.CheckedChanged += new System.EventHandler(this.filterTasksByDatesCheckEdit_CheckedChanged);
            // 
            // printEmployeeWorkButton
            // 
            this.printEmployeeWorkButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.printEmployeeWorkButton.Location = new System.Drawing.Point(772, 6);
            this.printEmployeeWorkButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.printEmployeeWorkButton.Name = "printEmployeeWorkButton";
            this.printEmployeeWorkButton.Size = new System.Drawing.Size(69, 21);
            this.printEmployeeWorkButton.TabIndex = 11;
            this.printEmployeeWorkButton.Text = "Print";
            this.printEmployeeWorkButton.Click += new System.EventHandler(this.printEmployeeWorkButton_Click);
            // 
            // labelControl2
            // 
            this.labelControl2.Appearance.Options.UseFont = true;
            this.labelControl2.Location = new System.Drawing.Point(12, 9);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(64, 13);
            this.labelControl2.TabIndex = 10;
            this.labelControl2.Text = "Department:";
            // 
            // PrintDeptsCheckedComboBoxEdit
            // 
            this.PrintDeptsCheckedComboBoxEdit.EditValue = "";
            this.PrintDeptsCheckedComboBoxEdit.Location = new System.Drawing.Point(437, 7);
            this.PrintDeptsCheckedComboBoxEdit.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.PrintDeptsCheckedComboBoxEdit.Name = "PrintDeptsCheckedComboBoxEdit";
            this.PrintDeptsCheckedComboBoxEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.PrintDeptsCheckedComboBoxEdit.Size = new System.Drawing.Size(127, 20);
            this.PrintDeptsCheckedComboBoxEdit.TabIndex = 9;
            // 
            // printTaskViewButton
            // 
            this.printTaskViewButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.printTaskViewButton.Location = new System.Drawing.Point(569, 6);
            this.printTaskViewButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.printTaskViewButton.Name = "printTaskViewButton";
            this.printTaskViewButton.Size = new System.Drawing.Size(69, 21);
            this.printTaskViewButton.TabIndex = 8;
            this.printTaskViewButton.Text = "Print";
            this.printTaskViewButton.Click += new System.EventHandler(this.printTaskViewButton_Click);
            // 
            // RefreshTasksButton
            // 
            this.RefreshTasksButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.RefreshTasksButton.Location = new System.Drawing.Point(362, 6);
            this.RefreshTasksButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.RefreshTasksButton.Name = "RefreshTasksButton";
            this.RefreshTasksButton.Size = new System.Drawing.Size(69, 21);
            this.RefreshTasksButton.TabIndex = 7;
            this.RefreshTasksButton.Text = "Refresh";
            this.RefreshTasksButton.Click += new System.EventHandler(this.RefreshTasksButton_Click);
            // 
            // departmentComboBox2
            // 
            this.departmentComboBox2.EditValue = "CNC Rough";
            this.departmentComboBox2.Location = new System.Drawing.Point(89, 7);
            this.departmentComboBox2.Margin = new System.Windows.Forms.Padding(4);
            this.departmentComboBox2.Name = "departmentComboBox2";
            this.departmentComboBox2.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.departmentComboBox2.Size = new System.Drawing.Size(127, 20);
            this.departmentComboBox2.TabIndex = 5;
            this.departmentComboBox2.SelectedIndexChanged += new System.EventHandler(this.departmentComboBox2_SelectedIndexChanged);
            // 
            // gridControl1
            // 
            this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridControl1.EmbeddedNavigator.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.gridControl1.Font = new System.Drawing.Font("Calibri", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridControl1.Location = new System.Drawing.Point(12, 36);
            this.gridControl1.MainView = this.gridView1;
            this.gridControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemHyperLinkEdit1,
            this.repositoryItemDateEdit4,
            this.repositoryItemDateEdit5,
            this.repositoryItemSpinEdit1,
            this.repositoryItemCheckedComboBoxEdit1,
            this.resourceRepositoryItemComboBox});
            this.gridControl1.Size = new System.Drawing.Size(1572, 750);
            this.gridControl1.TabIndex = 0;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            this.gridControl1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.gridControl1_MouseDown);
            // 
            // gridView1
            // 
            this.gridView1.Appearance.ColumnFilterButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.gridView1.Appearance.ColumnFilterButton.Options.UseBackColor = true;
            this.gridView1.Appearance.ColumnFilterButton.Options.UseFont = true;
            this.gridView1.Appearance.EvenRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.gridView1.Appearance.EvenRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.gridView1.Appearance.EvenRow.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.gridView1.Appearance.EvenRow.Options.UseBackColor = true;
            this.gridView1.Appearance.EvenRow.Options.UseBorderColor = true;
            this.gridView1.Appearance.EvenRow.Options.UseFont = true;
            this.gridView1.Appearance.FixedLine.Options.UseTextOptions = true;
            this.gridView1.Appearance.FixedLine.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.gridView1.Appearance.HeaderPanel.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.gridView1.Appearance.HeaderPanel.Options.UseFont = true;
            this.gridView1.Appearance.HeaderPanel.Options.UseTextOptions = true;
            this.gridView1.Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.gridView1.Appearance.OddRow.BackColor = System.Drawing.Color.White;
            this.gridView1.Appearance.OddRow.Options.UseBackColor = true;
            this.gridView1.Appearance.OddRow.Options.UseFont = true;
            this.gridView1.AppearancePrint.EvenRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.gridView1.AppearancePrint.EvenRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridView1.AppearancePrint.EvenRow.Options.UseBackColor = true;
            this.gridView1.AppearancePrint.EvenRow.Options.UseFont = true;
            this.gridView1.AppearancePrint.HeaderPanel.Font = new System.Drawing.Font("Trebuchet MS", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridView1.AppearancePrint.HeaderPanel.FontStyleDelta = System.Drawing.FontStyle.Bold;
            this.gridView1.AppearancePrint.HeaderPanel.Options.UseFont = true;
            this.gridView1.AppearancePrint.OddRow.BackColor = System.Drawing.Color.White;
            this.gridView1.AppearancePrint.OddRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridView1.AppearancePrint.OddRow.Options.UseBackColor = true;
            this.gridView1.AppearancePrint.OddRow.Options.UseFont = true;
            this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID3,
            this.colProjectStatus,
            this.colJobNumber,
            this.colProjectNumber,
            this.colComponent,
            this.colTaskID,
            this.colTaskName,
            this.colNotes2,
            this.colToolMaker,
            this.colHours1,
            this.colDuration,
            this.colStartDate,
            this.colFinishDate,
            this.colPredecessors,
            this.colDueDate,
            this.colMachine1,
            this.colPersonnel,
            this.colStatus});
            this.gridView1.GridControl = this.gridControl1;
            this.gridView1.GroupSummary.AddRange(new DevExpress.XtraGrid.GridSummaryItem[] {
            new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Hours", this.colHours1, "", "Total Hours:")});
            this.gridView1.Name = "gridView1";
            this.gridView1.OptionsBehavior.EditingMode = DevExpress.XtraGrid.Views.Grid.GridEditingMode.Inplace;
            this.gridView1.OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.MouseDown;
            this.gridView1.OptionsPrint.EnableAppearanceEvenRow = true;
            this.gridView1.OptionsPrint.EnableAppearanceOddRow = true;
            this.gridView1.OptionsPrint.PrintFooter = false;
            this.gridView1.OptionsPrint.PrintGroupFooter = false;
            this.gridView1.OptionsView.ColumnAutoWidth = false;
            this.gridView1.OptionsView.EnableAppearanceEvenRow = true;
            this.gridView1.OptionsView.EnableAppearanceOddRow = true;
            this.gridView1.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.colStartDate, DevExpress.Data.ColumnSortOrder.Ascending)});
            this.gridView1.ShownEditor += new System.EventHandler(this.gridView1_ShownEditor);
            this.gridView1.CellValueChanged += new DevExpress.XtraGrid.Views.Base.CellValueChangedEventHandler(this.gridView1_CellValueChanged);
            this.gridView1.CustomUnboundColumnData += new DevExpress.XtraGrid.Views.Base.CustomColumnDataEventHandler(this.gridView1_CustomUnboundColumnData);
            this.gridView1.PrintInitialize += new DevExpress.XtraGrid.Views.Base.PrintInitializeEventHandler(this.gridView1_PrintInitialize);
            this.gridView1.Click += new System.EventHandler(this.gridView1_Click);
            // 
            // colID3
            // 
            this.colID3.FieldName = "ID";
            this.colID3.Name = "colID3";
            // 
            // colProjectStatus
            // 
            this.colProjectStatus.Caption = "Project Status";
            this.colProjectStatus.FieldName = "ProjectStatus";
            this.colProjectStatus.Name = "colProjectStatus";
            this.colProjectStatus.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colProjectStatus.UnboundType = DevExpress.Data.UnboundColumnType.String;
            // 
            // colJobNumber
            // 
            this.colJobNumber.Caption = "Job #";
            this.colJobNumber.FieldName = "JobNumber";
            this.colJobNumber.Name = "colJobNumber";
            this.colJobNumber.OptionsColumn.AllowEdit = false;
            this.colJobNumber.Visible = true;
            this.colJobNumber.VisibleIndex = 0;
            this.colJobNumber.Width = 69;
            // 
            // colProjectNumber
            // 
            this.colProjectNumber.AppearanceCell.Options.UseTextOptions = true;
            this.colProjectNumber.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colProjectNumber.Caption = "Project #";
            this.colProjectNumber.FieldName = "ProjectNumber";
            this.colProjectNumber.Name = "colProjectNumber";
            this.colProjectNumber.OptionsColumn.AllowEdit = false;
            this.colProjectNumber.Visible = true;
            this.colProjectNumber.VisibleIndex = 1;
            this.colProjectNumber.Width = 69;
            // 
            // colComponent
            // 
            this.colComponent.ColumnEdit = this.repositoryItemHyperLinkEdit1;
            this.colComponent.FieldName = "Component";
            this.colComponent.Name = "colComponent";
            this.colComponent.OptionsColumn.AllowEdit = false;
            this.colComponent.Visible = true;
            this.colComponent.VisibleIndex = 2;
            this.colComponent.Width = 199;
            // 
            // repositoryItemHyperLinkEdit1
            // 
            this.repositoryItemHyperLinkEdit1.AutoHeight = false;
            this.repositoryItemHyperLinkEdit1.Name = "repositoryItemHyperLinkEdit1";
            this.repositoryItemHyperLinkEdit1.SingleClick = true;
            // 
            // colTaskID
            // 
            this.colTaskID.FieldName = "TaskID";
            this.colTaskID.Name = "colTaskID";
            // 
            // colTaskName
            // 
            this.colTaskName.FieldName = "TaskName";
            this.colTaskName.Name = "colTaskName";
            this.colTaskName.OptionsColumn.AllowEdit = false;
            this.colTaskName.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colTaskName.Visible = true;
            this.colTaskName.VisibleIndex = 4;
            this.colTaskName.Width = 109;
            // 
            // colNotes2
            // 
            this.colNotes2.FieldName = "Notes";
            this.colNotes2.Name = "colNotes2";
            this.colNotes2.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.True;
            this.colNotes2.Visible = true;
            this.colNotes2.VisibleIndex = 5;
            this.colNotes2.Width = 205;
            // 
            // colToolMaker
            // 
            this.colToolMaker.Caption = "Tool Maker";
            this.colToolMaker.FieldName = "ToolMaker2";
            this.colToolMaker.Name = "colToolMaker";
            this.colToolMaker.OptionsColumn.AllowEdit = false;
            this.colToolMaker.UnboundType = DevExpress.Data.UnboundColumnType.String;
            this.colToolMaker.Visible = true;
            this.colToolMaker.VisibleIndex = 3;
            this.colToolMaker.Width = 115;
            // 
            // colHours1
            // 
            this.colHours1.AppearanceCell.Options.UseTextOptions = true;
            this.colHours1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colHours1.ColumnEdit = this.repositoryItemSpinEdit1;
            this.colHours1.FieldName = "Hours";
            this.colHours1.Name = "colHours1";
            this.colHours1.Visible = true;
            this.colHours1.VisibleIndex = 6;
            this.colHours1.Width = 51;
            // 
            // repositoryItemSpinEdit1
            // 
            this.repositoryItemSpinEdit1.AutoHeight = false;
            this.repositoryItemSpinEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemSpinEdit1.Name = "repositoryItemSpinEdit1";
            // 
            // colDuration
            // 
            this.colDuration.FieldName = "Duration";
            this.colDuration.Name = "colDuration";
            this.colDuration.Visible = true;
            this.colDuration.VisibleIndex = 7;
            this.colDuration.Width = 65;
            // 
            // colStartDate
            // 
            this.colStartDate.AppearanceCell.Options.UseTextOptions = true;
            this.colStartDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colStartDate.DisplayFormat.FormatString = "MM/dd/yy";
            this.colStartDate.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.colStartDate.FieldName = "StartDate";
            this.colStartDate.Name = "colStartDate";
            this.colStartDate.Visible = true;
            this.colStartDate.VisibleIndex = 8;
            this.colStartDate.Width = 65;
            // 
            // colFinishDate
            // 
            this.colFinishDate.AppearanceCell.Options.UseTextOptions = true;
            this.colFinishDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colFinishDate.DisplayFormat.FormatString = "MM/dd/yy";
            this.colFinishDate.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.colFinishDate.FieldName = "FinishDate";
            this.colFinishDate.Name = "colFinishDate";
            this.colFinishDate.Visible = true;
            this.colFinishDate.VisibleIndex = 9;
            this.colFinishDate.Width = 65;
            // 
            // colPredecessors
            // 
            this.colPredecessors.FieldName = "Predecessors";
            this.colPredecessors.Name = "colPredecessors";
            this.colPredecessors.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            // 
            // colDueDate
            // 
            this.colDueDate.AppearanceCell.Options.UseTextOptions = true;
            this.colDueDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colDueDate.Caption = "Due Date";
            this.colDueDate.DisplayFormat.FormatString = "MM/dd/yy";
            this.colDueDate.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.colDueDate.FieldName = "DueDate";
            this.colDueDate.Name = "colDueDate";
            this.colDueDate.OptionsColumn.AllowEdit = false;
            this.colDueDate.UnboundType = DevExpress.Data.UnboundColumnType.DateTime;
            this.colDueDate.Visible = true;
            this.colDueDate.VisibleIndex = 10;
            this.colDueDate.Width = 65;
            // 
            // colMachine1
            // 
            this.colMachine1.ColumnEdit = this.repositoryItemCheckedComboBoxEdit1;
            this.colMachine1.FieldName = "Machine";
            this.colMachine1.Name = "colMachine1";
            this.colMachine1.Visible = true;
            this.colMachine1.VisibleIndex = 11;
            this.colMachine1.Width = 104;
            // 
            // repositoryItemCheckedComboBoxEdit1
            // 
            this.repositoryItemCheckedComboBoxEdit1.AutoHeight = false;
            this.repositoryItemCheckedComboBoxEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemCheckedComboBoxEdit1.Items.AddRange(new DevExpress.XtraEditors.Controls.CheckedListBoxItem[] {
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("Mazak", "Mazak"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("Mazak 1", "Mazak 1"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("Mazak 2", "Mazak 2"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("Mazak 3", "Mazak 3"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("Haas", "Haas"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("Old 640", "Old 640"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("New 640", "New 640"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("430", "430"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("950", "950"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("Sodick Mill", "Sodick Mill"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem("Makino", "Makino")});
            this.repositoryItemCheckedComboBoxEdit1.Name = "repositoryItemCheckedComboBoxEdit1";
            this.repositoryItemCheckedComboBoxEdit1.SeparatorChar = '/';
            this.repositoryItemCheckedComboBoxEdit1.QueryPopUp += new System.ComponentModel.CancelEventHandler(this.repositoryItemCheckedComboBoxEdit1_QueryPopUp);
            // 
            // colPersonnel
            // 
            this.colPersonnel.ColumnEdit = this.resourceRepositoryItemComboBox;
            this.colPersonnel.FieldName = "Personnel";
            this.colPersonnel.Name = "colPersonnel";
            this.colPersonnel.Visible = true;
            this.colPersonnel.VisibleIndex = 12;
            this.colPersonnel.Width = 117;
            // 
            // resourceRepositoryItemComboBox
            // 
            this.resourceRepositoryItemComboBox.AutoHeight = false;
            this.resourceRepositoryItemComboBox.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.resourceRepositoryItemComboBox.Name = "resourceRepositoryItemComboBox";
            this.resourceRepositoryItemComboBox.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            // 
            // colStatus
            // 
            this.colStatus.FieldName = "Status";
            this.colStatus.Name = "colStatus";
            this.colStatus.OptionsColumn.AllowEdit = false;
            this.colStatus.OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            this.colStatus.Visible = true;
            this.colStatus.VisibleIndex = 13;
            // 
            // repositoryItemDateEdit4
            // 
            this.repositoryItemDateEdit4.AutoHeight = false;
            this.repositoryItemDateEdit4.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemDateEdit4.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemDateEdit4.CalendarTimeProperties.EditFormat.FormatString = "d";
            this.repositoryItemDateEdit4.CalendarTimeProperties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.repositoryItemDateEdit4.Name = "repositoryItemDateEdit4";
            // 
            // repositoryItemDateEdit5
            // 
            this.repositoryItemDateEdit5.AutoHeight = false;
            this.repositoryItemDateEdit5.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemDateEdit5.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemDateEdit5.Name = "repositoryItemDateEdit5";
            // 
            // xtraTabPage7
            // 
            this.xtraTabPage7.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.xtraTabPage7.Controls.Add(this.restoreProjectButton);
            this.xtraTabPage7.Controls.Add(this.workLoadViewPrintPreviewButton);
            this.xtraTabPage7.Controls.Add(this.workLoadViewPrint2Button);
            this.xtraTabPage7.Controls.Add(this.workLoadViewPrintButton);
            this.xtraTabPage7.Controls.Add(this.changeViewRadioGroup);
            this.xtraTabPage7.Controls.Add(this.refreshLabelControl);
            this.xtraTabPage7.Controls.Add(this.resourceButton);
            this.xtraTabPage7.Controls.Add(this.editProjectButton);
            this.xtraTabPage7.Controls.Add(this.createProjectButton);
            this.xtraTabPage7.Controls.Add(this.backDateButton);
            this.xtraTabPage7.Controls.Add(this.forwardDateButton);
            this.xtraTabPage7.Controls.Add(this.kanBanButton);
            this.xtraTabPage7.Controls.Add(this.copyButton);
            this.xtraTabPage7.Controls.Add(this.RefreshProjectsButton);
            this.xtraTabPage7.Controls.Add(this.gridControl3);
            this.xtraTabPage7.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.xtraTabPage7.Name = "xtraTabPage7";
            this.xtraTabPage7.Size = new System.Drawing.Size(1596, 799);
            this.xtraTabPage7.Text = "Project View";
            // 
            // restoreProjectButton
            // 
            this.restoreProjectButton.AppearanceHovered.BackColor = System.Drawing.Color.DimGray;
            this.restoreProjectButton.AppearanceHovered.ForeColor = System.Drawing.Color.White;
            this.restoreProjectButton.AppearanceHovered.Options.UseBackColor = true;
            this.restoreProjectButton.AppearanceHovered.Options.UseForeColor = true;
            this.restoreProjectButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.restoreProjectButton.Location = new System.Drawing.Point(646, 7);
            this.restoreProjectButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.restoreProjectButton.Name = "restoreProjectButton";
            this.restoreProjectButton.Size = new System.Drawing.Size(85, 21);
            this.restoreProjectButton.TabIndex = 20;
            this.restoreProjectButton.Text = "Restore Project";
            this.restoreProjectButton.ToolTip = "Creates a new project.";
            this.restoreProjectButton.ToolTipAnchor = DevExpress.Utils.ToolTipAnchor.Cursor;
            this.restoreProjectButton.Click += new System.EventHandler(this.restoreProjectButton_Click);
            // 
            // workLoadViewPrintPreviewButton
            // 
            this.workLoadViewPrintPreviewButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.workLoadViewPrintPreviewButton.Location = new System.Drawing.Point(1241, 8);
            this.workLoadViewPrintPreviewButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.workLoadViewPrintPreviewButton.Name = "workLoadViewPrintPreviewButton";
            this.workLoadViewPrintPreviewButton.Size = new System.Drawing.Size(75, 23);
            this.workLoadViewPrintPreviewButton.TabIndex = 19;
            this.workLoadViewPrintPreviewButton.Text = "Print Preview";
            this.workLoadViewPrintPreviewButton.ToolTip = "Prints the grid onto an 8.5 x 11 sheet of paper.";
            this.workLoadViewPrintPreviewButton.Click += new System.EventHandler(this.workLoadViewPrintPreviewButton_Click);
            // 
            // workLoadViewPrint2Button
            // 
            this.workLoadViewPrint2Button.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.workLoadViewPrint2Button.Location = new System.Drawing.Point(1171, 8);
            this.workLoadViewPrint2Button.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.workLoadViewPrint2Button.Name = "workLoadViewPrint2Button";
            this.workLoadViewPrint2Button.Size = new System.Drawing.Size(64, 23);
            this.workLoadViewPrint2Button.TabIndex = 18;
            this.workLoadViewPrint2Button.Text = "Print Part";
            this.workLoadViewPrint2Button.ToolTip = "Prints the grid onto an 8.5 x 11 sheet of paper.";
            this.workLoadViewPrint2Button.Click += new System.EventHandler(this.workLoadViewPrint2Button_Click);
            // 
            // workLoadViewPrintButton
            // 
            this.workLoadViewPrintButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.workLoadViewPrintButton.Location = new System.Drawing.Point(1101, 8);
            this.workLoadViewPrintButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.workLoadViewPrintButton.Name = "workLoadViewPrintButton";
            this.workLoadViewPrintButton.Size = new System.Drawing.Size(64, 23);
            this.workLoadViewPrintButton.TabIndex = 15;
            this.workLoadViewPrintButton.Text = "Print Full";
            this.workLoadViewPrintButton.ToolTip = "Prints the full grid onto an 11 x 17 sheet of paper.";
            this.workLoadViewPrintButton.Click += new System.EventHandler(this.workLoadViewPrintButton_Click);
            // 
            // changeViewRadioGroup
            // 
            this.changeViewRadioGroup.Location = new System.Drawing.Point(864, 4);
            this.changeViewRadioGroup.Name = "changeViewRadioGroup";
            this.changeViewRadioGroup.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
            new DevExpress.XtraEditors.Controls.RadioGroupItem(true, "Project"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem(false, "Workload")});
            this.changeViewRadioGroup.Size = new System.Drawing.Size(175, 27);
            this.changeViewRadioGroup.TabIndex = 14;
            this.changeViewRadioGroup.SelectedIndexChanged += new System.EventHandler(this.changeViewRadioGroup_SelectedIndexChanged);
            // 
            // refreshLabelControl
            // 
            this.refreshLabelControl.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.refreshLabelControl.Appearance.BackColor = System.Drawing.Color.Transparent;
            this.refreshLabelControl.Appearance.BackColor2 = System.Drawing.Color.Transparent;
            this.refreshLabelControl.Appearance.BorderColor = System.Drawing.Color.Black;
            this.refreshLabelControl.Appearance.Options.UseBackColor = true;
            this.refreshLabelControl.Appearance.Options.UseBorderColor = true;
            this.refreshLabelControl.Appearance.Options.UseTextOptions = true;
            this.refreshLabelControl.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.refreshLabelControl.Location = new System.Drawing.Point(1361, 13);
            this.refreshLabelControl.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
            this.refreshLabelControl.LookAndFeel.UseDefaultLookAndFeel = false;
            this.refreshLabelControl.Name = "refreshLabelControl";
            this.refreshLabelControl.Size = new System.Drawing.Size(65, 13);
            this.refreshLabelControl.TabIndex = 13;
            this.refreshLabelControl.Text = "Last Refresh:";
            // 
            // resourceButton
            // 
            this.resourceButton.AppearanceHovered.BackColor = System.Drawing.Color.DimGray;
            this.resourceButton.AppearanceHovered.ForeColor = System.Drawing.Color.White;
            this.resourceButton.AppearanceHovered.Options.UseBackColor = true;
            this.resourceButton.AppearanceHovered.Options.UseForeColor = true;
            this.resourceButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.resourceButton.Location = new System.Drawing.Point(761, 7);
            this.resourceButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.resourceButton.Name = "resourceButton";
            this.resourceButton.Size = new System.Drawing.Size(85, 21);
            this.resourceButton.TabIndex = 12;
            this.resourceButton.Text = "Resources";
            this.resourceButton.ToolTip = "Allows user to create and assign resources to departments.";
            this.resourceButton.ToolTipAnchor = DevExpress.Utils.ToolTipAnchor.Cursor;
            this.resourceButton.Click += new System.EventHandler(this.resourceButton_Click);
            // 
            // editProjectButton
            // 
            this.editProjectButton.AppearanceHovered.BackColor = System.Drawing.Color.DimGray;
            this.editProjectButton.AppearanceHovered.ForeColor = System.Drawing.Color.White;
            this.editProjectButton.AppearanceHovered.Options.UseBackColor = true;
            this.editProjectButton.AppearanceHovered.Options.UseForeColor = true;
            this.editProjectButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.editProjectButton.Location = new System.Drawing.Point(291, 7);
            this.editProjectButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.editProjectButton.Name = "editProjectButton";
            this.editProjectButton.Size = new System.Drawing.Size(85, 21);
            this.editProjectButton.TabIndex = 11;
            this.editProjectButton.Text = "Edit Project";
            this.editProjectButton.ToolTip = "Edits a selected project.";
            this.editProjectButton.ToolTipAnchor = DevExpress.Utils.ToolTipAnchor.Cursor;
            this.editProjectButton.Click += new System.EventHandler(this.editProjectButton_Click);
            // 
            // createProjectButton
            // 
            this.createProjectButton.AppearanceHovered.BackColor = System.Drawing.Color.DimGray;
            this.createProjectButton.AppearanceHovered.ForeColor = System.Drawing.Color.White;
            this.createProjectButton.AppearanceHovered.Options.UseBackColor = true;
            this.createProjectButton.AppearanceHovered.Options.UseForeColor = true;
            this.createProjectButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.createProjectButton.Location = new System.Drawing.Point(115, 7);
            this.createProjectButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.createProjectButton.Name = "createProjectButton";
            this.createProjectButton.Size = new System.Drawing.Size(85, 21);
            this.createProjectButton.TabIndex = 10;
            this.createProjectButton.Text = "Create Project";
            this.createProjectButton.ToolTip = "Creates a new project.";
            this.createProjectButton.ToolTipAnchor = DevExpress.Utils.ToolTipAnchor.Cursor;
            this.createProjectButton.Click += new System.EventHandler(this.createProjectButton_Click);
            // 
            // backDateButton
            // 
            this.backDateButton.AppearanceHovered.BackColor = System.Drawing.Color.DimGray;
            this.backDateButton.AppearanceHovered.ForeColor = System.Drawing.Color.White;
            this.backDateButton.AppearanceHovered.Options.UseBackColor = true;
            this.backDateButton.AppearanceHovered.Options.UseForeColor = true;
            this.backDateButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.backDateButton.Location = new System.Drawing.Point(467, 7);
            this.backDateButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.backDateButton.Name = "backDateButton";
            this.backDateButton.Size = new System.Drawing.Size(85, 21);
            this.backDateButton.TabIndex = 9;
            this.backDateButton.Text = "Back Date";
            this.backDateButton.ToolTip = "Back dates selected components.";
            this.backDateButton.ToolTipAnchor = DevExpress.Utils.ToolTipAnchor.Cursor;
            this.backDateButton.Click += new System.EventHandler(this.backDateButton_Click);
            // 
            // forwardDateButton
            // 
            this.forwardDateButton.AppearanceHovered.BackColor = System.Drawing.Color.DimGray;
            this.forwardDateButton.AppearanceHovered.ForeColor = System.Drawing.Color.White;
            this.forwardDateButton.AppearanceHovered.Options.UseBackColor = true;
            this.forwardDateButton.AppearanceHovered.Options.UseForeColor = true;
            this.forwardDateButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.forwardDateButton.Location = new System.Drawing.Point(379, 7);
            this.forwardDateButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.forwardDateButton.Name = "forwardDateButton";
            this.forwardDateButton.Size = new System.Drawing.Size(85, 21);
            this.forwardDateButton.TabIndex = 8;
            this.forwardDateButton.Text = "Forward Date";
            this.forwardDateButton.ToolTip = "Forward dates selected components.";
            this.forwardDateButton.ToolTipAnchor = DevExpress.Utils.ToolTipAnchor.Cursor;
            this.forwardDateButton.Click += new System.EventHandler(this.forwardDateButton_Click);
            // 
            // kanBanButton
            // 
            this.kanBanButton.AppearanceHovered.BackColor = System.Drawing.Color.DimGray;
            this.kanBanButton.AppearanceHovered.ForeColor = System.Drawing.Color.White;
            this.kanBanButton.AppearanceHovered.Options.UseBackColor = true;
            this.kanBanButton.AppearanceHovered.Options.UseForeColor = true;
            this.kanBanButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.kanBanButton.Location = new System.Drawing.Point(555, 7);
            this.kanBanButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.kanBanButton.Name = "kanBanButton";
            this.kanBanButton.Size = new System.Drawing.Size(85, 21);
            this.kanBanButton.TabIndex = 7;
            this.kanBanButton.Text = "Kan Ban";
            this.kanBanButton.ToolTip = "Creates or modifies Kan Ban for selected project.";
            this.kanBanButton.ToolTipAnchor = DevExpress.Utils.ToolTipAnchor.Cursor;
            this.kanBanButton.Click += new System.EventHandler(this.kanBanButton_Click);
            // 
            // copyButton
            // 
            this.copyButton.AppearanceHovered.BackColor = System.Drawing.Color.DimGray;
            this.copyButton.AppearanceHovered.ForeColor = System.Drawing.Color.White;
            this.copyButton.AppearanceHovered.Options.UseBackColor = true;
            this.copyButton.AppearanceHovered.Options.UseForeColor = true;
            this.copyButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.copyButton.Location = new System.Drawing.Point(203, 7);
            this.copyButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.copyButton.Name = "copyButton";
            this.copyButton.Size = new System.Drawing.Size(85, 21);
            this.copyButton.TabIndex = 6;
            this.copyButton.Text = "Copy Project";
            this.copyButton.ToolTip = "Copies a selected project.";
            this.copyButton.ToolTipAnchor = DevExpress.Utils.ToolTipAnchor.Cursor;
            this.copyButton.Click += new System.EventHandler(this.copyButton_Click);
            // 
            // RefreshProjectsButton
            // 
            this.RefreshProjectsButton.AppearanceHovered.BackColor = System.Drawing.Color.DimGray;
            this.RefreshProjectsButton.AppearanceHovered.ForeColor = System.Drawing.Color.White;
            this.RefreshProjectsButton.AppearanceHovered.Options.UseBackColor = true;
            this.RefreshProjectsButton.AppearanceHovered.Options.UseForeColor = true;
            this.RefreshProjectsButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.RefreshProjectsButton.Location = new System.Drawing.Point(17, 7);
            this.RefreshProjectsButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.RefreshProjectsButton.Name = "RefreshProjectsButton";
            this.RefreshProjectsButton.Size = new System.Drawing.Size(85, 21);
            this.RefreshProjectsButton.TabIndex = 5;
            this.RefreshProjectsButton.Text = "Refresh";
            this.RefreshProjectsButton.ToolTip = "Refreshes data grid to what\'s in the database.";
            this.RefreshProjectsButton.Click += new System.EventHandler(this.RefreshProjectsButton_Click);
            // 
            // xtraTabPage3
            // 
            this.xtraTabPage3.Controls.Add(this.chartRadioGroup);
            this.xtraTabPage3.Controls.Add(this.rangeControl2);
            this.xtraTabPage3.Controls.Add(this.labelControl5);
            this.xtraTabPage3.Controls.Add(this.labelControl4);
            this.xtraTabPage3.Controls.Add(this.rangeControl1);
            this.xtraTabPage3.Controls.Add(this.timeFrameComboBoxEdit);
            this.xtraTabPage3.Controls.Add(this.TimeUnitsComboBox);
            this.xtraTabPage3.Controls.Add(this.RefreshChartButton);
            this.xtraTabPage3.Controls.Add(this.chartControl1);
            this.xtraTabPage3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.xtraTabPage3.Name = "xtraTabPage3";
            this.xtraTabPage3.Size = new System.Drawing.Size(1596, 799);
            this.xtraTabPage3.Text = "Chart View";
            // 
            // chartRadioGroup
            // 
            this.chartRadioGroup.Location = new System.Drawing.Point(336, 8);
            this.chartRadioGroup.Name = "chartRadioGroup";
            this.chartRadioGroup.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
            new DevExpress.XtraEditors.Controls.RadioGroupItem(true, "Department"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem(false, "Personnel")});
            this.chartRadioGroup.Size = new System.Drawing.Size(215, 28);
            this.chartRadioGroup.TabIndex = 20;
            // 
            // labelControl5
            // 
            this.labelControl5.Appearance.Options.UseFont = true;
            this.labelControl5.Location = new System.Drawing.Point(11, 38);
            this.labelControl5.Name = "labelControl5";
            this.labelControl5.Size = new System.Drawing.Size(61, 13);
            this.labelControl5.TabIndex = 18;
            this.labelControl5.Text = "Time Frame:";
            // 
            // labelControl4
            // 
            this.labelControl4.Appearance.Options.UseFont = true;
            this.labelControl4.Location = new System.Drawing.Point(16, 12);
            this.labelControl4.Name = "labelControl4";
            this.labelControl4.Size = new System.Drawing.Size(57, 13);
            this.labelControl4.TabIndex = 17;
            this.labelControl4.Text = "Time Units:";
            // 
            // timeFrameComboBoxEdit
            // 
            this.timeFrameComboBoxEdit.Location = new System.Drawing.Point(81, 35);
            this.timeFrameComboBoxEdit.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.timeFrameComboBoxEdit.Name = "timeFrameComboBoxEdit";
            this.timeFrameComboBoxEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.timeFrameComboBoxEdit.Size = new System.Drawing.Size(149, 20);
            this.timeFrameComboBoxEdit.TabIndex = 10;
            this.timeFrameComboBoxEdit.SelectedIndexChanged += new System.EventHandler(this.timeFrameComboBoxEdit_SelectedIndexChanged);
            // 
            // TimeUnitsComboBox
            // 
            this.TimeUnitsComboBox.EditValue = "Days";
            this.TimeUnitsComboBox.Location = new System.Drawing.Point(81, 9);
            this.TimeUnitsComboBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.TimeUnitsComboBox.Name = "TimeUnitsComboBox";
            this.TimeUnitsComboBox.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.TimeUnitsComboBox.Properties.Items.AddRange(new object[] {
            "Weeks",
            "Days"});
            this.TimeUnitsComboBox.Size = new System.Drawing.Size(149, 20);
            this.TimeUnitsComboBox.TabIndex = 7;
            this.TimeUnitsComboBox.SelectedIndexChanged += new System.EventHandler(this.TimeUnitsComboBox_SelectedIndexChanged);
            // 
            // RefreshChartButton
            // 
            this.RefreshChartButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.RefreshChartButton.Location = new System.Drawing.Point(249, 8);
            this.RefreshChartButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.RefreshChartButton.Name = "RefreshChartButton";
            this.RefreshChartButton.Size = new System.Drawing.Size(69, 21);
            this.RefreshChartButton.TabIndex = 5;
            this.RefreshChartButton.Text = "Refresh";
            this.RefreshChartButton.Click += new System.EventHandler(this.RefreshChartButton_Click);
            // 
            // xtraTabPage4
            // 
            this.xtraTabPage4.Controls.Add(this.labelControl6);
            this.xtraTabPage4.Controls.Add(this.panel1);
            this.xtraTabPage4.Controls.Add(this.projectComboBox);
            this.xtraTabPage4.Controls.Add(this.RefreshGanttButton);
            this.xtraTabPage4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.xtraTabPage4.Name = "xtraTabPage4";
            this.xtraTabPage4.Size = new System.Drawing.Size(1596, 799);
            this.xtraTabPage4.Text = "Gantt View";
            // 
            // labelControl6
            // 
            this.labelControl6.Appearance.Options.UseFont = true;
            this.labelControl6.Location = new System.Drawing.Point(14, 12);
            this.labelControl6.Name = "labelControl6";
            this.labelControl6.Size = new System.Drawing.Size(38, 13);
            this.labelControl6.TabIndex = 8;
            this.labelControl6.Text = "Project:";
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.splitContainerControl1);
            this.panel1.Location = new System.Drawing.Point(14, 36);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1569, 749);
            this.panel1.TabIndex = 6;
            // 
            // splitContainerControl1
            // 
            this.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerControl1.Location = new System.Drawing.Point(0, 0);
            this.splitContainerControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.splitContainerControl1.Name = "splitContainerControl1";
            this.splitContainerControl1.Panel1.Controls.Add(this.resourcesTree1);
            this.splitContainerControl1.Panel1.Text = "Panel1";
            this.splitContainerControl1.Panel2.Controls.Add(this.schedulerControl2);
            this.splitContainerControl1.Panel2.Text = "Panel2";
            this.splitContainerControl1.Size = new System.Drawing.Size(1569, 749);
            this.splitContainerControl1.SplitterPosition = 241;
            this.splitContainerControl1.TabIndex = 0;
            this.splitContainerControl1.Text = "splitContainerControl1";
            // 
            // resourcesTree1
            // 
            this.resourcesTree1.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
            this.colCaption});
            this.resourcesTree1.Cursor = System.Windows.Forms.Cursors.Default;
            this.resourcesTree1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.resourcesTree1.FixedLineWidth = 1;
            this.resourcesTree1.Location = new System.Drawing.Point(0, 0);
            this.resourcesTree1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.resourcesTree1.MinWidth = 21;
            this.resourcesTree1.Name = "resourcesTree1";
            this.resourcesTree1.OptionsBehavior.Editable = false;
            this.resourcesTree1.SchedulerControl = this.schedulerControl2;
            this.resourcesTree1.Size = new System.Drawing.Size(241, 749);
            this.resourcesTree1.TabIndex = 5;
            this.resourcesTree1.TreeLevelWidth = 17;
            this.resourcesTree1.CustomDrawNodeCell += new DevExpress.XtraTreeList.CustomDrawNodeCellEventHandler(this.resourcesTree1_CustomDrawNodeCell);
            // 
            // colCaption
            // 
            this.colCaption.AppearanceHeader.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.colCaption.AppearanceHeader.Options.UseFont = true;
            this.colCaption.AppearanceHeader.Options.UseTextOptions = true;
            this.colCaption.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colCaption.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colCaption.Caption = "Tasks";
            this.colCaption.FieldName = "TaskName";
            this.colCaption.MinWidth = 21;
            this.colCaption.Name = "colCaption";
            this.colCaption.Visible = true;
            this.colCaption.VisibleIndex = 0;
            // 
            // schedulerControl2
            // 
            this.schedulerControl2.ActiveViewType = DevExpress.XtraScheduler.SchedulerViewType.Gantt;
            this.schedulerControl2.DataStorage = this.schedulerStorage2;
            this.schedulerControl2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.schedulerControl2.GroupType = DevExpress.XtraScheduler.SchedulerGroupType.Resource;
            this.schedulerControl2.Location = new System.Drawing.Point(0, 0);
            this.schedulerControl2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.schedulerControl2.Name = "schedulerControl2";
            this.schedulerControl2.Size = new System.Drawing.Size(1323, 749);
            this.schedulerControl2.Start = new System.DateTime(2017, 12, 18, 0, 0, 0, 0);
            this.schedulerControl2.TabIndex = 0;
            this.schedulerControl2.Text = "schedulerControl2";
            this.schedulerControl2.Views.DayView.TimeRulers.Add(timeRuler4);
            this.schedulerControl2.Views.FullWeekView.Enabled = true;
            this.schedulerControl2.Views.FullWeekView.TimeRulers.Add(timeRuler5);
            this.schedulerControl2.Views.GanttView.ShowResourceHeaders = false;
            this.schedulerControl2.Views.TimelineView.ShowResourceHeaders = false;
            this.schedulerControl2.Views.WeekView.Enabled = false;
            this.schedulerControl2.Views.WorkWeekView.TimeRulers.Add(timeRuler6);
            this.schedulerControl2.AppointmentFlyoutShowing += new DevExpress.XtraScheduler.AppointmentFlyoutShowingEventHandler(this.schedulerControl2_AppointmentFlyoutShowing);
            // 
            // schedulerStorage2
            // 
            this.schedulerStorage2.AppointmentChanging += new DevExpress.XtraScheduler.PersistentObjectCancelEventHandler(this.schedulerStorage2_AppointmentChanging);
            this.schedulerStorage2.AppointmentsChanged += new DevExpress.XtraScheduler.PersistentObjectsEventHandler(this.schedulerStorage2_AppointmentsChanged);
            // 
            // projectComboBox
            // 
            this.projectComboBox.Location = new System.Drawing.Point(61, 8);
            this.projectComboBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.projectComboBox.Name = "projectComboBox";
            this.projectComboBox.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.projectComboBox.Size = new System.Drawing.Size(149, 20);
            this.projectComboBox.TabIndex = 4;
            this.projectComboBox.SelectedIndexChanged += new System.EventHandler(this.projectComboBox_SelectedIndexChanged);
            // 
            // RefreshGanttButton
            // 
            this.RefreshGanttButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat;
            this.RefreshGanttButton.Location = new System.Drawing.Point(221, 8);
            this.RefreshGanttButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.RefreshGanttButton.Name = "RefreshGanttButton";
            this.RefreshGanttButton.Size = new System.Drawing.Size(69, 21);
            this.RefreshGanttButton.TabIndex = 2;
            this.RefreshGanttButton.Text = "Refresh";
            this.RefreshGanttButton.Click += new System.EventHandler(this.RefreshGanttButton_Click);
            // 
            // tasksBindingSource
            // 
            this.tasksBindingSource.DataMember = "Tasks";
            this.tasksBindingSource.DataSource = this.workload_Tracking_System_DBDataSet;
            // 
            // resourcesBindingSource
            // 
            this.resourcesBindingSource.DataMember = "Resources";
            this.resourcesBindingSource.DataSource = this.workload_Tracking_System_DBDataSet;
            // 
            // resourcesTableAdapter
            // 
            this.resourcesTableAdapter.ClearBeforeFill = true;
            // 
            // componentsBindingSource
            // 
            this.componentsBindingSource.DataMember = "Components";
            this.componentsBindingSource.DataSource = this.workload_Tracking_System_DBDataSet;
            // 
            // componentsTableAdapter
            // 
            this.componentsTableAdapter.ClearBeforeFill = true;
            // 
            // MainWindow
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1628, 852);
            this.Controls.Add(this.xtraTabControl1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.Shadow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "MainWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Toolroom Scheduler";
            this.Load += new System.EventHandler(this.MainWindow_Load);
            ((System.ComponentModel.ISupportInitialize)(this.rangeControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(xyDiagram1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(series1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(sideBySideBarSeriesLabel1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chartControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.workload_Tracking_System_DBDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rangeControl2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemComboBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.projectsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.personnelComboBoxEdit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemHyperLinkEdit2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemImageEdit2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemPictureEdit1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemImageComboBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.stageComboBoxEdit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.genericDateEdit.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.genericDateEdit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.projectBandedGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DeptProgressGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerStorage1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).EndInit();
            this.xtraTabControl1.ResumeLayout(false);
            this.xtraTabPage1.ResumeLayout(false);
            this.xtraTabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.includeCompletesCheckEdit.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.includeQuotesCheckEdit.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.projectCheckedComboBoxEdit.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GroupByRadioGroup.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.departmentComboBox.Properties)).EndInit();
            this.xtraTabPage2.ResumeLayout(false);
            this.xtraTabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PrintEmployeeWorkCheckedComboBoxEdit.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.daysAheadSpinEdit.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.filterTasksByDatesCheckEdit.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PrintDeptsCheckedComboBoxEdit.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.departmentComboBox2.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemHyperLinkEdit1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckedComboBoxEdit1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.resourceRepositoryItemComboBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit4.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit5.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit5)).EndInit();
            this.xtraTabPage7.ResumeLayout(false);
            this.xtraTabPage7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.changeViewRadioGroup.Properties)).EndInit();
            this.xtraTabPage3.ResumeLayout(false);
            this.xtraTabPage3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartRadioGroup.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.timeFrameComboBoxEdit.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TimeUnitsComboBox.Properties)).EndInit();
            this.xtraTabPage4.ResumeLayout(false);
            this.xtraTabPage4.PerformLayout();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).EndInit();
            this.splitContainerControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.resourcesTree1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerControl2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerStorage2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.projectComboBox.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tasksBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.resourcesBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.behaviorManager1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.componentsBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraScheduler.SchedulerStorage schedulerStorage1;
        private DevExpress.XtraTab.XtraTabControl xtraTabControl1;
        private DevExpress.XtraTab.XtraTabPage xtraTabPage1;
        private DevExpress.XtraScheduler.SchedulerControl schedulerControl1;
        private DevExpress.XtraTab.XtraTabPage xtraTabPage2;
        private DevExpress.XtraEditors.ComboBoxEdit departmentComboBox;
        private DevExpress.LookAndFeel.DefaultLookAndFeel defaultLookAndFeel1;
        private DevExpress.XtraEditors.SimpleButton refreshButton;
        private DevExpress.XtraTab.XtraTabPage xtraTabPage3;
        private DevExpress.XtraTab.XtraTabPage xtraTabPage4;
        private Workload_Tracking_System_DBDataSet workload_Tracking_System_DBDataSet;
        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private DevExpress.XtraEditors.SimpleButton RefreshTasksButton;
        private DevExpress.XtraEditors.ComboBoxEdit departmentComboBox2;
        private System.Windows.Forms.BindingSource resourcesBindingSource;
        private Workload_Tracking_System_DBDataSetTableAdapters.ResourcesTableAdapter resourcesTableAdapter;
        private DevExpress.XtraScheduler.SchedulerControl schedulerControl2;
        private DevExpress.XtraEditors.SimpleButton RefreshGanttButton;
        private DevExpress.XtraEditors.ComboBoxEdit projectComboBox;
        private DevExpress.XtraScheduler.SchedulerStorage schedulerStorage2;
        private DevExpress.XtraScheduler.UI.ResourcesTree resourcesTree1;
        private System.Windows.Forms.Panel panel1;
        private DevExpress.XtraEditors.SplitContainerControl splitContainerControl1;
        private DevExpress.XtraScheduler.Native.ResourceTreeColumn colCaption;
        private DevExpress.XtraEditors.Repository.RepositoryItemHyperLinkEdit repositoryItemHyperLinkEdit1;
        private DevExpress.XtraEditors.SimpleButton RefreshChartButton;
        private DevExpress.XtraCharts.ChartControl chartControl1;
        private DevExpress.XtraEditors.ComboBoxEdit TimeUnitsComboBox;
        private DevExpress.XtraEditors.SimpleButton printTaskViewButton;
        private DevExpress.XtraEditors.ComboBoxEdit timeFrameComboBoxEdit;
        private DevExpress.XtraEditors.RangeControl rangeControl1;
        private DevExpress.Utils.Behaviors.BehaviorManager behaviorManager1;
        private Workload_Tracking_System_DBDataSetTableAdapters.TasksTableAdapter tasksTableAdapter;
        private DevExpress.XtraTab.XtraTabPage xtraTabPage7;
        private DevExpress.XtraGrid.GridControl gridControl3;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView3;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView5;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView4;
        private Workload_Tracking_System_DBDataSetTableAdapters.ProjectsTableAdapter projectsTableAdapter;
        private System.Windows.Forms.BindingSource projectsBindingSource;
        private DevExpress.XtraGrid.Columns.GridColumn colJobNumber1;
        private DevExpress.XtraGrid.Columns.GridColumn colProjectNumber2;
        private DevExpress.XtraGrid.Columns.GridColumn colDueDate1;
        private DevExpress.XtraGrid.Columns.GridColumn colPriority;
        private DevExpress.XtraGrid.Columns.GridColumn colStatus1;
        private DevExpress.XtraGrid.Columns.GridColumn colDesigner1;
        private DevExpress.XtraGrid.Columns.GridColumn colToolMaker2;
        private DevExpress.XtraGrid.Columns.GridColumn colRoughProgrammer1;
        private DevExpress.XtraGrid.Columns.GridColumn colElectrodeProgrammer1;
        private DevExpress.XtraGrid.Columns.GridColumn colFinishProgrammer1;
        private DevExpress.XtraGrid.Columns.GridColumn colEngineer1;
        private DevExpress.XtraGrid.Columns.GridColumn colKanBanWorkbookPath;
        private DevExpress.XtraGrid.Columns.GridColumn colComponent1;
        private DevExpress.XtraGrid.Columns.GridColumn colMaterial;
        private DevExpress.XtraGrid.Columns.GridColumn colNotes;
        private DevExpress.XtraGrid.Columns.GridColumn colPercentComplete;
        private DevExpress.XtraGrid.Columns.GridColumn colPictures;
        private DevExpress.XtraGrid.Columns.GridColumn colPosition;
        private DevExpress.XtraGrid.Columns.GridColumn colPriority1;
        private DevExpress.XtraGrid.Columns.GridColumn colQuantity;
        private DevExpress.XtraGrid.Columns.GridColumn colSpares;
        private DevExpress.XtraGrid.Columns.GridColumn colStatus2;
        private DevExpress.XtraGrid.Columns.GridColumn colTaskName1;
        private DevExpress.XtraGrid.Columns.GridColumn colStartDate2;
        private DevExpress.XtraGrid.Columns.GridColumn colFinishDate2;
        private DevExpress.XtraGrid.Columns.GridColumn colResource1;
        private DevExpress.XtraGrid.Columns.GridColumn colMachine;
        private DevExpress.XtraGrid.Columns.GridColumn colHours;
        private DevExpress.XtraGrid.Columns.GridColumn colStatus3;
        private System.Windows.Forms.BindingSource componentsBindingSource;
        private Workload_Tracking_System_DBDataSetTableAdapters.ComponentsTableAdapter componentsTableAdapter;
        private DevExpress.XtraGrid.Columns.GridColumn colInitials;
        private DevExpress.XtraGrid.Columns.GridColumn colDateCompleted;
        private DevExpress.XtraGrid.Columns.GridColumn colFinish;
        private DevExpress.XtraGrid.Columns.GridColumn colDuration1;
        private DevExpress.XtraGrid.Columns.GridColumn colPercentComplete1;
        private DevExpress.XtraEditors.SimpleButton RefreshProjectsButton;
        private DevExpress.XtraGrid.Columns.GridColumn colCustomer1;
        private DevExpress.XtraGrid.Columns.GridColumn colProject;
        private DevExpress.XtraEditors.Repository.RepositoryItemDateEdit repositoryItemDateEdit4;
        private DevExpress.XtraEditors.Repository.RepositoryItemDateEdit repositoryItemDateEdit5;
        private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit1;
        private System.Windows.Forms.BindingSource tasksBindingSource;
        private DevExpress.XtraGrid.Columns.GridColumn colJobNumber;
        private DevExpress.XtraGrid.Columns.GridColumn colProjectNumber;
        private DevExpress.XtraGrid.Columns.GridColumn colComponent;
        private DevExpress.XtraGrid.Columns.GridColumn colDuration;
        private DevExpress.XtraGrid.Columns.GridColumn colStartDate;
        private DevExpress.XtraGrid.Columns.GridColumn colFinishDate;
        private DevExpress.XtraGrid.Columns.GridColumn colPersonnel;
        private DevExpress.XtraGrid.Columns.GridColumn colHours1;
        private DevExpress.XtraGrid.Columns.GridColumn colStatus;
        private DevExpress.XtraGrid.Columns.GridColumn colDueDate;
        private DevExpress.XtraGrid.Columns.GridColumn colTaskName;
        private DevExpress.XtraGrid.Columns.GridColumn colToolMaker;
        private DevExpress.XtraGrid.Columns.GridColumn colProjectStatus;
        private DevExpress.XtraGrid.Columns.GridColumn colTaskID;
        private DevExpress.XtraGrid.Columns.GridColumn colPredecessors;
        private DevExpress.XtraGrid.Columns.GridColumn colTaskID1;
        private DevExpress.XtraGrid.Columns.GridColumn colPredecessors1;
        private DevExpress.XtraGrid.Columns.GridColumn colNotes1;
        private DevExpress.XtraGrid.Columns.GridColumn colID1;
        private DevExpress.XtraGrid.Columns.GridColumn colID2;
        private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit2;
        private DevExpress.XtraEditors.Repository.RepositoryItemComboBox repositoryItemComboBox3;
        private DevExpress.XtraGrid.Columns.GridColumn colNotes2;
        private DevExpress.XtraGrid.Columns.GridColumn colMachine1;
        private DevExpress.XtraGrid.Columns.GridColumn colID3;
        private DevExpress.XtraGrid.Columns.GridColumn colID4;
        private DevExpress.XtraEditors.Repository.RepositoryItemImageEdit repositoryItemImageEdit2;
        private DevExpress.XtraEditors.Repository.RepositoryItemPictureEdit repositoryItemPictureEdit1;
        private DevExpress.XtraEditors.Repository.RepositoryItemImageComboBox repositoryItemImageComboBox1;
        private DevExpress.XtraEditors.SimpleButton copyButton;
        private DevExpress.XtraEditors.SimpleButton kanBanButton;
        private DevExpress.XtraEditors.SimpleButton forwardDateButton;
        private DevExpress.XtraEditors.SimpleButton backDateButton;
        private DevExpress.XtraEditors.SimpleButton createProjectButton;
        private DevExpress.XtraEditors.SimpleButton editProjectButton;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckedComboBoxEdit repositoryItemCheckedComboBoxEdit1;
        private DevExpress.XtraEditors.SimpleButton resourceButton;
        private DevExpress.XtraEditors.CheckedComboBoxEdit PrintDeptsCheckedComboBoxEdit;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private DevExpress.XtraEditors.LabelControl labelControl5;
        private DevExpress.XtraEditors.LabelControl labelControl4;
        private DevExpress.XtraEditors.LabelControl labelControl6;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        private DevExpress.XtraEditors.RadioGroup GroupByRadioGroup;
        private DevExpress.XtraGrid.Columns.GridColumn colOverlapAllowed;
        private DevExpress.XtraEditors.Repository.RepositoryItemHyperLinkEdit repositoryItemHyperLinkEdit2;
        private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit2;
        private DevExpress.XtraGrid.Columns.GridColumn colIncludeHours;
        private DevExpress.XtraEditors.LabelControl labelControl7;
        private DevExpress.XtraEditors.CheckedComboBoxEdit projectCheckedComboBoxEdit;
        private DevExpress.XtraEditors.RangeControl rangeControl2;
        private DevExpress.XtraEditors.SimpleButton printEmployeeWorkButton;
        private DevExpress.XtraEditors.CheckEdit includeCompletesCheckEdit;
        private DevExpress.XtraEditors.CheckEdit includeQuotesCheckEdit;
        private DevExpress.XtraEditors.RadioGroup chartRadioGroup;
        private DevExpress.XtraEditors.Repository.RepositoryItemComboBox resourceRepositoryItemComboBox;
        private DevExpress.XtraGrid.Columns.GridColumn colApprentice;
        private DevExpress.XtraGrid.Columns.GridColumn colDateModified;
        private DevExpress.XtraGrid.Columns.GridColumn colLastKanBanGenerationDate;
        private DevExpress.XtraGrid.Columns.GridColumn colProjectNumber3;
        private DevExpress.XtraGrid.Columns.GridColumn colLatestFinishDate;
        private DevExpress.XtraEditors.LabelControl refreshLabelControl;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridView projectBandedGridView;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colJobNumberBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colProjectNumberBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colCustomerBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colProjectBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colMoldCostBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colStartDateBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colDueDateBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colAdjustedDeliveryDateBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colEngineerBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colDesignerBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colToolMakerBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colRoughProgrammerBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colElectrodeProgrammerBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colFinishProgrammerBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colApprenticeBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colManifoldBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colMoldBaseBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colGeneralNotesBGV;
        private DevExpress.XtraEditors.RadioGroup changeViewRadioGroup;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colStatusBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colStageBGV;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colDeliveryInWeeksBGV;
        private DevExpress.XtraEditors.Repository.RepositoryItemComboBox stageComboBoxEdit;
        private DevExpress.XtraEditors.Repository.RepositoryItemDateEdit genericDateEdit;
        private DevExpress.XtraEditors.Repository.RepositoryItemComboBox personnelComboBoxEdit;
        private DevExpress.XtraEditors.SimpleButton workLoadViewPrintPreviewButton;
        private DevExpress.XtraEditors.SimpleButton workLoadViewPrint2Button;
        private DevExpress.XtraEditors.SimpleButton workLoadViewPrintButton;
        private DevExpress.XtraGrid.Views.BandedGrid.GridBand SegoeUI;
        private DevExpress.XtraGrid.Views.BandedGrid.GridBand milestonesGridBand;
        private DevExpress.XtraGrid.Views.BandedGrid.GridBand personnelGridBand;
        private DevExpress.XtraGrid.Views.BandedGrid.GridBand generalInfoGridBand;
        private DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn colIDBGV;
        private DevExpress.XtraEditors.SimpleButton restoreProjectButton;
        private DevExpress.XtraGrid.Views.Grid.GridView DeptProgressGridView;
        private DevExpress.XtraGrid.Columns.GridColumn DepartmentColDPV;
        private DevExpress.XtraGrid.Columns.GridColumn PercentCompleteColDPV;
        private DevExpress.XtraGrid.Columns.GridColumn colStagePV;
        private DevExpress.XtraEditors.SpinEdit daysAheadSpinEdit;
        private DevExpress.XtraEditors.CheckEdit filterTasksByDatesCheckEdit;
        private DevExpress.XtraEditors.LabelControl labelControl8;
        private DevExpress.XtraEditors.CheckedComboBoxEdit PrintEmployeeWorkCheckedComboBoxEdit;
    }
}

