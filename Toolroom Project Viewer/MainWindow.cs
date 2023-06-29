using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.Xml;
using DevExpress.XtraGrid.Views.Base;
using System;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Base.ViewInfo;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraPrinting;
using DevExpress.XtraCharts;
using System.Text;

using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using ClassLibrary;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;
using System.Drawing;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using System.Reflection;
using System.IO;
using System.Threading.Tasks;
using Squirrel;
using System.Text.RegularExpressions;
using DevExpress.Data.Filtering;
using DevExpress.XtraSplashScreen;
using System.Threading;
using DevExpress.XtraReports.UI;
using DevExpress.XtraScheduler.Reporting;
using DevExpress.Spreadsheet;
using DevExpress.Utils.Menu;
using DevExpress.XtraScheduler.Drawing;
using ClassLibrary.Models;
using DevExpress.Utils;

namespace Toolroom_Project_Viewer
{
    public partial class MainWindow : DevExpress.XtraEditors.XtraForm
    {
        private string footerDateTime = "";
        private BindingList<CustomResource> CustomResourceCollection = new BindingList<CustomResource>();
        private BindingList<CustomAppointment> CustomEventList = new BindingList<CustomAppointment>();
        private BindingList<CustomDependency> CustomDependencyList = new BindingList<CustomDependency>();
        private List<ColorStruct> ColorList = new List<ColorStruct>();
        private RepositoryItemPopupContainerEdit repositoryItemPopupContainerEdit = new RepositoryItemPopupContainerEdit();
        private string PrintOrientation, PaperSize;
        private string[] DepartmentArr = { "Design", "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "CNC Electrodes", "Grind", "Inspection", "EDM Sinker", "EDM Wire (In-House)", "Polish" };

        object DraggedResourceId = null;
        public DataTable RoleTable { get; set; }
        private string TimeUnits { get; set; }
        private string NoResourceName { get; set; }
        private ProjectModel Project { get; set; }
        private ProjectModel CalendarProject { get; set; }
        public List<ProjectModel> ProjectsList { get; set; }
        public List<ComponentModel> ComponentsList { get; set; }
        public List<TaskModel> TasksList { get; set; }
        //private List<WorkLoadModel> WorkloadList { get; set; }
        public List<ProjectModel> DeletedProjects { get; set; } = new List<ProjectModel>();
        private List<UserModel> UserList { get; set; }

        private List<int> gridView3ExpandedRowsList = new List<int>();
        private List<int> gridView4ExpandedRowsList = new List<int>();
        private List<ExpandedProjectRows> expandedProjectRowsList = new List<ExpandedProjectRows>();
        private List<int> ProjectBandedGridViewExpandedGroupList = new List<int>();
        private List<int> gridView3SelectedRows = new List<int>();
        private List<int> gridView4SelectedRows = new List<int>();
        private List<int> gridView5SelectedRows = new List<int>();
        private DataTable ResourceDataTable;
        private string Role, Tasks;
        Regex TaskRegExpression, RoleRegExpression;
        private bool AllProjectItemsChecked, MoveSelectedAppointments, RightMouseButtonPressed, MoveSubsequentTaskWithLockedSpacing, AppointmentErrorSent, AppointmentResourceChanged, AppointmentDateChanged, EditAppointmentByForm;
        DateTime OldTaskStartDate;
        Appointment DraggedAppointment;
        private SchedulerHitInfo HitInfo { get; set; }

        private RefreshHelper helper1, helper2, deptTaskViewHelper;

        public MainWindow()
        {
            try
            {
                this.TimeUnits = "Days";
                ResourceDataTable = Database.GetResourceData();
                UserList = Database.GetUsers();
                InitializeComponent();
                SetRole();
                SetTasks();
                PopulateProjectCheckedComboBox();
                InitializeResources();

                GroupByRadioGroup.SelectedIndex = 0;
                changeViewRadioGroup.SelectedIndex = 0;
                chartRadioGroup.SelectedIndex = 0;

                PopulateDepartmentComboBoxes();
                PopulateProjectComboBox();
                PopulateProjectComboBox2();
                PopulateTimeFrameComboBox();               

                schedulerStorage2.Appointments.CommitIdToDataSource = false;               

                string forecastedHoursSheetPath = AppDomain.CurrentDomain.BaseDirectory + @"\Resources\Forecasted_Hours.xlsx";

                spreadsheetControl1.LoadDocument(forecastedHoursSheetPath);

                if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.EngineeringNumberVisible))
                {
                    projectBandedGridView.Columns["EngineeringProjectNumber"].Visible = false; 
                }
                else
                {
                    projectBandedGridView.Columns["EngineeringProjectNumber"].Visible = true;
                }
                //InitializeExample();
                AddRepositoryItemToGrid();

                gridView3.DetailHeight = int.MaxValue;
                gridView4.DetailHeight = int.MaxValue;

                AddVersionNumber();

                CheckForUpdates();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
                Console.WriteLine(e.ToString());
            }
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            try
            {
                helper1 = new RefreshHelper(gridView3, "JobNumber");
                helper2 = new RefreshHelper(gridView4, "Component");

                gridControl3.LevelTree.Nodes.Add("DeptProgresses", DeptProgressGridView);

                LoadProjects();
                LoadProjectView();
                LoadTaskView();

                //foreach (var component in ComponentsList)
                //{
                //    if (component.Tasks.Count() > 0)
                //    {
                //        component.AllTasksDated = component.CheckIfAllTasksDated();
                //        Database.UpdateComponent(component, "AllTasksDated");
                //    }
                //}

                InitializeAppointments();
                InitializeDepartmentPrintOptions();
                InitializePersonnelPrintOptions();

                footerDateTime = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
                gridView5.SortInfo.Add(new GridColumnSortInfo(colTaskID1, DevExpress.Data.ColumnSortOrder.Ascending));
                schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskID", "TaskID"));
                schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Component", "Component"));
                RoleTable = Database.GetRoleTable();

                //PopulateEmployeeComboBox();
                gridView1.ActiveFilterCriteria = FilterTaskView(departmentComboBox2.Text, false, false, filterTasksByDatesCheckEdit.Checked);
                gridView3.ActiveFilterCriteria = FilterTaskView3();
                //MessageBox.Show($"{schedulerControl1.TimelineView.GetBaseTimeScale().Width}"); 

                schedulerControl1.Start = DateTime.Today.AddDays(-7);
                schedulerControl1.OptionsCustomization.AllowAppointmentDelete = UsedAppointmentType.Custom;
                schedulerControl1.AllowAppointmentDelete += new AppointmentOperationEventHandler(schedulerControl1_AllowAppointmentDelete);
                schedulerControl1.GanttView.GetBaseTimeScale().Width = 60;
                schedulerControl1.GanttView.AppointmentDisplayOptions.StartTimeVisibility = AppointmentTimeVisibility.Never;
                schedulerControl1.GanttView.AppointmentDisplayOptions.EndTimeVisibility = AppointmentTimeVisibility.Never;
                //schedulerControl1.OptionsCustomization.AllowDisplayAppointmentForm = AllowDisplayAppointmentForm.Never;
                //schedulerControl1.OptionsCustomization.AllowInplaceEditor = UsedAppointmentType.None;
                //gridView3.Columns["IncludeHours"].VisibleIndex = 14;

                zoomTrackBarControl1.Properties.Maximum = 215;
                zoomTrackBarControl1.Properties.Minimum = 60;
                zoomTrackBarControl1.Value = schedulerControl1.GanttView.GetBaseTimeScale().Width;

                schedulerControl2.Start = DateTime.Today.AddDays(-7);
                schedulerControl2.Views.GanttView.ResourcesPerPage = 15;
                schedulerControl2.GroupType = SchedulerGroupType.Resource;
                schedulerControl2.ActiveViewType = SchedulerViewType.Gantt;
                schedulerControl2.OptionsCustomization.AllowDisplayAppointmentForm = AllowDisplayAppointmentForm.Never;
                schedulerControl2.OptionsCustomization.AllowInplaceEditor = UsedAppointmentType.None;

                schedulerControl3.GanttView.AppointmentDisplayOptions.StartTimeVisibility = AppointmentTimeVisibility.Never;
                schedulerControl3.GanttView.AppointmentDisplayOptions.EndTimeVisibility = AppointmentTimeVisibility.Never;
                schedulerControl3.TimelineView.AppointmentDisplayOptions.StartTimeVisibility = AppointmentTimeVisibility.Never;
                schedulerControl3.TimelineView.AppointmentDisplayOptions.EndTimeVisibility = AppointmentTimeVisibility.Never;
                schedulerControl3.OptionsCustomization.AllowDisplayAppointmentForm = AllowDisplayAppointmentForm.Never;
                schedulerControl3.OptionsCustomization.AllowInplaceEditor = UsedAppointmentType.None;
                schedulerControl3.MonthView.AppointmentDisplayOptions.StartTimeVisibility = AppointmentTimeVisibility.Never;
                schedulerControl3.MonthView.AppointmentDisplayOptions.EndTimeVisibility = AppointmentTimeVisibility.Never;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }

        private void AddVersionNumber()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            FileVersionInfo versionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);

            this.Text += $" v.{versionInfo.FileVersion}";
        }

        private async Task CheckForUpdates()
        {
            using (var manager = new UpdateManager(@"X:\TOOLROOM\Workload Tracking System\Releases"))
            {
                await manager.UpdateApp();
            }
        }
        private void LoadProjects()
        {
            var data = Database.GetProjects();

            ProjectsList = data.projects;
            ComponentsList = data.components;
            TasksList = data.tasks;            

            InitializeProjects();
        }
        private void LoadProjectView()
        {
            BindingList<ProjectModel> projects = new BindingList<ProjectModel>(ProjectsList);
            gridControl3.DataSource = projects;

            refreshLabelControl.Text = "Last Refresh: " + DateTime.Now.ToString("M/d/yyyy hh:mm:ss tt");
        }
        //private void LoadWorkloadView()
        //{
        //    WorkloadList = Database.GetWorkloads();
        //    BindingList<WorkLoadModel> workLoads = new BindingList<WorkLoadModel>(WorkloadList);
        //    gridControl2.DataSource = workLoads;
        //}
        private void InitializeProjects()
        {
            List<TaskModel> projectTasks;

            foreach (var project in ProjectsList)
            {
                project.Components = new List<ComponentModel>(ComponentsList.FindAll(x => x.ProjectNumber == project.ProjectNumber));

                projectTasks = new List<TaskModel>(TasksList.FindAll(x => x.ProjectNumber == project.ProjectNumber));

                foreach (var component in project.Components)
                {
                    component.Tasks = new List<TaskModel>(projectTasks.FindAll(x => x.Component == component.Component));
                }

                var result = from task in projectTasks
                             //where task.ProjectNumber == 37242
                             group task by task.TaskName into grp
                             select 
                             new 
                             { 
                                 Department = grp.Key, 
                                 PercentComplete = (double)grp.Where(x => x.Status == "Completed").Sum(x => x.Hours) / grp.Sum(x => x.Hours)
                             };

                foreach (var item in result.ToList())
                {
                    project.DeptProgresses.Add(new DeptProgress() { ProjectNumber = project.ProjectNumber, Department = item.Department, PercentComplete = item.PercentComplete }); 
                }
            }
        }
        private void LoadTaskView()
        {
            BindingList<TaskModel> tasks = new BindingList<TaskModel>(TasksList);

            gridControl1.DataSource = tasks;
        }

        #region Department Schedule View

        private void InitializeResources()
        {
            Database db = new Database();
            DataTable dt = new DataTable();

            schedulerDataStorage1.Resources.Clear();
            schedulerDataStorage1.Resources.CustomFieldMappings.Clear();

            ResourceStorage resourceStorage = new ResourceStorage(schedulerDataStorage1);
            ResourceMappingInfo resourceMappings = schedulerDataStorage1.Resources.Mappings;

            resourceMappings.Caption = "ResourceName";
            resourceMappings.Id = "ResourceName";

            schedulerDataStorage1.Resources.CustomFieldMappings.Add(new ResourceCustomFieldMapping("ResourceType", "ResourceType"));
            schedulerDataStorage1.Resources.CustomFieldMappings.Add(new ResourceCustomFieldMapping("Role", "Role"));
            schedulerDataStorage1.Resources.CustomFieldMappings.Add(new ResourceCustomFieldMapping("Department", "Department"));

            schedulerDataStorage1.Resources.DataSource = ResourceDataTable.AsEnumerable().GroupBy(x => x.Field<string>("ResourceName")).Select(x => x.FirstOrDefault()).CopyToDataTable();

            Console.WriteLine();
            Console.WriteLine("Resources");

            //for (int i = 0; i < schedulerDataStorage1.Resources.Items.Count; i++)
            //{
            //    Console.WriteLine($"{schedulerDataStorage1.Resources[i].Id} {schedulerDataStorage1.Resources[i].Caption}");
            //}
        }
        private void InitializeAppointments()
        {
            //bool grouped;

            schedulerDataStorage1.Appointments.Clear();
            schedulerDataStorage1.Appointments.CustomFieldMappings.Clear();

            schedulerDataStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("JobNumber", "JobNumber"));
            schedulerDataStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("ProjectNumber", "ProjectNumber"));
            schedulerDataStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskID", "TaskID"));
            schedulerDataStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskName", "TaskName"));
            schedulerDataStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Component", "Component"));
            schedulerDataStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Hours", "Hours"));
            schedulerDataStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Predecessors", "Predecessors"));
            schedulerDataStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("DueDate", "DueDate"));

            AppointmentMappingInfo appointmentMappings = schedulerDataStorage1.Appointments.Mappings;

            appointmentMappings.AppointmentId = "ID";
            appointmentMappings.Start = "StartDate";
            appointmentMappings.End = "FinishDate";
            appointmentMappings.Subject = "Subject";
            //appointmentMappings.Location = "Location";
            appointmentMappings.PercentComplete = "PercentComplete";
            appointmentMappings.ResourceId = "Resources";
            appointmentMappings.Description = "Notes";            

            // TODO: Uncomment the code below to setup resource tree.

            //if (GroupByRadioGroup.SelectedIndex == 0)
            //{
            //    if (departmentComboBox.Text.Contains("CNC"))
            //    {
            //        appointmentMappings.ResourceId = "Machine";
            //    }
            //    else if (departmentComboBox.Text.Contains("Program"))
            //    {
            //        appointmentMappings.ResourceId = "Resource";
            //    }

            //    grouped = true;
            //}
            //else
            //{
            //    grouped = false;
            //}

            //foreach (var item in TasksList)
            //{

            //}

            //schedulerDataStorage1.Appointments.DataSource = Database.GetAppointmentData("All");

            BindingList<TaskModel> tasks = new BindingList<TaskModel>(TasksList);

            schedulerDataStorage1.Appointments.DataSource = tasks;

            for (int i = 0; i < schedulerDataStorage1.Appointments.Items.Count; i++)
            {
                Console.WriteLine($"{schedulerDataStorage1.Appointments[i].Id} Project#: {schedulerDataStorage1.Appointments[i].CustomFields["ProjectNumber"]} {schedulerDataStorage1.Appointments[i].Start} Resource: {schedulerDataStorage1.Appointments[i].ResourceId} TaskName: {schedulerDataStorage1.Appointments[i].CustomFields["TaskName"]}");
            }

            //Console.WriteLine();
            //Console.WriteLine($"{departmentComboBox.Text} Appointments");
        }
        private bool UpdateTaskStorage1(TaskModel movedTask, Appointment apt, PersistentObjectsEventArgs e)
        {            
            bool retryHit = false;
            ProjectModel globalProject = null;
            ComponentModel globalComponent = null;
            TaskModel globalTask = null;
            //AdvPersistentObjectsEventArgs e2 = (AdvPersistentObjectsEventArgs)e;

            //MessageBox.Show(e2.PropertyName);

            gridView1.BeginUpdate();
            gridView5.BeginUpdate();

            movedTask.Resources = GenerateResourceIDsString(apt.ResourceIds);
            movedTask.Machine = GetResourceFromResourceIDs(apt.ResourceIds, "Machine");
            movedTask.Personnel = GetResourceFromResourceIDs(apt.ResourceIds, "Person");
            movedTask.PercentComplete = apt.PercentComplete;
            //movedTask.StartDate = apt.Start;
            //movedTask.FinishDate = apt.End;

            gridView5.EndUpdate();
            gridView1.EndUpdate();

            Retry:

            globalProject = ProjectsList.Find(x => x.ProjectNumber == movedTask.ProjectNumber);

            globalComponent = ComponentsList.Find(x => x.Component == movedTask.Component && x.ProjectNumber == movedTask.ProjectNumber); 

            globalTask = globalComponent.Tasks.Find(x => x.TaskID == movedTask.TaskID);

            if ((globalProject == null || globalComponent == null || globalTask == null) && retryHit == false)
            {
                RefreshProjectGrid();
                retryHit = true;
                goto Retry;
            }

            try
            {
                if (AppointmentDateChanged)
                {
                    gridView5.BeginUpdate();
                    schedulerControl1.BeginUpdate();

                    if (RightMouseButtonPressed)
                    {
                        foreach (TaskModel task in globalComponent.Tasks)
                        {
                            Console.WriteLine($"Task: {task.TaskName,-13} Start Date: {((DateTime)task.StartDate).ToShortDateString(),-10} Finish Date: {GeneralOperations.AddBusinessDays((DateTime)task.StartDate, task.Duration).ToShortDateString()}");
                        }

                        Console.WriteLine();

                        if (MoveSubsequentTaskWithLockedSpacing)
                        {
                            if (((DateTime)movedTask.StartDate - OldTaskStartDate).Days > 0)
                            {
                                globalComponent.UpdateSuccessorTaskDates(movedTask, GeneralOperations.GetWorkingDays(OldTaskStartDate, (DateTime)movedTask.StartDate));
                            }
                            else
                            {
                                foreach (int taskID in movedTask.GetPredecessorList())
                                {
                                    globalComponent.UpdatePredecessorTaskDates(taskID, GeneralOperations.GetWorkingDays((DateTime)movedTask.StartDate, OldTaskStartDate));
                                }
                            }

                            Database.UpdateTaskDates(globalComponent.Tasks);

                            MoveSubsequentTaskWithLockedSpacing = false;
                        }
                        else
                        {
                            globalComponent.UpdateTaskDates(movedTask, OldTaskStartDate);
                        }

                        return true;
                    }
                    else
                    {
                        if (globalComponent.UpdateTaskDates(movedTask))
                        {
                            return true;
                        }
                    } 
                }
                else
                {
                    Database.UpdateTask(movedTask);
                    return true;
                }
            }
            finally
            {
                schedulerControl1.EndUpdate();
                schedulerControl1.RefreshData();
                gridView5.EndUpdate();

                gridView3.BeginUpdate();
                globalProject.LatestFinishDate = globalProject.GetLatestFinishDate();
                gridView3.EndUpdate();
            }

            return false;
        }

        /// <summary>
        /// Gets the last selected machine in resource list.
        /// </summary>
        private string GetResourceFromResourceIDs(AppointmentResourceIdCollection appointmentResourceIdCollection, string resourceType)
        {
            string id = "";
            foreach (var item in appointmentResourceIdCollection)
            {
                // This just validates that the selected resource is a machine and not a person.  It assumes that the resource list is comprised of both people and machines.
                if (ResourceDataTable.AsEnumerable().Where(x => x.Field<string>("ResourceName") == item.ToString() && x.Field<string>("ResourceType") == resourceType).Count() >= 1)
                {
                    if (item.ToString() == "No Machine" || item.ToString() == "No Personnel")
                    {
                        id = "";
                    }
                    else
                    {
                        id = item.ToString();
                    }

                    Console.WriteLine($"Machine: {id}");
                }
            }

            return id;
        }

        private string GenerateResourceIDsString(AppointmentResourceIdCollection appointmentResourceIdCollection)
        {
            AppointmentResourceIdCollectionXmlPersistenceHelper helper = new AppointmentResourceIdCollectionXmlPersistenceHelper(appointmentResourceIdCollection);
            return helper.ToXml();
        }

        private ResourceIdCollection GenerateResourceIDsString(string xml)
        {
            ResourceIdCollection result = new ResourceIdCollection();
            if (String.IsNullOrEmpty(xml))
                return result;

            return AppointmentResourceIdCollectionXmlPersistenceHelper.ObjectFromXml(result, xml);
        }

        private Color GetDeptColor(string department)
        {
            Color departmentColor;

            if (department == "Design")
            {
                departmentColor = Color.LightBlue;
            }
            else if (department == "Program Rough")
            {
                departmentColor = Color.LightCoral;
            }
            else if (department == "Program Electrodes")
            {
                departmentColor = Color.LightGreen;
            }
            else if (department == "Program Finish")
            {
                departmentColor = Color.LightPink;
            }
            else if (department == "CNC Rough")
            {
                departmentColor = Color.Orange;
            }
            else if (department == "CNC Electrodes")
            {
                departmentColor = Color.Green;
            }
            else if (department == "CNC Finish")
            {
                departmentColor = Color.Aquamarine;
            }
            else if (department == "EDM Sinker")
            {
                departmentColor = Color.DodgerBlue;
            }
            else if (department.Contains("Grind"))
            {
                departmentColor = Color.NavajoWhite;
            }
            else if (department == "Heat Treat")
            {
                departmentColor = Color.Red;
            }
            else if (department.Contains("EDM Wire"))
            {
                departmentColor = Color.Gold;
            }
            else if (department.Contains("Polish"))
            {
                departmentColor = Color.Aqua;
            }
            else if (department.Contains("Inspection"))
            {
                departmentColor = Color.BlueViolet;
            }
            else if (department == "Hole Pop")
            {
                departmentColor = Color.Honeydew;
            }
            else
            {
                departmentColor = Color.LightGray;
            }

            return departmentColor;
        }

        private Color GetTaskReadinessColor(TaskModel task)
        {
            var result = from tempTask in TasksList
                         where tempTask.ProjectNumber == task.ProjectNumber && tempTask.Component == task.Component && task.HasMatchingPredecessor(tempTask.TaskID)
                         select tempTask;

            foreach (var item in result)
            {
                if (item.Status != "Completed")
                {
                    return Color.LightSalmon;
                }
            }

            return Color.LightGreen;
        }

        private void SetTasks()
        {
            string department = departmentComboBox.Text;

            if (department == "Design")
            {
                Tasks = "Design";
            }
            else if (department == "Program Rough")
            {
                Tasks = "Program Rough";
            }
            else if (department == "Program Finish")
            {
                Tasks = "Program Finish";
            }
            else if (department == "Program Electrodes")
            {
                Tasks = "Program Electrodes";
            }
            else if (department == "Programming")
            {
                Tasks = @"Program\w*";
            }
            else if (department == "CNC Rough")
            {
                Tasks = @"CNC Rough";
            }
            else if (department == "Rough")
            {
                Tasks = @"Rough";
            }
            else if (department == "CNC Finish")
            {
                Tasks = @"^CNC Finish\w*";
            }
            else if (department == "CNC Electrodes")
            {
                Tasks = @"^CNC Electrodes";
            }
            else if (department == "CNCs")
            {
                Tasks = @"^CNC\w*";
            }
            else if (department == "CNC People")
            {
                Tasks = "All";
            }
            else if (department == "EDM Sinker")
            {
                Tasks = @"^EDM Sinker\w*";
            }
            else if (department == "EDM Wire (In-House)")
            {
                Tasks = @"^EDM Wire \(In-House\)\w*";
            }
            else if (department == "Grind")
            {
                Tasks = @"^*Grind\w*";
            }
            else if (department == "Polish")
            {
                Tasks = @"Polish \(In-House\)\w*";
            }
            else if (department == "Inspection")
            {
                Tasks = @"^Inspection\w*";
            }
            else if (department == "Mold Service")
            {
                Tasks = @"Mold Service";
            }
            else if (department == "Manual")
            {
                Tasks = @"Manual";
            }
            else if (department == "All")
            {
                Tasks = ".*";
            }

            TaskRegExpression = new Regex(Tasks);
        }

        private void SetRole()
        {
            string department = departmentComboBox.Text;
            schedulerControl1.ActiveView.ResourcesPerPage = 0;

            if (department == "Design")
            {
                Role = "Design";
                NoResourceName = "No Personnel";
            }
            else if (department == "Program Rough")
            {
                Role = "Rough Programmer";
                NoResourceName = "No Personnel";
            }
            else if (department == "Program Finish")
            {
                Role = "Finish Programmer";
                NoResourceName = "No Personnel";
            }
            else if (department == "Program Electrodes")
            {
                Role = "Electrode Programmer";
                NoResourceName = "No Personnel";
            }
            else if (department == "Programming")
            {
                Role = "Programmer";
                NoResourceName = "No Personnel";
            }
            else if (department == "CNC Rough")
            {
                Role = "Rough Mill";
                NoResourceName = "No Machine";
            }
            else if (department == "Rough") // This option is not available in the Department Schedule View.  And that's okay.
            {
                Role = "Rough";
            }
            else if (department == "CNC Finish")
            {
                Role = "Finish Mill";
                NoResourceName = "No Machine";
            }
            else if (department == "CNC Electrodes")
            {
                Role = "Graphite Mill";
                NoResourceName = "No Machine";
            }
            else if (department == "CNCs")
            {
                Role = "Mill";
                NoResourceName = "No Machine";
            }
            else if (department == "CNC People")
            {
                Role = "CNC Operator";
                NoResourceName = "No Personnel";
                schedulerControl1.ActiveView.ResourcesPerPage = 8;
            }
            else if (department == "EDM Sinker")
            {
                Role = @"^(EDM Sinker)$";
                NoResourceName = "No Machine";
            }
            else if (department == "EDM Wire (In-House)")
            {
                Role = @"^EDM Wire$";
                NoResourceName = "No Machine";
            }
            else if (department == "Grind")
            {
                Role = "Tool Maker";
                NoResourceName = "No Personnel";
                schedulerControl1.ActiveView.ResourcesPerPage = 8;
            }
            else if (department == "Polish")
            {
                Role = "Tool Maker";
                NoResourceName = "No Personnel";
                schedulerControl1.ActiveView.ResourcesPerPage = 8;
            }
            else if (department == "Inspection")
            {
                Role = "CMM Operator";
                NoResourceName = "No Personnel";
            }
            else if (department == "Mold Service")
            {
                Role = "Tool Maker";
                NoResourceName = "No Personnel";
                schedulerControl1.ActiveView.ResourcesPerPage = 8;
            }
            else if (department == "Manual")
            {
                Role = "Manual";
                NoResourceName = "No Personnel";
                schedulerControl1.ActiveView.ResourcesPerPage = 8;
            }
            else if (department == "All")
            {
                Role = "All";
                schedulerControl1.ActiveView.ResourcesPerPage = 8;
            }

            RoleRegExpression = new Regex(Role);
        }

        public static bool Like(string toSearch, string toFind)
        {
            return new Regex(@"\A" + new Regex(@"\.|\$|\^|\{|\[|\(|\||\)|\*|\+|\?|\\").Replace(toFind, ch => @"\" + ch).Replace('_', '.').Replace("%", ".*") + @"\z", RegexOptions.Singleline).IsMatch(toSearch);
        }

        private void RefreshDepartmentScheduleView()
        {
            ResourceDataTable = Database.GetResourceData();
            InitializeResources();
            LoadProjects();
            BindingList<TaskModel> tasks = new BindingList<TaskModel>(TasksList);

            schedulerDataStorage1.Appointments.DataSource = tasks;
            //InitializeAppointments();
            schedulerControl1.RefreshData();
        }

        private void PopulateProjectCheckedComboBox()
        {
            projectCheckedComboBoxEdit.Properties.Items.Clear();

            foreach (var item in Database.GetJobNumberComboList())
            {
                projectCheckedComboBoxEdit.Properties.Items.Add(item, true);
            }

            AllProjectsChecked();
        }

        private void AllProjectsChecked()
        {
            if (projectCheckedComboBoxEdit.Properties.Items.Count == projectCheckedComboBoxEdit.Properties.Items.Count(x => x.CheckState == CheckState.Checked))
            {
                AllProjectItemsChecked = true;
            }
            else
            {
                AllProjectItemsChecked = false;
            }
        }

        private bool IsFormOpen(string formName)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                //iterate through
                if (frm.Name == formName)
                {
                    return true;
                }
            }

            return false;
        }

        private void departmentComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //InitializeAppointments();
            SetRole();
            SetTasks();
            //appointmentstorage.FilterCriteria = FilterTaskView(departmentComboBox.Text, includeQuotesCheckEdit.Checked, includeCompletesCheckEdit.Checked);
            //schedulerControl1.SchedulerDataStorage.Appointments.FilterCriteria = FilterTaskView(departmentComboBox.Text, includeQuotesCheckEdit.Checked, includeCompletesCheckEdit.Checked);
            schedulerControl1.ActiveView.LayoutChanged();
            //if (departmentComboBox.Text.Contains("Program") || departmentComboBox.Text.Contains("CNC"))
            //{
            //    InitializeResources();
            //}
            //else
            //{

            //}
        }

        private void projectCheckedComboBoxEdit_EditValueChanged(object sender, EventArgs e)
        {
            AllProjectsChecked();
            schedulerControl1.ActiveView.LayoutChanged();
        }

        private void refreshButton_Click(object sender, EventArgs e)
        {
            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                RefreshDepartmentScheduleView();                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                SplashScreenManager.CloseForm();
            }
        }

        private void GroupByRadioGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            RadioGroup edit = sender as RadioGroup;

            if (edit.SelectedIndex == 0)
            {
                //InitializeAppointments();
                //InitializeResources();
                SetRole();
                schedulerControl1.ActiveView.LayoutChanged();
                schedulerControl1.GroupType = SchedulerGroupType.Resource;

            }
            else
            {
                //InitializeAppointments();
                //InitializeResources();

                if (departmentComboBox.Text.Contains("Rough"))
                {
                    Role = "Rough";
                }

                schedulerControl1.GroupType = SchedulerGroupType.None;
                schedulerControl1.ActiveView.LayoutChanged();
            }
        }

        private void includeCompletesCheckEdit_CheckStateChanged(object sender, EventArgs e)
        {
            schedulerControl1.ActiveView.LayoutChanged();
        }

        private void includeQuotesCheckEdit_CheckedChanged(object sender, EventArgs e)
        {
            schedulerControl1.ActiveView.LayoutChanged();
        }
        private void zoomTrackBarControl1_ValueChanged(object sender, EventArgs e)
        {
            schedulerControl1.GanttView.GetBaseTimeScale().Width = zoomTrackBarControl1.Value;
            Console.WriteLine(zoomTrackBarControl1.Value);
        }
        private void schedulerControl1_AllowAppointmentDelete(object sender, AppointmentOperationEventArgs e)
        {
            e.Allow = false;
        }        
        private void schedulerControl1_AppointmentDrop(object sender, AppointmentDragEventArgs e)
        {
            // Use this event to handle moving multiple selected tasks.

            // Gets rid of the resource that the task originally placed on before it was moved while preserving the other resources.
            foreach (var id in e.SourceAppointment.ResourceIds)
            {
                if (!Equals(id, DraggedResourceId))
                {
                    e.EditedAppointment.ResourceIds.Add(id);
                }
            }
        }
        private void schedulerControl1_AppointmentFlyoutShowing(object sender, AppointmentFlyoutShowingEventArgs e)
        {
            TaskModel task = new TaskModel();

            task.JobNumber = e.FlyoutData.Appointment.CustomFields["JobNumber"].ToString();
            task.ProjectNumber = (int)e.FlyoutData.Appointment.CustomFields["ProjectNumber"];
            task.Component = e.FlyoutData.Appointment.CustomFields["Component"].ToString();
            task.TaskName = e.FlyoutData.Appointment.CustomFields["TaskName"].ToString();
            task.Hours = (int)e.FlyoutData.Appointment.CustomFields["Hours"];
            task.DueDate = ProjectsList.Find(x => x.ProjectNumber == task.ProjectNumber).DueDate;
            task.ComponentPicture = ComponentsList.Find(x => x.ProjectNumber == task.ProjectNumber && x.Component == task.Component).picture;
            task.Notes = e.FlyoutData.Appointment.Description;

            e.Control = CreateLabel(task);
        }
        private void schedulerControl1_AppointmentResized(object sender, AppointmentResizeEventArgs e)
        {
            //MessageBox.Show("Resize");
        }
        private void schedulerControl1_AppointmentsDrag(object sender, AppointmentsDragEventArgs e)
        {
            //MessageBox.Show("Drag");
        }
        private void schedulerControl1_DragDrop(object sender, DragEventArgs e)
        {
            //MessageBox.Show("DragDrop");
        }
        private void schedulerControl1_DragOver(object sender, DragEventArgs e)
        {
            Point pos = schedulerControl1.PointToClient(Cursor.Position);
            SchedulerViewInfoBase viewInfo = schedulerControl1.ActiveView.ViewInfo;
            HitInfo = viewInfo.CalcHitInfo(pos, false);
        }
        private void schedulerControl1_AllowAppointmentDrag(object sender, AppointmentOperationEventArgs e)
        {
            AppointmentErrorSent = false;
            AppointmentDateChanged = false;
            // Prevents user from dragging multiple tasks since doing so causes undesirable results.
            if (schedulerControl1.SelectedAppointments.Count > 1)
            {
                //MessageBox.Show("Cannot drag multiple tasks.");
                e.Allow = false;
            }
        }
        private void schedulerControl1_MouseDown(object sender, MouseEventArgs e)
        {
            var scheduler = sender as DevExpress.XtraScheduler.SchedulerControl;
            var hitInfo = scheduler.ActiveView.CalcHitInfo(e.Location, false);

            if (e.Button == MouseButtons.Right)
            {
                RightMouseButtonPressed = true;

                if (hitInfo.HitTest == SchedulerHitTest.AppointmentContent)
                {
                    Appointment apt = ((AppointmentViewInfo)hitInfo.ViewInfo).Appointment;
                    DraggedAppointment = apt;
                }
            }

            if (hitInfo.HitTest != DevExpress.XtraScheduler.Drawing.SchedulerHitTest.AppointmentContent)
                return;

            DraggedResourceId = ((DevExpress.XtraScheduler.Internal.Implementations.ResourceBase)hitInfo.ViewInfo.Resource).Id;

            AppointmentResourceChanged = false;
            EditAppointmentByForm = false;
        }
        private void schedulerControl1_EditAppointmentFormShowing(object sender, AppointmentFormEventArgs e)
        {
            AppointmentErrorSent = false;
            AppointmentDateChanged = false;
            EditAppointmentByForm = true;
        }
        private void schedulerControl1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                RightMouseButtonPressed = false;
            }
        }
        private void schedulerDataStorage1_AppointmentChanging(object sender, PersistentObjectCancelEventArgs e)
        {
            try
            {
                // May want to go this route in the future.  But this would require switching schedulerStorage1 to a datastorage

                AdvPersistentObjectCancelEventArgs e2 = (AdvPersistentObjectCancelEventArgs)e;

                //MessageBox.Show(e2.PropertyName);

                if (e2.PropertyName == "Start")
                {
                    //TaskModel changingTask = ((Appointment)e.Object).GetSourceObject(schedulerDataStorage1) as TaskModel;

                    //ComponentModel tempComponent = ComponentsList.Find(x => x.Component == changingTask.Component && x.ProjectNumber == changingTask.ProjectNumber);

                    //if (RightMouseButtonPressed)
                    //{
                    //    tempComponent.UpdateTask(changingTask, ChangingTask);
                    //}
                    if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeDates))
                    {
                        e.Cancel = true;
                        throw new Exception("This login is not authorized to make changes to dates.");
                    }

                    AppointmentDateChanged = true;
                }
                else if (e2.PropertyName == "End")
                {
                    if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeDates))
                    {
                        e.Cancel = true;
                        throw new Exception("This login is not authorized to make changes to dates.");
                    }

                    AppointmentDateChanged = true;
                }
                else if (e2.PropertyName == "Reminders")
                {
                    e.Cancel = true;
                }
                else if (e2.PropertyName == "ResourceIds")
                {
                    //MessageBox.Show(((AppointmentResourceIdCollection)e2.OldValue)[0].ToString());
                    //MessageBox.Show(((Appointment)e.Object).ResourceIds[0].ToString());

                    if (!EditAppointmentByForm && ((DevExpress.XtraScheduler.Internal.Implementations.ResourceBase)HitInfo.ViewInfo.Resource).Id != DraggedResourceId)
                    {
                        if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeProjectData))
                        {
                            e.Cancel = true;
                            throw new Exception("This login is not authorized to make changes to project data.");
                        } 
                    }
                }
                else
                {
                    if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeProjectData))
                    {
                        e.Cancel = true;
                        throw new Exception("This login is not authorized to make changes to project data.");
                    }
                }

                if (RightMouseButtonPressed)
                {
                    TaskModel changingTask = ((Appointment)e.Object).GetSourceObject(schedulerDataStorage1) as TaskModel;

                    OldTaskStartDate = (DateTime)changingTask.StartDate;  // (DateTime)changingTask.StartDate

                    Console.WriteLine($"Old Date: {OldTaskStartDate}");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                if (!AppointmentErrorSent)
                {
                    MessageBox.Show(ex.Message); 
                }

                Console.WriteLine(ex.ToString());
                AppointmentErrorSent = true;
            }
        }
        private void schedulerDataStorage1_AppointmentsChanged(object sender, PersistentObjectsEventArgs e)
        {
            //LoadProjects();
            if (!IsFormOpen("WaitForm1") && !MoveSelectedAppointments)
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                Thread.Sleep(250);
            }

            TaskModel movedTask;

            try
            {
                foreach (Appointment apt in e.Objects)
                {
                    movedTask = apt.GetSourceObject(schedulerDataStorage1) as TaskModel;
                    
                    if (UpdateTaskStorage1(movedTask, apt, e))
                    {

                    }
                    else
                    {
                        RefreshDepartmentScheduleView();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if (IsFormOpen("WaitForm1") && !MoveSelectedAppointments)
                {
                    SplashScreenManager.CloseForm();
                }
            }

            //MessageBox.Show("AppointmentChanged");
        }
        private void schedulerControl1_PopupMenuShowing(object sender, DevExpress.XtraScheduler.PopupMenuShowingEventArgs e)
        {
            if (schedulerControl1.SelectedAppointments.Count > 1)
            {
                e.Menu = null;
                //MessageBox.Show(schedulerControl1.SelectedAppointments.Count.ToString());

                XtraInputBoxArgs args = new XtraInputBoxArgs();

                args.Caption = "Number of Days Delta";
                args.Prompt = "Number of Days";
                args.DefaultButtonIndex = 0;
                //args.Showing += Args_Showing;
                SpinEdit editor = new SpinEdit();
                editor.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
                editor.Properties.Mask.EditMask = "N0";
                editor.Properties.Mask.UseMaskAsDisplayFormat = true;

                args.Editor = editor;

                var result = XtraInputBox.Show(args)?.ToString();

                if (result != null && result.Length > 0)
                {
                    if (int.TryParse(result, out int numOfDays))
                    {
                        MoveSelectedAppointments = true;

                        List<int> selectedAptIDs = new List<int>();

                        selectedAptIDs = schedulerControl1.SelectedAppointments.Select(x => int.Parse(x.Id.ToString())).ToList();

                        foreach (int apt in selectedAptIDs)
                        {
                            Appointment appointment = schedulerDataStorage1.Appointments.Items.Where(x => int.Parse(x.Id.ToString()) == apt).First();

                            appointment.Start = appointment.Start.AddDays(numOfDays);
                        }

                        MoveSelectedAppointments = false;
                    }
                }
            }
            else
            {
                e.Menu.Items.Remove(e.Menu.Items.FirstOrDefault(x => x.Caption == "&Copy"));

                if (e.Menu.Items.Count(x => x.Caption == "Mo&ve") > 0)
                {
                    DXMenuItem menuItem = e.Menu.Items.FirstOrDefault(x => x.Caption == "Mo&ve");
                    menuItem.Caption = "Move All Component Tasks with Locked Spacing";
                }

                if (e.Menu.Items.Count(x => x.Caption == "Open Kan Ban") == 0)
                {
                    e.Menu.Items.Insert(1, new SchedulerMenuItem("Open Kan Ban", schedulerDataStorage1_OpenKanBan));
                }

                if (e.Menu.Items.Count(x => x.Caption == "Move Subsequent Component Tasks with Lock Spacing") == 0)
                {
                    e.Menu.Items.Insert(2, new SchedulerMenuItem("Move Subsequent Component Tasks with Lock Spacing", schedulerDataStorage1_MoveSubsequentComponentTasksWithLockedSpacing));
                }
            }
        }
        private void schedulerDataStorage1_OpenKanBan(object sender, EventArgs e)
        {
            string filePath = ProjectsList.Find(x => x.ProjectNumber == int.Parse(DraggedAppointment.CustomFields["ProjectNumber"].ToString())).KanBanWorkbookPath;
            string component = DraggedAppointment.CustomFields["Component"].ToString();
            ExcelInteractions.OpenKanBanWorkbook(filePath, component);
        }
        private void schedulerDataStorage1_MoveSubsequentComponentTasksWithLockedSpacing(object sender, EventArgs e)
        {
            MoveSubsequentTaskWithLockedSpacing = true;

            //MessageBox.Show($"Final Start: {HitInfo.ViewInfo.Interval.Start.ToShortDateString()} Finish: {HitInfo.ViewInfo.Interval.End.ToShortDateString()}");

            DraggedAppointment.Start = HitInfo.ViewInfo.Interval.Start;
            DraggedAppointment.End = HitInfo.ViewInfo.Interval.End;

        }
        private void schedulerDataStorage1_FilterAppointment(object sender, PersistentObjectCancelEventArgs e)
        {
            Appointment apt = (Appointment)e.Object;

            try
            {
                if (AllProjectItemsChecked == false)
                {
                    e.Cancel = !projectCheckedComboBoxEdit.Properties.Items.Where(x => x.Value.ToString().Contains($"#{apt.CustomFields["ProjectNumber"]}") && x.CheckState == CheckState.Checked && TaskRegExpression.IsMatch($"{apt.CustomFields["TaskName"]}")).Any(); // 
                }
                else
                {
                    e.Cancel = !TaskRegExpression.IsMatch($"{apt.CustomFields["TaskName"]}");
                }

                if (includeCompletesCheckEdit.Checked == false)
                {
                    if (apt.PercentComplete == 100)
                    {
                        e.Cancel = true;
                    }
                }

                if (includeQuotesCheckEdit.Checked == false)
                {
                    if (apt.CustomFields["JobNumber"].ToString().Contains("quote") || apt.CustomFields["JobNumber"].ToString().Contains("Quote"))
                    {
                        e.Cancel = true;
                    }
                }

                //e.Cancel = !apt.Location.Contains(Tasks);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }
        private void schedulerDataStorage1_FilterResource(object sender, PersistentObjectCancelEventArgs e)
        {
            try
            {
                SchedulerDataStorage storage = (SchedulerDataStorage)sender;
                Resource res = (Resource)e.Object;

                if (Role != "All")
                {
                    if (GroupByRadioGroup.SelectedIndex == 0)
                    {
                        if (Role != "Manual")
                        {
                            //TODO: Modify this to use resource custom fields.
                            //e.Cancel = !(RoleRegExpression.IsMatch(res.CustomFields["Role"].ToString()) || res.Id.ToString() == NoResourceName);
                            e.Cancel = ResourceDataTable.AsEnumerable().Where(x => x.Field<string>("ResourceName") == res.Id.ToString() && (RoleRegExpression.IsMatch(x.Field<string>("Role")) || x.Field<string>("ResourceName") == NoResourceName || x.Field<string>("ResourceName") == "None")).Count() < 1;
                        }
                        else
                        {
                            e.Cancel = res.CustomFields["ResourceType"].ToString() != "Person";
                        }
                    }
                    else if (GroupByRadioGroup.SelectedIndex == 1)
                    {
                        Console.WriteLine($"Resource: {res.Id}");
                        //TODO: Modify this to use resource custom fields.
                        e.Cancel = ResourceDataTable.AsEnumerable().Where(x => x.Field<string>("ResourceName") == res.Id.ToString() && (x.Field<string>("Department").Contains(departmentComboBox.Text) || x.Field<string>("ResourceName") == NoResourceName || x.Field<string>("ResourceName") == "None")).Count() < 1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }
        private void schedulerControl1_AppointmentViewInfoCustomizing(object sender, AppointmentViewInfoCustomizingEventArgs e)
        {
            if (Tasks == ".*")
            {
                e.ViewInfo.Appearance.BackColor = GetDeptColor(e.ViewInfo.Appointment.CustomFields["TaskName"].ToString());
            }
            else
            {
                e.ViewInfo.Appearance.BackColor = GetTaskReadinessColor((TaskModel)e.ViewInfo.Appointment.GetSourceObject(schedulerDataStorage1));
            }
        }

        #endregion

        #region Department Task View

        private void InitializeDepartmentPrintOptions()
        {
            // string[] departmentArr = { "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "CNC Electrodes", "Grind", "Inspection", "EDM Sinker", "EDM Wire (In-House)", "Polish" };

            foreach (string item in DepartmentArr)
            {
                PrintDeptsCheckedComboBoxEdit.Properties.Items.Add(item, CheckState.Unchecked, true);
            }

            PrintDeptsCheckedComboBoxEdit.Properties.SeparatorChar = ',';
            // This line sets which items in the checkedcombobox are checked by default.
            PrintDeptsCheckedComboBoxEdit.SetEditValue("Program Rough, Program Finish, Program Electrodes, CNC Rough, CNC Finish, CNC Electrodes, Grind, Inspection, EDM Sinker, EDM Wire (In-House)");
        }
        private void InitializePersonnelPrintOptions()
        {
            PrintEmployeeWorkCheckedComboBoxEdit.Properties.Items.Clear();

            var result = (from task in TasksList
                          where task.Personnel != null && task.Personnel != "" && task.Status != null
                          orderby task.Personnel
                          select task.Personnel).Distinct().ToList();

            foreach (string item in result) // ResourceDataTable.AsEnumerable().Where(x => x.Field<string>("Role") != "Engineer" && x.Field<string>("ResourceType") == "Person").Select(x => x.Field<string>("ResourceName"))
            {
                PrintEmployeeWorkCheckedComboBoxEdit.Properties.Items.Add(item, CheckState.Unchecked, true);
            }

            //BindingList personnel = new BindingList(ResourceDataTable.AsEnumerable().Where(x => x.Field<string>("Role") != "Engineer" && x.Field<string>("ResourceType") == "Person"));

            //PrintEmployeeWorkCheckedComboBoxEdit.Properties.DataSource = new BindingList<string>(Database.GetPersonnel());

            //PrintEmployeeWorkCheckedComboBoxEdit.Properties.ValueMember = "ResourceName";

            //PrintEmployeeWorkCheckedComboBoxEdit.Properties.DisplayMember = "ResourceName";

            //PrintEmployeeWorkCheckedComboBoxEdit.Properties.SeparatorChar = ',';

            //PrintEmployeeWorkCheckedComboBoxEdit.SetEditValue("Program Rough, Program Finish, Program Electrodes, CNC Rough, CNC Finish, CNC Electrodes, Grind, Inspection, EDM Sinker, EDM Wire (In-House)");
        }
        private CriteriaOperator FilterTaskView(string department, bool includeQuotes, bool includeCompleteTasks, bool filterTasksByDates)
        {
            List<CriteriaOperator> criteriaOperators = new List<CriteriaOperator>();

            DateTime outlookDate = DateTime.Today.AddDays((int)daysAheadSpinEdit.Value);

            if (includeQuotes == false)
            {
                criteriaOperators.Add(new NotOperator(new FunctionOperator(FunctionOperatorType.Contains, new OperandProperty("JobNumber"), "Quote"))); // Excludes tasks with quote in jobnumber. 
            }

            if (includeCompleteTasks == false)
            {
                criteriaOperators.Add(new NullOperator("Status"));  // Excludes tasks with Status set to null. 
            }

            if (filterTasksByDates == true)
            {
                criteriaOperators.Add(CriteriaOperator.Or( new BetweenOperator(outlookDate, new OperandProperty("StartDate"), new OperandProperty("FinishDate")), new BinaryOperator(outlookDate, new OperandProperty("FinishDate"), BinaryOperatorType.GreaterOrEqual)));
            }

            if (department == "Design")
            {
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "Program")
            {
                criteriaOperators.Add(new FunctionOperator(FunctionOperatorType.StartsWith, new OperandProperty("TaskName"), department));
            }
            else if (department == "Program Rough")
            {
                //gridView1.ActiveFilterString = "[TaskName] = 'Program Rough'  AND [Status] = NULL";
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "Program Finish")
            {
                //gridView1.ActiveFilterString = "[TaskName] = 'Program Finish' AND [Status] = NULL";
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "Program Electrodes")
            {
                //gridView1.ActiveFilterString = "[TaskName] = 'Program Electrodes' AND [Status] = NULL";
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "CNC")
            {
                criteriaOperators.Add(new FunctionOperator(FunctionOperatorType.StartsWith, new OperandProperty("TaskName"), department));
            }
            else if (department == "CNC Rough")
            {
                //gridView1.ActiveFilterString = "[TaskName] = 'CNC Rough' AND [Status] = NULL";
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "CNC Finish")
            {
                //gridView1.ActiveFilterString = "[TaskName] = 'CNC Finish' AND [Status] = NULL";
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "CNC Electrodes")
            {
                //gridView1.ActiveFilterString = "[TaskName] = 'CNC Electrodes' AND [Status] = NULL";
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "Grind")
            {
                //gridView1.ActiveFilterString = "[TaskName] LIKE '%Grind' AND [Status] = NULL";
                criteriaOperators.Add(new FunctionOperator(FunctionOperatorType.EndsWith, new OperandProperty("TaskName"), department));
            }
            else if (department == "EDM Sinker")
            {
                //gridView1.ActiveFilterString = "[TaskName] = 'EDM Sinker' AND [Status] = NULL";
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "EDM Wire (In-House)")
            {
                //gridView1.ActiveFilterString = "[TaskName] = 'EDM Wire (In-House)' AND [Status] = NULL";
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "Polish")
            {
                //gridView1.ActiveFilterString = "[TaskName] LIKE 'Polish%' AND [Status] = NULL";
                criteriaOperators.Add(new FunctionOperator(FunctionOperatorType.StartsWith, new OperandProperty("TaskName"), department));
            }
            else if (department == "Inspection")
            {
                //gridView1.ActiveFilterString = "[TaskName] LIKE 'Inspection%' AND [Status] = NULL";
                criteriaOperators.Add(new FunctionOperator(FunctionOperatorType.StartsWith, new OperandProperty("TaskName"), department));
            }
            else if (department == "Mold Service")
            {
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "Manual")
            {
                criteriaOperators.Add(new BinaryOperator("TaskName", department, BinaryOperatorType.Equal));
            }
            else if (department == "All")
            {
                //gridView1.ActiveFilterString = String.Empty;
                //gridView1.ClearColumnsFilter();
            }

            footerDateTime = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

            return CriteriaOperator.And(criteriaOperators);
        }

        // This filter is used by the print resources function.
        private void FilterTaskView2(string resource)
        {
            List<CriteriaOperator> criteriaOperators = new List<CriteriaOperator>();

            criteriaOperators.Add(new NotOperator(new FunctionOperator(FunctionOperatorType.Contains, new OperandProperty("JobNumber"), "Quote")));
            criteriaOperators.Add(new NullOperator("Status"));

            criteriaOperators.Add(new BinaryOperator("Personnel", resource, BinaryOperatorType.Equal));

            gridView1.ActiveFilterCriteria = CriteriaOperator.And(criteriaOperators);

            footerDateTime = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
        }
        private CriteriaOperator FilterTaskView3()
        {
            List<CriteriaOperator> criteriaOperators = new List<CriteriaOperator>();

            criteriaOperators.Add(new NotOperator(new FunctionOperator(FunctionOperatorType.Contains, new OperandProperty("Stage"), "On Hold")));
            criteriaOperators.Add(new NotOperator(new FunctionOperator(FunctionOperatorType.Contains, new OperandProperty("Stage"), "Completed")));
            criteriaOperators.Add(new NotOperator(new FunctionOperator(FunctionOperatorType.Contains, new OperandProperty("Stage"), "Closed")));

            return CriteriaOperator.And(criteriaOperators);
        }
        // This header is for when the datagrid gets printed.
        private void CreateHeaderRTFString()
        {
            StringBuilder tableRtf = new StringBuilder();

            tableRtf.Append(@"{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Microsoft Sans Serif;}}");
            for (int j = 0; j < 1; j++) // j represents the number of rows to create.
            {
                // Start the row.
                tableRtf.Append(@"\trowd\b");

                tableRtf.Append(@"\clbrdrt\brdrs\clbrdrl\brdrs\clbrdrb\brdrs\clbrdrr\brdrs");
                // First cell with width 1000. Font style to bold.
                tableRtf.Append(@"\cellx3838");
                tableRtf.Append(@"\clbrdrt\brdrs\clbrdrl\brdrs\clbrdrb\brdrs\clbrdrr\brdrs");
                tableRtf.Append(@" Job Number: All");
                tableRtf.Append(@"\intbl\cell");
                tableRtf.Append(@"\cellx7676\qc");
                tableRtf.Append(@"\intbl");
                tableRtf.Append(" " + "Department: " + departmentComboBox2.Text);
                tableRtf.Append(@"\cell");
                tableRtf.Append(@"\clbrdrt\brdrs\clbrdrl\brdrs\clbrdrb\brdrs\clbrdrr\brdrs");
                tableRtf.Append(@"\cellx11514");
                tableRtf.Append(" " + "Components: All");
                tableRtf.Append(@"\intbl\cell");
                // Append the row in StringBuilder.
                tableRtf.Append(@"\b0 \row");
            }

            tableRtf.Append(@"\pard");
            tableRtf.Append(@"}");

            //richTextBox1.Rtf = tableRtf.ToString();
        }

        private void OpenKanBanWorkbook(int rowIndex) // This method is for opening KanBan Workbooks from the Department Task View grid.
        {
            TaskModel task = gridView1.GetRow(rowIndex) as TaskModel;

            if (rowIndex >= 0)
            {
                ExcelInteractions.OpenKanBanWorkbook(Database.GetKanBanWorkbookPath(task.JobNumber, task.ProjectNumber), task.Component);
            }
        }

        public static void ChangeCellBorderColor(TableCell cell)
        {
            //Specify the border style and the background color for the header cells 
            cell.Borders.Bottom.LineStyle = TableBorderLineStyle.None;
            cell.Borders.Left.LineStyle = TableBorderLineStyle.None;
            cell.Borders.Right.LineStyle = TableBorderLineStyle.None;
            cell.Borders.Top.LineStyle = TableBorderLineStyle.None;
        }

        private void departmentComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            gridView1.ActiveFilterCriteria = FilterTaskView(departmentComboBox2.Text, false, false, filterTasksByDatesCheckEdit.Checked);
        }
        private void filterTasksByDatesCheckEdit_CheckedChanged(object sender, EventArgs e)
        {
            gridView1.ActiveFilterCriteria = FilterTaskView(departmentComboBox2.Text, false, false, filterTasksByDatesCheckEdit.Checked);
        }
        private void daysAheadSpinEdit_ValueChanged(object sender, EventArgs e)
        {
            gridView1.ActiveFilterCriteria = FilterTaskView(departmentComboBox2.Text, false, false, filterTasksByDatesCheckEdit.Checked);
        }
        private void RefreshTasksButton_Click(object sender, EventArgs e)
        {
            try
            {
                deptTaskViewHelper = new RefreshHelper(gridView1, "ProjectNumber");
                RoleTable = Database.GetRoleTable();
                deptTaskViewHelper.SaveViewInfo();
                LoadProjects();
                LoadTaskView();
                InitializePersonnelPrintOptions();
                deptTaskViewHelper.LoadViewInfo();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }
        private void printTaskViewButton_Click(object sender, EventArgs e)
        {
            // Check whether the GridControl can be previewed.
            if (!gridControl1.IsPrintingAvailable)
            {
                MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                return;
            }

            gridView1.Columns["Personnel"].OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.True;
            gridView1.Columns["TaskName"].OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            //// Used richEditControl to generate desired Rtf code.
            //gridView1.OptionsPrint.RtfPageHeader = @"{\rtf1\deff0{\fonttbl{\f0 Calibri;}{\f1 Microsoft Sans Serif;}}{\colortbl ;\red0\green0\blue255 ;}{\*\defchp \b\f1\fs22}{\stylesheet {\ql\b\f1\fs22 Normal;}{\*\cs1\b\f1\fs22 Default Paragraph Font;}{\*\cs2\sbasedon1\b\f1\fs22 Line Number;}{\*\cs3\b\ul\f1\fs22\cf1 Hyperlink;}{\*\ts4\tsrowd\b\f1\fs22\ql\tscellpaddfl3\tscellpaddl108\tscellpaddfb3\tscellpaddfr3\tscellpaddr108\tscellpaddft3\tsvertalt\cltxlrtb Normal Table;}{\*\ts5\tsrowd\sbasedon4\b\f1\fs22\ql\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\tscellpaddfl3\tscellpaddl108\tscellpaddfr3\tscellpaddr108\tsvertalt\cltxlrtb Table Simple 1;}}{\*\listoverridetable}{\info{\creatim\yr2018\mo1\dy10\hr10\min20}{\version1}}\nouicompat\splytwnine\htmautsp\sectd\trowd\irow0\irowband-1\lastrow\ts5\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\trleft-108\trautofit1\trpaddfl3\trpaddl108\trpaddfr3\trpaddr108\tbllkhdrcols\tbllkhdrrows\tbllknocolband\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx3810\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx7710\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx11610\pard\plain\ql\intbl\yts5{\b\f1\fs22\cf0 Job Number: All}\b\f1\fs22\cell\pard\plain\qc\intbl\yts5{\b\f1\fs22\cf0 Department: " + departmentComboBox2.Text + @"}\b\f1\fs22\cell\pard\plain\qr\intbl\yts5{\b\f1\fs22\cf0 Component: All}\b\f1\fs22\cell\trowd\irow0\irowband-1\lastrow\ts5\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\trleft-108\trautofit1\trpaddfl3\trpaddl108\trpaddfr3\trpaddr108\tbllkhdrcols\tbllkhdrrows\tbllknocolband\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx3810\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx7710\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx11610\row\pard\plain\ql\b\f1\fs22\par}";
            ////gridView1.OptionsPrint.RtfPageHeader = richEditControl1.RtfText;
            //gridView1.OptionsPrint.RtfPageFooter = @"{\rtf1\ansi {\fonttbl\f0\ Microsoft Sans Serif;} \f0\pard \fs18 \qr \b Report Date: " + footerDateTime + @"\b0 \par}";

            for (int i = 0; i < PrintDeptsCheckedComboBoxEdit.Properties.Items.Count; i++)
            {
                if (PrintDeptsCheckedComboBoxEdit.Properties.Items[i].CheckState == CheckState.Checked)
                {
                    Console.WriteLine(PrintDeptsCheckedComboBoxEdit.Properties.Items[i].Value + " " + i);
                    departmentComboBox2.SelectedIndex = i;
                    
                    // Used richEditControl to generate desired Rtf code.
                    gridView1.OptionsPrint.RtfPageHeader = @"{\rtf1\deff0{\fonttbl{\f0 Calibri;}{\f1 Microsoft Sans Serif;}}{\colortbl ;\red0\green0\blue255 ;}{\*\defchp \b\f1\fs22}{\stylesheet {\ql\b\f1\fs22 Normal;}{\*\cs1\b\f1\fs22 Default Paragraph Font;}{\*\cs2\sbasedon1\b\f1\fs22 Line Number;}{\*\cs3\b\ul\f1\fs22\cf1 Hyperlink;}{\*\ts4\tsrowd\b\f1\fs22\ql\tscellpaddfl3\tscellpaddl108\tscellpaddfb3\tscellpaddfr3\tscellpaddr108\tscellpaddft3\tsvertalt\cltxlrtb Normal Table;}{\*\ts5\tsrowd\sbasedon4\b\f1\fs22\ql\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\tscellpaddfl3\tscellpaddl108\tscellpaddfr3\tscellpaddr108\tsvertalt\cltxlrtb Table Simple 1;}}{\*\listoverridetable}{\info{\creatim\yr2018\mo1\dy10\hr10\min20}{\version1}}\nouicompat\splytwnine\htmautsp\sectd\trowd\irow0\irowband-1\lastrow\ts5\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\trleft-108\trautofit1\trpaddfl3\trpaddl108\trpaddfr3\trpaddr108\tbllkhdrcols\tbllkhdrrows\tbllknocolband\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx3810\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx7710\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx11610\pard\plain\ql\intbl\yts5{\b\f1\fs22\cf0 Job Number: All}\b\f1\fs22\cell\pard\plain\qc\intbl\yts5{\b\f1\fs22\cf0 Department: " + departmentComboBox2.Text + @"}\b\f1\fs22\cell\pard\plain\qr\intbl\yts5{\b\f1\fs22\cf0 Component: All}\b\f1\fs22\cell\trowd\irow0\irowband-1\lastrow\ts5\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\trleft-108\trautofit1\trpaddfl3\trpaddl108\trpaddfr3\trpaddr108\tbllkhdrcols\tbllkhdrrows\tbllknocolband\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx3810\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx7710\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx11610\row\pard\plain\ql\b\f1\fs22\par}";
                    //gridView1.OptionsPrint.RtfPageHeader = richEditControl1.RtfText;
                    gridView1.OptionsPrint.RtfPageFooter = @"{\rtf1\ansi {\fonttbl\f0\ Microsoft Sans Serif;} \f0\pard \fs18 \qr \b Report Date: " + footerDateTime + @"\b0 \par}";
                    gridView1.OptionsPrint.AutoWidth = false;
                    //gridView1.GridControl.ShowPrintPreview();
                    gridView1.GridControl.Print();
                }
            }

            // Print the gridView control.
            //gridView1.GridControl.Print();

            // Open the Preview window.
            //gridView1.ShowPrintPreview();
        }

        private void printEmployeeWorkButton_Click(object sender, EventArgs e)
        {
            // Check whether the GridControl can be previewed.
            if (!gridControl1.IsPrintingAvailable)
            {
                MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                return;
            }

            departmentComboBox2.SelectedItem = "All";

            gridView1.Columns["Personnel"].OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
            gridView1.Columns["TaskName"].OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.True;

            gridView1.SortInfo.ClearAndAddRange(new[] {new GridColumnSortInfo(colStartDate, DevExpress.Data.ColumnSortOrder.Ascending)});

            for (int i = 0; i < PrintEmployeeWorkCheckedComboBoxEdit.Properties.Items.Count; i++)
            {
                if (PrintEmployeeWorkCheckedComboBoxEdit.Properties.Items[i].CheckState == CheckState.Checked)
                {
                    FilterTaskView2(PrintEmployeeWorkCheckedComboBoxEdit.Properties.Items[i].Value.ToString());

                    var filteredRows = gridView1.DataController.GetAllFilteredAndSortedRows();

                    var count = filteredRows.Count;

                    if (count > 0)
                    {
                        Console.WriteLine(PrintEmployeeWorkCheckedComboBoxEdit.Properties.Items[i].Value + " " + i);
                        // Used richEditControl to generate desired Rtf code.
                        gridView1.OptionsPrint.RtfPageHeader = @"{\rtf1\deff0{\fonttbl{\f0 Calibri;}{\f1 Microsoft Sans Serif;}}{\colortbl ;\red0\green0\blue255 ;}{\*\defchp \b\f1\fs22}{\stylesheet {\ql\b\f1\fs22 Normal;}{\*\cs1\b\f1\fs22 Default Paragraph Font;}{\*\cs2\sbasedon1\b\f1\fs22 Line Number;}{\*\cs3\b\ul\f1\fs22\cf1 Hyperlink;}{\*\ts4\tsrowd\b\f1\fs22\ql\tscellpaddfl3\tscellpaddl108\tscellpaddfb3\tscellpaddfr3\tscellpaddr108\tscellpaddft3\tsvertalt\cltxlrtb Normal Table;}{\*\ts5\tsrowd\sbasedon4\b\f1\fs22\ql\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\tscellpaddfl3\tscellpaddl108\tscellpaddfr3\tscellpaddr108\tsvertalt\cltxlrtb Table Simple 1;}}{\*\listoverridetable}{\info{\creatim\yr2018\mo1\dy10\hr10\min20}{\version1}}\nouicompat\splytwnine\htmautsp\sectd\trowd\irow0\irowband-1\lastrow\ts5\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\trleft-108\trautofit1\trpaddfl3\trpaddl108\trpaddfr3\trpaddr108\tbllkhdrcols\tbllkhdrrows\tbllknocolband\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx3810\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx7710\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx11610\pard\plain\ql\intbl\yts5{\b\f1\fs22\cf0 Job Number: All}\b\f1\fs22\cell\pard\plain\qc\intbl\yts5{\b\f1\fs22\cf0 Personnel: " + PrintEmployeeWorkCheckedComboBoxEdit.Properties.Items[i].Value + @"}\b\f1\fs22\cell\pard\plain\qr\intbl\yts5{\b\f1\fs22\cf0 Component: All}\b\f1\fs22\cell\trowd\irow0\irowband-1\lastrow\ts5\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\trleft-108\trautofit1\trpaddfl3\trpaddl108\trpaddfr3\trpaddr108\tbllkhdrcols\tbllkhdrrows\tbllknocolband\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx3810\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx7710\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx11610\row\pard\plain\ql\b\f1\fs22\par}";
                        //gridView1.OptionsPrint.RtfPageHeader = richEditControl1.RtfText;
                        gridView1.OptionsPrint.RtfPageFooter = @"{\rtf1\ansi {\fonttbl\f0\ Microsoft Sans Serif;} \f0\pard \fs18 \qr \b Report Date: " + footerDateTime + @"\b0 \par}";
                        gridView1.OptionsPrint.AutoWidth = false;
                        //gridView1.GridControl.ShowPrintPreview();
                        gridView1.GridControl.Print();
                    }
                }
            }

            gridView1.ActiveFilterCriteria = FilterTaskView(departmentComboBox2.Text, false, false, filterTasksByDatesCheckEdit.Checked);
        }
        private void taskViewExportButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel File (*.xlsx) |*.xlsx";
            saveFileDialog.InitialDirectory = @"C:\Users\" + Environment.UserName + @"\Desktop";
            saveFileDialog.FileName = "Tool Room Tasks " + DateTime.Today.Month + "-" + DateTime.Today.Day + "-" + DateTime.Today.Year;
            saveFileDialog.DefaultExt = "xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                gridView1.ExportToXlsx(saveFileDialog.FileName);
            }
        }
        private void gridControl1_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                string clickInfo = "";
                GridControl grid = sender as GridControl;
                if (grid == null) return;
                // Get a View at the current point.
                BaseView view = grid.GetViewAt(e.Location);
                if (view == null) return;
                // Retrieve information on the current View element.
                BaseHitInfo baseHI = view.CalcHitInfo(e.Location);
                GridHitInfo gridHI = baseHI as GridHitInfo;
                if (gridHI != null)
                    clickInfo = gridHI.HitTest.ToString();

                //MessageBox.Show(clickInfo);

                if (clickInfo == "RowCell" && gridHI.Column.ToString() == "Component")
                {
                    OpenKanBanWorkbook(gridHI.RowHandle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }
        private void gridView1_PrintInitialize(object sender, PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;


            //pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.Letter;

            pb.PageSettings.TopMargin = 20;
            pb.PageSettings.BottomMargin = 20;
            pb.PageSettings.LeftMargin = 25;
            pb.PageSettings.RightMargin = 25;
            pb.Document.AutoFitToPagesWidth = 1;
            pb.PageSettings.Landscape = true;
        }
        private void gridView1_ValidatingEditor(object sender, BaseContainerValidateEditorEventArgs e)
        {
            ColumnView view = sender as ColumnView;

            GridColumn column = (e as EditFormValidateEditorEventArgs)?.Column ?? view.FocusedColumn;

            if (column.FieldName == "StartDate" || column.FieldName == "FinishDate")
            {
                if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeDates))
                {
                    MessageBox.Show("This login is not authorized to make changes to dates.  Hit ESC to cancel editing.");
                    e.Valid = false;
                } 
            }
        }
        private void gridView1_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            try
            {
                GridView view = sender as GridView;
                TaskModel task = view.GetFocusedRow() as TaskModel;
                ComponentModel component = ComponentsList.Find(x => x.Component == task.Component && x.ProjectNumber == task.ProjectNumber);

                gridView1.BeginUpdate();
                gridView5.BeginUpdate();
                schedulerControl1.BeginUpdate();

                deptTaskViewHelper = new RefreshHelper(gridView1, "ProjectNumber");

                if (e.Column.FieldName == "StartDate" || e.Column.FieldName == "FinishDate")
                {
                    component.ChangeTaskDate(e.Column.FieldName, task);
                }
                else if (e.Column.FieldName == "Machine" || e.Column.FieldName == "Personnel")
                {
                    task.Resources = GeneralOperations.GenerateResourceIDsString(schedulerDataStorage1, task.Machine, task.Personnel);
                    Database.UpdateTask(task, e);
                }
                else
                {
                    Database.UpdateTask(task, e);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());

                deptTaskViewHelper.SaveViewInfo();
                LoadTaskView();
                deptTaskViewHelper.LoadViewInfo();
            }
            finally
            {
                gridView5.EndUpdate();
                gridView1.EndUpdate();
                schedulerControl1.EndUpdate();
                schedulerControl1.RefreshData();
            }

        }

        private void gridView1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("gridView1_Click");
        }

        private void gridView1_CustomUnboundColumnData(object sender, CustomColumnDataEventArgs e)
        {
            GridView view = sender as GridView;

            ProjectModel pi = ProjectsList.Find(x => x.ProjectNumber == (int)view.GetListSourceRowCellValue(e.ListSourceRowIndex, "ProjectNumber"));

            if (e.IsGetData)
            {
                if (e.Column.FieldName == "DueDate")
                {
                    e.Value = pi.DueDate;
                }
                else if (e.Column.FieldName == "ProjectStatus")
                {
                    e.Value = pi.Status;
                }
                else if (e.Column.FieldName == "ToolMaker2")
                {
                    e.Value = pi.ToolMaker;
                }
            }
        }

        private void repositoryItemCheckedComboBoxEdit1_QueryPopUp(object sender, CancelEventArgs e)
        {
            string task = (string)gridView1.GetRowCellValue(gridView1.GetSelectedRows()[0], "TaskName");
            string role = "";

            if (task == "CNC Rough")
            {
                role = "Rough Mill";
            }
            else if (task == "CNC Finish")
            {
                role = "Finish Mill";
            }
            else if(task == "CNC Electrodes")
            {
                role = "Graphite Mill";
            }
            else if(task == "EDM Sinker")
            {
                role = "EDM Sinker";
            }
            else if (task == "EDM Wire (In-House)")
            {
                role = "EDM Wire";
            }
            else
            {
                e.Cancel = true;
            }

            if (role != "")
            {
                repositoryItemCheckedComboBoxEdit1.DataSource = GetResourceList(role, "Machine");
            }

        }

        private void gridView1_ShownEditor(object sender, EventArgs e)
        {
            ComboBoxEdit comboBoxEdit = null;

            if (gridView1.ActiveEditor.EditorTypeName == "ComboBoxEdit")
            {
                comboBoxEdit = gridView1.ActiveEditor as ComboBoxEdit;
            }

            if (comboBoxEdit != null && gridView1.FocusedColumn.FieldName == "Personnel")
            {
                string task = (string)gridView1.GetRowCellValue(gridView1.GetSelectedRows()[0], "TaskName");
                string role = "";

                comboBoxEdit.Properties.Items.Clear();

                if (task == "Program Rough")
                {
                    role = "Rough Programmer";
                }
                else if (task == "Program Finish")
                {
                    role = "Finish Programmer";
                }
                else if (task == "Program Electrodes")
                {
                    role = "Electrode Programmer";
                }
                else if (task.EndsWith("Grind") || task == "Polish")
                {
                    role = "Tool Maker";
                }
                else if (task == "CNC Rough")
                {
                    role = "Rough CNC Operator";
                }
                else if (task == "CNC Finish")
                {
                    role = "Finish CNC Operator";
                }
                else if( task == "CNC Electrodes")
                {
                    role = "Electrode CNC Operator";
                }
                else if (task == "EDM Wire (In-House)")
                {
                    role = "EDM Wire Operator";
                }
                else if (task == "EDM Sinker")
                {
                    role = "EDM Sinker Operator";
                }
                else if (task == "Hole Pop")
                {
                    role = "Hole Popper Operator";
                }
                else if (task.StartsWith("Inspection"))
                {
                    role = "CMM Operator";
                }

                comboBoxEdit.Properties.Items.Clear();
                comboBoxEdit.Properties.Items.AddRange(GetResourceList(role, "Person").ToArray());
            }
        }

        #endregion

        #region Project View

        private class ExpandedProjectRows
        {
            public int RowHandle { get; set; }
            public List<int> ExpandedComponentRows { get; set; }
            public List<int> SelectedComponentRows { get; set; }
            public int FocusedComponentRow { get; set; }
            public List<int> SelectedTaskRows { get; set; }
            public int FocusedTaskRow { get; set; }

            public ExpandedProjectRows()
            {
                ExpandedComponentRows = new List<int>();
                SelectedComponentRows = new List<int>();
            }

            public void AddComponentRow(int rowHandle)
            {
                ExpandedComponentRows.Add(rowHandle);
            }

            public void RemoveComponentRow(int rowHandle)
            {
                ExpandedComponentRows.Remove(rowHandle);
            }
        }

        private void ExpandStoredRows()
        {
            this.gridView3.MasterRowExpanded -= this.gridView_MasterRowExpanded;
            this.gridView4.MasterRowExpanded -= this.gridView_MasterRowExpanded;
            //this.gridView3.MasterRowExpanded -= new DevExpress.XtraGrid.Views.Grid.CustomMasterRowEventHandler(this.gridView_MasterRowExpanded);
            //this.gridView3.MasterRowCollapsed -= new DevExpress.XtraGrid.Views.Grid.CustomMasterRowEventHandler(this.gridView_MasterRowCollapsed);

            foreach (int row in gridView3ExpandedRowsList)
            {
                Console.WriteLine(row);
                gridView3.SetMasterRowExpanded(row, true);
            }

            foreach (int row in gridView4ExpandedRowsList)
            {
                Console.WriteLine(row);
                gridView4.SetRowExpanded(row, true);
            }

            this.gridView3.MasterRowExpanded += this.gridView_MasterRowExpanded;
            this.gridView4.MasterRowExpanded += this.gridView_MasterRowExpanded;
        }

        private void DetermineExpandedRows()
        {
            var count = gridView3.RowCount;
            expandedProjectRowsList.Clear();
            gridView4SelectedRows.Clear();
            ExpandedProjectRows epr;

            for (int projectRow = 0; projectRow < gridView3.RowCount; projectRow++)
            {
                if (gridView3.GetMasterRowExpanded(projectRow))
                {
                    epr = new ExpandedProjectRows();
                    epr.RowHandle = projectRow;
                    
                    //Console.WriteLine(projectRow + " is expanded.");

                    var childView = gridView3.GetVisibleDetailView(projectRow) as GridView;
                    
                    for (int componentRow = 0; componentRow < childView.DataRowCount; componentRow++)
                    {
                        if (childView.GetMasterRowExpanded(componentRow) == true)
                        {
                            epr.AddComponentRow(componentRow);
                            //Console.WriteLine(componentRow + " is expanded."); 
                        }

                        if (childView.IsRowSelected(componentRow))
                        {
                            epr.SelectedComponentRows.Add(componentRow);
                        }
                    }

                    epr.FocusedComponentRow = childView.FocusedRowHandle;
                    expandedProjectRowsList.Add(epr);
                }
                else
                {
                    Console.WriteLine(projectRow + " is collapsed.");
                }
            }
        }

        private void GetSelectedRows()
        {
            //gridView3.SelectedRowsCount;
            gridView3SelectedRows = gridView3.GetSelectedRows().ToList();
            gridView4SelectedRows = gridView4.GetSelectedRows().ToList();
            gridView5SelectedRows = gridView5.GetSelectedRows().ToList();
        }

        private void SelectRows()
        {
            foreach (int row in gridView3SelectedRows)
            {
                gridView3.FocusedRowHandle = row;
                gridControl3.Focus();
            }

            //foreach (int row in gridView4SelectedRows)
            //{
            //    gridView4.FocusedRowHandle = row;
            //    gridControl3.Focus();
            //}

            //foreach (int row in gridView5SelectedRows)
            //{
            //    gridView5.FocusedRowHandle = row;
            //    gridControl3.Focus();
            //}
        }

        private void RecursiveExpand()
        {
            foreach (ExpandedProjectRows projectRow in expandedProjectRowsList)
            {
                gridView3.SetMasterRowExpanded(projectRow.RowHandle, true);

                var childView = gridView3.GetVisibleDetailView(projectRow.RowHandle) as GridView;

                foreach (int componentRow in projectRow.ExpandedComponentRows)
                {
                    childView.SetMasterRowExpanded(componentRow, true);
                }

                foreach (int row in projectRow.SelectedComponentRows)
                {
                    childView.SelectRow(row);
                    childView.Focus();
                }

                childView.FocusedRowHandle = projectRow.FocusedComponentRow;
            }
        }

        private void RefreshProjectGrid()
        {
            GetSelectedRows();
            DetermineExpandedRows();
            ColorList = Database.GetColorEntries();
            LoadProjects();
            LoadProjectView();
            CollapseGroups();
            RecursiveExpand();
            SelectRows();
        }

        private bool KanBanExists(string kanBanWorkbookPath)
        {
            if (kanBanWorkbookPath != "")
            {
                FileInfo fi = new FileInfo(kanBanWorkbookPath);

                if (fi.Exists)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private void CreateProject()
        {
            using (var form = new ProjectCreationForm(schedulerDataStorage1))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    if (form.DataValidated)
                    {
                        RefreshProjectGrid();
                        int rowHandle = gridView3.LocateByValue("ProjectNumber", form.Project.ProjectNumber);
                        if (rowHandle != GridControl.InvalidRowHandle)
                            gridView3.FocusedRowHandle = rowHandle;
                        gridView3.SetMasterRowExpanded(gridView3.FocusedRowHandle, true);
                    }
                }
                else if (result == DialogResult.Cancel)
                {

                }
            }
        }
        private void EditProject(ProjectModel project)
        {
            using (var form = new ProjectCreationForm(project, schedulerDataStorage1))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    if (form.DataValidated)
                    {
                        if (project.KanBanWorkbookPath.ToString().Length > 0)
                        {
                            MessageBox.Show("A project has changed.  Need to regenerate and reprint Kan Ban.");
                            //gridView3.Appearance.FocusedRow.BackColor = Color.Red;
                        }

                        RefreshProjectGrid();
                    }
                }
                else if (result == DialogResult.Cancel)
                {

                }
            }
        }
        private void DeleteProject(object sender, KeyEventArgs e)
        {
            if (MessageBox.Show("Delete Project?", "Confirmation", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;

            GridView view = sender as GridView;
            ProjectModel project = view.GetFocusedRow() as ProjectModel;

            try
            {
                PreserveNotes();

                if (Database.RemoveProject(project))
                {
                    DeletedProjects.Add(project);

                    view.DeleteRow(view.FocusedRowHandle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }
        private GridView GetFocusedView()
        {
            if (gridView3.IsFocusedView)
            {
                return gridView3;
            }
            else if (projectBandedGridView.IsFocusedView)
            {
                return projectBandedGridView;
            }

            return null;
        }
        private bool IsPastDate(DateTime? target, DateTime? actual)
        {
            if (target == null || actual == null)
            {
                return false;
            }
            else if (actual > target)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private void CollapseGroups()
        {
            BandedGridView view = projectBandedGridView;

            int count = 0;
            for (int i = 0; i < view.RowCount; i++)
            {
                int rowHandle = view.GetVisibleRowHandle(i);
                if (view.IsGroupRow(rowHandle))
                {
                    count++;

                    // Using GetGroupRowValue causes an error to occur when it is null.
                    if (view.GetGroupRowDisplayText(rowHandle).Contains("On Hold") || view.GetGroupRowDisplayText(rowHandle).Contains("Completed") || view.GetGroupRowDisplayText(rowHandle).Contains("Outsourced") || view.GetGroupRowDisplayText(rowHandle).Contains("Closed"))
                    {
                        view.CollapseGroupRow(rowHandle);

                        //MessageBox.Show(view.GetGroupRowDisplayText(rowHandle));
                    }
                }
            }
        }
        private void PreserveNotes()
        {
            if (MessageBox.Show("Do you want to generate / update the Kan Ban for this project to preserve notes?", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                GenerateKanBan();
            }
        }
        private void GenerateKanBan()
        {
            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                GridView gridView = gridControl3.MainView as GridView;

                if (gridView.SelectedRowsCount != 1)
                {
                    MessageBox.Show("Please select a project.");
                    return;
                }
                else
                {
                    ProjectModel project = gridView.GetFocusedRow() as ProjectModel;
                    string path;

                    project = Database.GetProject(project.ProjectNumber);

                    if (KanBanExists(project.KanBanWorkbookPath))
                    {
                        DialogResult result = XtraMessageBox.Show("A Kan Ban for this project already exists.\n\nDo you want to create a new one?\n\n" +
                            "(Click 'Yes' to create new one (All info preserved).  Click 'No' to cancel.)", "Warning",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                        if (result == DialogResult.Yes)
                        {
                            path = ExcelInteractions.GenerateKanBanWorkbook2(project);

                            if (path != "")
                            {
                                Database.SetKanBanWorkbookPath(path, project.ProjectNumber);
                            }
                        }
                        else if (result == DialogResult.No)
                        {
                            // This space was formerly occupied by the 'EditKanBan' method.

                            return;
                        }
                        else if (result == DialogResult.Cancel)
                        {
                            return;
                        }
                    }
                    else
                    {
                        path = ExcelInteractions.GenerateKanBanWorkbook2(project);

                        if (path != "")
                        {
                            Database.SetKanBanWorkbookPath(path, project.ProjectNumber);
                        }
                    }

                    RefreshProjectGrid();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                SplashScreenManager.CloseForm();
            }
        }
        private void GenerateKanBanReport()
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));

            GridView gridView = gridControl3.MainView as GridView;

            ProjectModel project = gridView.GetFocusedRow() as ProjectModel;

            project = Database.GetProject(project.ProjectNumber);

            KanBanXtraReport kanBan = new KanBanXtraReport(project);

            ReportPrintTool printTool = new ReportPrintTool(kanBan);

            printTool.Report.CreateDocument(true);
            printTool.ShowPreviewDialog();

            SplashScreenManager.CloseForm();
        }
        private Color GetRowColor(int rowHandle, BandedGridView bandedGridView)
        {
            if (rowHandle % 2 == 0)
            {
                return bandedGridView.Appearance.OddRow.BackColor;
            }
            else
            {
                return bandedGridView.Appearance.EvenRow.BackColor;
            }
        }
        private Color? GetColorFromUser(string columnType, Color rowColor)
        {
            using (var ssw = new SelectStatusWindow(columnType, rowColor))
            {
                Point windowLocation = new Point(Cursor.Position.X - (int)(ssw.Width * .3), Cursor.Position.Y - (int)(ssw.Height * .5));

                ssw.Location = windowLocation;

                var result = ssw.ShowDialog();

                if (result == DialogResult.Cancel)
                {
                    ssw.Close();
                }
                else if (result == DialogResult.OK)
                {
                    return ssw.SelectedColor;
                }

                return null;
            }
        }

        private void SetSelectedCellColor(Color? color, GridCell[] cells, BandedGridView bandedGridView)
        {
            if (color == null)
            {
                return;
            }

            ColorStruct colorItem;
            ProjectModel project;
            Color rowColor;

            //ColorList.Clear();

            foreach (var cell in cells)
            {
                project = bandedGridView.GetRow(cell.RowHandle) as ProjectModel;

                rowColor = GetRowColor(cell.RowHandle, bandedGridView);

                colorItem = ColorList.Find(r => r.ColumnFieldName == cell.Column.FieldName && r.ProjectID == project.ID); // Somehow the same color was added twice for the same role for the same project.

                if (colorItem == null && rowColor != ((Color)color))
                {
                    colorItem = new ColorStruct { ProjectID = project.ID, ColumnFieldName = cell.Column.FieldName, ARGBColor = ((Color)color).ToArgb() };

                    if (Database.AddColorEntry(colorItem))
                    {
                        ColorList.Add(colorItem); 
                    }
                }
                else if (colorItem != null && rowColor == ((Color)color))
                {
                    ColorList.Remove(colorItem);

                    Database.DeleteColorEntry(colorItem);
                }
                else
                {
                    colorItem.ARGBColor = ((Color)color).ToArgb();

                    Database.UpdateColorEntry(colorItem);
                }
            }

            bandedGridView.LayoutChanged();
        }

        private void AddRepositoryItemToGrid()
        {
            RichTextBox richTextBox = new RichTextBox();
            richTextBox.Dock = DockStyle.Fill;

            SimpleButton fontButton = new SimpleButton();
            fontButton.Appearance.Font = new Font(fontButton.Font.FontFamily, fontButton.Font.Size, FontStyle.Regular);
            fontButton.Text = "Font";
            fontButton.Left = 40;
            fontButton.Top = 20;
            fontButton.Width = 50;
            fontButton.Height = 20;
            fontButton.Click += new EventHandler(fontButton_Clicked);

            //ColorEdit textColorPicker = new ColorEdit();
            //textColorPicker.Left = 50;
            //textColorPicker.Top = 40;
            //textColorPicker.Width = 50;
            //textColorPicker.Height = 25;
            //textColorPicker.Dock = DockStyle.None;
            //textColorPicker.Color = Color.Black;
            //textColorPicker.ColorChanged += new EventHandler(editorColorPickerControl_ColorChanged);

            //SimpleButton boldButton = new SimpleButton();
            //boldButton.Appearance.Font = new Font(boldButton.Font.FontFamily, boldButton.Font.Size, FontStyle.Bold);
            //boldButton.Text = "B";
            //boldButton.Left = 105;
            //boldButton.Top = 40;
            //boldButton.Width = 20;
            //boldButton.Height = 20;
            //boldButton.Click += new EventHandler(boldButton_Clicked);

            //SimpleButton underlineButton = new SimpleButton();
            //underlineButton.Appearance.Font = new Font(underlineButton.Font.FontFamily, underlineButton.Font.Size, FontStyle.Underline);
            //underlineButton.Text = "U";
            //underlineButton.Left = 130;
            //underlineButton.Top = 40;
            //underlineButton.Width = 20;
            //underlineButton.Height = 20;
            //underlineButton.Click += new EventHandler(underlineButton_Clicked);

            //SimpleButton plainButton = new SimpleButton();
            //plainButton.Appearance.Font = new Font(plainButton.Font.FontFamily, plainButton.Font.Size, FontStyle.Regular);
            //plainButton.Text = "P";
            //plainButton.Left = 155;
            //plainButton.Top = 40;
            //plainButton.Width = 20;
            //plainButton.Height = 20;
            //plainButton.Click += new EventHandler(plainButton_Clicked);

            SimpleButton editorOKButton = new SimpleButton();
            editorOKButton.Text = "OK";
            editorOKButton.Left = 40;
            editorOKButton.Top = 50;
            editorOKButton.Width = 50;
            editorOKButton.Height = 30;
            editorOKButton.Dock = DockStyle.None;
            editorOKButton.Click += new EventHandler(editorOKButton_Clicked);

            SimpleButton editorCancelButton = new SimpleButton();
            editorCancelButton.Text = "Cancel";
            editorCancelButton.Left = 110;
            editorCancelButton.Top = 50;
            editorCancelButton.Width = 50;
            editorCancelButton.Height = 30;
            editorCancelButton.Dock = DockStyle.None;
            editorCancelButton.Click += new EventHandler(editorCancelButton_Clicked);

            Panel panel = new Panel();
            panel.Height = 80;
            panel.Dock = DockStyle.Bottom;
            panel.Controls.Add(editorOKButton);
            panel.Controls.Add(editorCancelButton);
            panel.Controls.Add(fontButton);

            //panel.Controls.Add(textColorPicker);
            //panel.Controls.Add(boldButton);
            //panel.Controls.Add(underlineButton);
            //panel.Controls.Add(plainButton);

            RepositoryItemRichTextEdit repositoryItemRichTextEdit = new RepositoryItemRichTextEdit();

            PopupContainerControl popupContainerControl = new PopupContainerControl();
            popupContainerControl.Controls.Add(richTextBox);
            popupContainerControl.Controls.Add(panel);
            popupContainerControl.Height = 200;

            PopupContainerEdit popupContainerEdit = new PopupContainerEdit();
            popupContainerEdit.Properties.PopupControl = popupContainerControl;

            // The initialization of this instance of repositoryItemPopupContainer edit is at the top of this class.
            repositoryItemPopupContainerEdit.PopupControl = popupContainerControl;

            gridControl3.RepositoryItems.Add(repositoryItemPopupContainerEdit);
            projectBandedGridView.Columns["GeneralNotes"].ColumnEdit = repositoryItemRichTextEdit;
        }
        private void PopulateShownEditor(object sender)
        {
            GridView gridView = sender as GridView;

            PopupContainerEdit popupContainerEdit = null;
            ComboBoxEdit comboBoxEdit = null;

            if (gridView != null)
            {

                if (gridView.ActiveEditor.EditorTypeName == "PopupContainerEdit")
                {
                    popupContainerEdit = gridView.ActiveEditor as PopupContainerEdit;
                }
                else if (gridView.ActiveEditor.EditorTypeName == "ComboBoxEdit")
                {
                    comboBoxEdit = gridView.ActiveEditor as ComboBoxEdit;
                }

                if (popupContainerEdit != null)
                {
                    RichTextBox richTextBox = (RichTextBox)popupContainerEdit.Properties.PopupControl.Controls[0];

                    richTextBox.Rtf = gridView.GetFocusedRowCellValue("GeneralNotes").ToString();
                }

                if (comboBoxEdit != null)
                {
                    string column = gridView.FocusedColumn.FieldName;
                    string role = "";

                    if (column == "RoughProgrammer")
                    {
                        role = "Rough Programmer";
                    }
                    else if (column == "FinishProgrammer")
                    {
                        role = "Finish Programmer";
                    }
                    else if (column == "ElectrodeProgrammer")
                    {
                        role = "Electrode Programmer";
                    }
                    else if (column == "ToolMaker")
                    {
                        role = "Tool Maker";
                    }
                    else if (column == "Engineer")
                    {
                        role = "Engineer";
                    }
                    else if (column == "Designer")
                    {
                        role = "Designer";
                    }
                    else if (column == "Apprentice")
                    {
                        role = "Apprentice";
                    }
                    else if (column == "Stage")
                    {
                        return;
                    }

                    if (role != "")
                    {
                        comboBoxEdit.Properties.Items.Clear();
                        comboBoxEdit.Properties.Items.AddRange(GetResourceList(role, "Person"));
                    }
                }
            }
        }
        private void editorOKButton_Clicked(object sender, EventArgs e)
        {
            Console.WriteLine("okButton_Clicked");

            BandedGridView bandedGridView = GetFocusedView() as BandedGridView;

            PopupContainerEdit popupContainerEdit = bandedGridView.ActiveEditor as PopupContainerEdit;
            RichTextBox richTextBox = popupContainerEdit.Properties.PopupControl.Controls[0] as RichTextBox;


            popupContainerEdit.EditValue = richTextBox.Rtf;
            popupContainerEdit.ClosePopup();

            //Control button = sender as Control;
            ////Close the dropdown accepting the user's choice 
            //(button.Parent.Parent as PopupContainerControl).OwnerEdit.ClosePopup();
        }

        private void editorCancelButton_Clicked(object sender, EventArgs e)
        {
            BandedGridView bandedGridView = GetFocusedView() as BandedGridView;

            PopupContainerEdit popupContainerEdit = bandedGridView.ActiveEditor as PopupContainerEdit;

            popupContainerEdit.CancelPopup();
        }

        private void fontButton_Clicked(object sender, EventArgs e)
        {
            BandedGridView bandedGridView = GetFocusedView() as BandedGridView;

            PopupContainerEdit popupContainerEdit = bandedGridView.ActiveEditor as PopupContainerEdit;

            RichTextBox richTextBox = (RichTextBox)popupContainerEdit.Properties.PopupControl.Controls[0];

            FontDialog fontDialog = new FontDialog();
            fontDialog.ShowColor = true;

            if (fontDialog.ShowDialog() != DialogResult.Cancel)
            {
                richTextBox.SelectionFont = fontDialog.Font;
                richTextBox.SelectionColor = fontDialog.Color;
            }
        }
        private void gridControl3_Load(object sender, EventArgs e)
        {
            footerDateTime = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
            ColorList = Database.GetColorEntries();
            //CollapseGroups();  Calling CollapseGroups from here doesn't work.  For reason the rows do not yet exist on the grid.
        }
        private CriteriaOperator FilterGridView3()
        {
            List<CriteriaOperator> criteriaOperators = new List<CriteriaOperator>();
            criteriaOperators.Add(new NotOperator(new BinaryOperator("Stage", "7 - Completed")));

            return GroupOperator.And(criteriaOperators);
        }
        private void gridView3_PrintInitialize(object sender, PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;

            pb.PageSettings.TopMargin = 25;
            pb.PageSettings.RightMargin = 25;
            pb.PageSettings.BottomMargin = 25;
            pb.PageSettings.LeftMargin = 25;
            pb.Document.AutoFitToPagesWidth = 1;
        }
        private void projectBandedGridView_PrintInitialize(object sender, PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;

            pb.PageSettings.TopMargin = 25;
            pb.PageSettings.RightMargin = 25;
            pb.PageSettings.BottomMargin = 25;
            pb.PageSettings.LeftMargin = 25;
            pb.Document.AutoFitToPagesWidth = 1;

            if (PaperSize == "Tabloid")
            {
                pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.Tabloid;
                pb.PageSettings.PrinterName = @"\\S-PS1-SMDRV\P-1336 HP CP5225 - Color";
            }
            else if (PaperSize == "Letter")
            {
                pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.Letter;
            }

            if (PrintOrientation == "Landscape")
            {
                pb.PageSettings.Landscape = true;
            }
            else if (PrintOrientation == "Portrait")
            {
                pb.PageSettings.Landscape = false;
            }

            projectBandedGridView.OptionsPrint.RtfPageFooter = @"{\rtf1\ansi {\fonttbl\f0\ Microsoft Sans Serif;} \f0\pard \fs18 \qr \b Report Date: " + footerDateTime + @"\b0 \par}";
        }
        private void workLoadViewPrintPreviewButton_Click(object sender, EventArgs e)
        {
            PrintOrientation = "Landscape";
            PaperSize = "Tabloid";

            GridView gridView = gridControl3.MainView as GridView;

            gridView.ShowPrintPreview();

            //gridControl3.ShowPrintPreview();
        }
        private void projectBandedGridView_ShownEditor(object sender, EventArgs e)
        {
            try
            {
                PopulateShownEditor(sender);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }
        private void projectBandedGridView_CustomColumnSort(object sender, CustomColumnSortEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null) return;

            try
            {
                if (e.Column.FieldName == "Stage")
                {
                    object val1 = view.GetListSourceRowCellValue(e.ListSourceRowIndex1, "StageNumber");
                    object val2 = view.GetListSourceRowCellValue(e.ListSourceRowIndex2, "StageNumber");
                    e.Result = System.Collections.Comparer.Default.Compare(val1, val2);
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }
        private void projectBandedGridView_CustomRowCellEditForEditing(object sender, CustomRowCellEditEventArgs e)
        {
            if (e.Column.FieldName == "GeneralNotes")
            {
                e.RepositoryItem = repositoryItemPopupContainerEdit;
            }
        }
        private void projectBandedGridView_ValidatingEditor(object sender, BaseContainerValidateEditorEventArgs e)
        {
            ColumnView view = sender as ColumnView;

            GridColumn column = (e as EditFormValidateEditorEventArgs)?.Column ?? view.FocusedColumn;

            if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeProjectData) && column.FieldName != "GeneralNotes")
            {
                e.ErrorText = "This login is not authorized to make changes to project level data.  Hit ESC to cancel editing.";
                e.Valid = false;
            }
            else if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeGeneralNotes) && column.FieldName == "GeneralNotes")
            {
                e.ErrorText = "This login is not authorized to make changes to project level data.  Hit ESC to cancel editing.";
                e.Valid = false;
            }
            else if (column.FieldName == "ProjectNumber") // column.FieldName == "MWONumber" || 
            {
                if (e.Value.ToString() != "" && int.TryParse(e.Value.ToString(), out int result) == false)
                {
                    e.ErrorText = "Please enter a number.  Hit ESC to cancel editing.";
                    e.Valid = false;
                }
                else if (Database.ProjectExists(int.Parse(e.Value.ToString())))
                {
                    e.ErrorText = "A project with that Project Number already exists.  Hit ESC to cancel editing.";
                    e.Valid = false;
                }
            }
            else if (column.FieldName == "DeliveryInWeeks")
            {
                if (e.Value.ToString() != "" && double.TryParse(e.Value.ToString(), out double result) == false)
                {
                    e.ErrorText = "Please enter a number.  Hit ESC to cancel editing.";
                    e.Valid = false;
                }
            }
        }
        private void projectBandedGridView_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            BandedGridView view = sender as BandedGridView;
            
            ProjectModel project = view.GetRow(e.RowHandle) as ProjectModel;

            try
            {
                if (e.Column.FieldName == "DeliveryInWeeks" && project.StartDate != null)
                {
                    project.DueDate = Convert.ToDateTime(project.StartDate).AddDays(Convert.ToDouble(e.Value) * 7);
                }
                else if (e.Column.FieldName == "StartDate" && project.DeliveryInWeeks != null)
                {
                    project.DueDate = Convert.ToDateTime(e.Value).AddDays(Convert.ToDouble(project.DeliveryInWeeks) * 7);
                }
                else if (e.Column.FieldName == "Stage")
                {
                    if (e.Value.ToString() == "Completed")
                    {
                        PreserveNotes();
                    }
                }

                if (view.IsNewItemRow(e.RowHandle))
                {
                    return;
                }
                else
                {
                    Console.WriteLine("projectBandedGridView Cell Value Changed Event");
                    //Console.WriteLine("Changed Cell Value: " + e.Value.ToString());

                    if (!Database.UpdateProjectRecord(project, e))
                    {
                        LoadProjects();
                        return;
                    }
                    else
                    {
                        if (e.Column.FieldName == "ProjectNumber" || e.Column.FieldName == "JobNumber" || e.Column.FieldName == "DeliveryInWeeks" || e.Column.FieldName == "StartDate" || e.Column.FieldName == "DueDate")
                        {
                            project.UpdateKanBan = true;
                        }
                    }

                    if (e.Column.FieldName.Contains("Programmer") || e.Column.FieldName == "Designer")
                    {
                        Database.SetTaskPersonnel(project.ProjectNumber, GeneralOperations.FindMatchingDepartment(e.Column.FieldName, Database.GetDepartmentRoles()), e.Value.ToString(), schedulerDataStorage1);
                        RefreshProjectGrid();
                        RefreshDepartmentScheduleView();
                    } 
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }
        private void projectBandedGridView_RowUpdated(object sender, RowObjectEventArgs e)
        {
            GridView view = sender as GridView;

            ProjectModel project = e.Row as ProjectModel;

            try
            {
                if (view.IsNewItemRow(e.RowHandle))
                {
                    if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanCreateProjects))
                    {
                        MessageBox.Show("This login is not authorized to create projects.");
                        return;
                    }

                    Database.CreateProjectEntry(project);

                    RefreshProjectGrid();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                Console.WriteLine(ex.ToString());
            }
        }
        private void projectBandedGridView_ValidateRow(object sender, ValidateRowEventArgs e)
        {
            GridView view = sender as GridView;

            ProjectModel project = e.Row as ProjectModel;

            if (view.IsNewItemRow(e.RowHandle))
            {
                if (project.JobNumber == null || project.JobNumber.Length == 0)
                {
                    view.SetColumnError(colJobNumberBGV, "Please enter a Job Number");
                    e.Valid = false;
                }
                else if (project.ProjectNumber == 0)
                {
                    view.SetColumnError(colProjectNumberBGV, "Please enter a Project Number");
                    e.Valid = false;
                }
                else if (project.DueDate == new DateTime(0001, 1, 1))
                {
                    view.SetColumnError(colDueDateBGV, "Please enter a due date.");
                    e.Valid = false;
                }
            }
        }
        private void projectBandedGridView_InvalidRowException(object sender, InvalidRowExceptionEventArgs e)
        {
            //Suppress displaying the error message box
            e.ExceptionMode = ExceptionMode.NoAction;
        }
        private void projectBandedGridView_MouseDown(object sender, MouseEventArgs e)
        {
            Console.WriteLine("bandedGridView1 Mouse down event.");
            BandedGridView bandedGridView = sender as BandedGridView;
            List<string> PersonnelColumns = new List<string> { "Engineer", "Designer", "ToolMaker", "RoughProgrammer", "FinishProgrammer", "ElectrodeProgrammer" };
            List<string> OtherColumns = new List<string> { "AdjustDeliveryDate", "StartDate", "FinishDate", "GeneralNotes" };
            var hitInfo = bandedGridView.CalcHitInfo(e.Location);
            Color? color;
            Color rowColor;

            if (hitInfo.InRowCell)
            {
                int rowHandle = hitInfo.RowHandle;
                GridColumn column = hitInfo.Column;

                if (e.Button == MouseButtons.Right)
                {
                    if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeProjectData))
                    {
                        MessageBox.Show("This login is not authorized to make changes to work load tab.");
                        return;
                    }

                    var cells = bandedGridView.GetSelectedCells();

                    foreach (var cell in cells)
                    {
                        bandedGridView.UnselectCell(cell);
                    }

                    bandedGridView.SelectCell(rowHandle, column);

                    rowColor = GetRowColor(rowHandle, bandedGridView);

                    if (PersonnelColumns.Exists(x => x == column.FieldName))
                    {
                        color = GetColorFromUser("Personnel", rowColor);
                        
                        cells = bandedGridView.GetSelectedCells();

                        SetSelectedCellColor(color, cells, bandedGridView);
                    }
                    else if (OtherColumns.Exists(x => x == column.FieldName))
                    {
                        color = GetColorFromUser("Other", rowColor);

                        cells = bandedGridView.GetSelectedCells();

                        SetSelectedCellColor(color, cells, bandedGridView);
                    }
                    else if (column.FieldName == "JobFolderPath")
                    {
                        FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                        folderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer;
                        folderBrowserDialog.SelectedPath = @"X:\TOOLROOM\";
                        var result = folderBrowserDialog.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            Database.SetJobFolderPath((int)bandedGridView.GetRowCellValue(rowHandle, "ID"), folderBrowserDialog.SelectedPath);

                        }
                    }
                }
                else if (e.Button == MouseButtons.Left)
                {


                }

            }
        }
        private void projectBandedGridView_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            BandedGridView bandedGridView = sender as BandedGridView;
            ProjectModel project = bandedGridView.GetRow(e.RowHandle) as ProjectModel;

            if (e.RowHandle >= 0)
            {
                var data = ColorList.FirstOrDefault(p => p.ColumnFieldName == e.Column.FieldName && p.ProjectID == project.ID);

                if (data != null)
                {
                    //Console.WriteLine(e.Column + " " + e.RowHandle + " " + data.Color);

                    e.Appearance.BackColor = data.Color;
                }
            }
        }
        private void projectBandedGridView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && e.Modifiers == Keys.Control)
            {
                if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanDeleteProjects))
                {
                    MessageBox.Show("This login is not authorized to delete projects.");
                    return;
                }

                DeleteProject(sender, e);
            }
        }
        private void changeViewRadioGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            RadioGroup edit = sender as RadioGroup;

            if (edit.SelectedIndex == 0)
            {
                gridControl3.MainView = gridView3;
                workLoadViewPrintButton.Visible = false;
                workLoadViewPrint2Button.Visible = false;
                gridView3.OptionsPrint.PrintDetails = true;
            }
            else if (edit.SelectedIndex == 1)
            {
                gridControl3.MainView = projectBandedGridView;
                workLoadViewPrintButton.Visible = true;
                workLoadViewPrint2Button.Visible = true;
                gridView3.OptionsPrint.PrintDetails = false;
                CollapseGroups();
            }
            else
            {
                gridControl3.MainView = gridView3;
                GridLevelNode node = new GridLevelNode();
                node.RelationName = "DeptProgresses";
                node.LevelTemplate = DeptProgressGridView;
                GridLevelNode deleteNode = gridControl3.LevelTree.Nodes["Components"];
                BaseView oldView = deleteNode.LevelTemplate;
                oldView.Dispose();
                //gridControl3.LevelTree.Nodes.Remove(deleteNode);
                //deleteNode = gridControl3.LevelTree.Nodes["Tasks"];
                //deleteNode.Dispose();
                //gridControl3.LevelTree.Nodes.Remove(deleteNode);
                gridControl3.LevelTree.Nodes.Add(node);

                gridView3.OptionsPrint.PrintDetails = true;
            }
        }
        private void gridView3_CustomRowFilter(object sender, RowFilterEventArgs e)
        {
            ColumnView view = sender as ColumnView;
            ProjectModel project = view.GetRow(e.ListSourceRow) as ProjectModel;

            if (project.Components.Count == 0)
            {
                e.Visible = false;

                e.Handled = true;
            }
        }
        private void gridView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && e.Modifiers == Keys.Control)
            {
                DeleteProject(sender, e);
            }
        }
        private void GridView3_ValidatingEditor(object sender, BaseContainerValidateEditorEventArgs e)
        {
            ColumnView view = sender as ColumnView;
            GridColumn column = (e as EditFormValidateEditorEventArgs)?.Column ?? view.FocusedColumn;

            if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeProjectData))
            {
                e.ErrorText = "This login is not authorized to make changes to project level data.  Hit ESC to cancel editing.";
                e.Valid = false;
            }
            else if (column.FieldName == "ProjectNumber")
            {
                if (int.TryParse(e.Value.ToString(), out int result) == true)
                {
                    if (Database.ProjectExists(result))
                    {
                        e.ErrorText = "There is already a project with that number.  Hit ESC to cancel editing.";
                        e.Valid = false;
                    }
                }
                else
                {
                    e.ErrorText = "Project number must be a number.  Hit ESC to cancel editing.";
                    e.Valid = false;
                }
            }
        }
        private void GridView3_InvalidValueException(object sender, InvalidValueExceptionEventArgs e)
        {
            ColumnView view = sender as ColumnView;
            if (view == null) return;
            e.ExceptionMode = ExceptionMode.DisplayError;
            e.WindowCaption = "Input Error";
            // Destroy the editor and discard the changes made within the edited cell. 
            view.HideEditor();
        }
        private void gridView3_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                GridView view = sender as GridView;
                ProjectModel project = gridView3.GetFocusedRow() as ProjectModel;

                if (!Database.UpdateProjectRecord(project, e))
                {
                    RefreshProjectGrid();
                }
                else
                {
                    project.DateModified = DateTime.Now;

                    if (e.Column.FieldName == "ProjectNumber" || e.Column.FieldName == "JobNumber" || e.Column.FieldName == "DueDate")
                    {
                        project.UpdateKanBan = true;
                    }

                    if (e.Column.FieldName.Contains("Programmer") || e.Column.FieldName == "Designer")
                    {
                        Database.SetTaskPersonnel(project.ProjectNumber, GeneralOperations.FindMatchingDepartment(e.Column.FieldName, Database.GetDepartmentRoles()), e.Value.ToString(), schedulerDataStorage1);
                        RefreshProjectGrid();
                        RefreshDepartmentScheduleView();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
            finally
            {
                SplashScreenManager.CloseForm();
            }
        }
        private void gridView3_ShownEditor(object sender, EventArgs e)
        {
            try
            {
                PopulateShownEditor(sender);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void gridView3_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;
            ProjectModel project = view.GetRow(e.RowHandle) as ProjectModel;
            bool pastDueDate = IsPastDate(project.DueDate, project.LatestFinishDate);  //(DateTime?)view.GetRowCellValue(e.RowHandle, "LatestFinishDate")

            if (e.Column.FieldName == "DueDate")
            {
                e.Appearance.BackColor = pastDueDate ? Color.Red : Color.Empty;
                e.Appearance.ForeColor = pastDueDate ? Color.White : Color.Black;
            }
        }
        private void gridView3_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView View = sender as GridView;
            ProjectModel project = View.GetRow(e.RowHandle) as ProjectModel;

            if (e.RowHandle >= 0)
            {
                if (project.UpdateKanBan)
                {
                    e.Appearance.BackColor = Color.Salmon;
                    e.Appearance.BackColor2 = Color.SeaShell;
                    e.HighPriority = true;
                }

                if (!project.AllTasksDated)
                {
                    e.Appearance.BackColor = Color.Orange;
                    e.Appearance.BackColor2 = Color.SeaShell;
                    e.HighPriority = true;
                }
            }
        }
        private void gridView_MasterRowExpanded(object sender, CustomMasterRowEventArgs e)
        {
            GridView gridView = sender as GridView;
            ExpandedProjectRows epr = new ExpandedProjectRows();
            
            if (gridView.Name == "gridView3")
            {
                //epr.RowHandle = e.RowHandle;
                //expandedProjectRowsList.Add(epr);
                //gridView3ExpandedRowsList.Add(e.RowHandle);
            }
            else if(gridView.Name == "gridView4")
            {
                //epr = expandedProjectRowsList.Find(x => x.RowHandle == gridView.GetParentRowHandle(e.RowHandle));
                //epr.ExpandedComponentRows.Add(e.RowHandle);
                //gridView4ExpandedRowsList.Add(e.RowHandle);
            }

            //MessageBox.Show(gridView.Name + " row expanded.");
        }

        private void gridView_MasterRowCollapsed(object sender, CustomMasterRowEventArgs e)
        {
            GridView gridView = sender as GridView;

            if (gridView.Name == "gridView3")
            {
                //gridView3ExpandedRowsList.Remove(e.RowHandle);
            }
            else if (gridView.Name == "gridView4")
            {
                //gridView4ExpandedRowsList.Remove(e.RowHandle);
            }

            //MessageBox.Show(gridView.Name + " row collapsed.");
        }
        private void gridView4_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView View = sender as GridView;
            ComponentModel component = View.GetRow(e.RowHandle) as ComponentModel;

            if (e.RowHandle >= 0)
            {
                if (!component.AllTasksDated)
                {
                    e.Appearance.BackColor = Color.Orange;
                    e.Appearance.BackColor2 = Color.SeaShell;
                    e.HighPriority = true;
                }
            }
        }
        private void gridView4_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                GridView view = sender as GridView;
                ComponentModel component = view.GetFocusedRow() as ComponentModel;

                Database.UpdateComponent(component, e);

                ProjectsList.Find(x => x.ProjectNumber == component.ProjectNumber).UpdateKanBan = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                SplashScreenManager.CloseForm();
            }
        }
        private void GridView4_ValidatingEditor(object sender, BaseContainerValidateEditorEventArgs e)
        {
            ColumnView view = sender as ColumnView;
            GridColumn column = (e as EditFormValidateEditorEventArgs)?.Column ?? view.FocusedColumn;

            if (column.FieldName == "Component")
            {
                if (e.Value.ToString().Length > ComponentModel.ComponentCharacterLimit)
                {
                    e.ErrorText = $"A component name cannot be longer than {ComponentModel.ComponentCharacterLimit}";
                    e.Valid = false; 
                }
            }
        }

        private void RepositoryItemImageEdit2_Validating(object sender, CancelEventArgs e)
        {
            var edit = sender as ImageEdit;
            try
            {
                bool isGoodComponentPicture = ComponentModel.IsGoodComponentPicture(ComponentModel.NullByteArrayCheck(edit.EditValue));

                if (isGoodComponentPicture == false)
                {
                    e.Cancel = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                MessageBox.Show(ex.Message);
            }
        }
        private void gridView4_CustomRowCellEditForEditing(object sender, CustomRowCellEditEventArgs e)
        {
            if (e.Column.FieldName == "Picture")
            {
                e.RepositoryItem = repositoryItemImageEdit2;
            }
        }
        private void GridView5_ValidatingEditor(object sender, BaseContainerValidateEditorEventArgs e)
        {
            ColumnView view = sender as ColumnView;
            GridColumn column = (e as EditFormValidateEditorEventArgs)?.Column ?? view.FocusedColumn;
            TaskModel task = view.GetFocusedRow() as TaskModel;
            //MessageBox.Show(view.GetRowCellValue(view.FocusedRowHandle, "TaskID").ToString());
            if (column.FieldName == "Duration")
            {
                if (e.Value.ToString() == "")
                {
                    e.ErrorText = "Duration cannot be blank.";
                    e.Valid = false;
                }
                else if (e.Value.ToString().Contains(' '))
                {
                    string[] durationArr = e.Value.ToString().Split(' ');

                    if (!int.TryParse(durationArr[0], out int result))
                    {
                        e.ErrorText = "Number of days must be a whole number.";
                        e.Valid = false;
                    }

                    if (durationArr[1] != "Day(s)")
                    {
                        e.ErrorText = "Missing unit of duration 'Day(s)'.";
                        e.Valid = false;
                    }
                }
                else
                {
                    e.ErrorText = "Missing a space for duration.";
                    e.Valid = false;
                }
            }
            else if (column.FieldName == "TaskID")
            {
                e.ErrorText = "TaskID is not editable.";
                e.Valid = false;
            }
            else if (column.FieldName == "Predecessors")
            {
                string[] predecessorArr;

                if (e.Value.ToString().Contains(','))
                {
                    predecessorArr = e.Value.ToString().Split(',');

                    foreach (string predecessor in predecessorArr)
                    {
                        if (predecessor == view.GetRowCellValue(view.FocusedRowHandle, "TaskID").ToString())
                        {
                            e.ErrorText = "A task can't be it's own predecessor.";
                            e.Valid = false;
                        }
                    }
                }
                else if (e.Value.ToString() == view.GetRowCellValue(view.FocusedRowHandle, "TaskID").ToString())
                {
                    e.ErrorText = "A task can't be it's own predecessor.";
                    e.Valid = false;
                }
            }
            else if (column.FieldName == "StartDate" || column.FieldName == "FinishDate")
            {
                if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeDates))
                {
                    e.ErrorText = "This login is not authorized to make changes to dates.";
                    e.Valid = false;
                }

                if (column.FieldName == "FinishDate" && task.StartDate != null)
                {
                    if (e.Value != null && DateTime.Parse(e.Value.ToString()) < task.StartDate)
                    {
                        e.ErrorText = "You cannot have a finish date before a task's start date.";
                        e.Valid = false;
                    }
                }
            }
        }
        private void GridView5_InvalidValueException(object sender, InvalidValueExceptionEventArgs e)
        {
            ColumnView view = sender as ColumnView;
            if (view == null) return;

            e.ExceptionMode = ExceptionMode.DisplayError;
            e.WindowCaption = "Input Error";
            //e.ErrorText = "The value should be greater than 0 and less than 1,000,000";
            // Destroy the editor and discard the changes made within the edited cell. 
            view.HideEditor();
        }
        private void gridView5_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                GridView view = sender as GridView;
                TaskModel task = view.GetFocusedRow() as TaskModel;
                ProjectModel project = ProjectsList.Find(x => x.ProjectNumber == task.ProjectNumber);
                ComponentModel component = ComponentsList.Find(x => x.Component == task.Component && x.ProjectNumber == task.ProjectNumber);

                schedulerControl1.BeginUpdate();
                gridView1.BeginUpdate();
                gridView5.BeginUpdate();

                if ((e.Column.FieldName == "StartDate" || e.Column.FieldName == "FinishDate") && e.Value != DBNull.Value)
                {
                    component.ChangeTaskDate(e.Column.FieldName, task);  // Database update handled in this method.
                    //project.IsProjectOnTime();
                }
                else
                {
                    if (e.Column.FieldName == "Machine" || e.Column.FieldName == "Personnel")
                    {
                        task.Resources = GeneralOperations.GenerateResourceIDsString(schedulerDataStorage1, task.Machine, task.Personnel);
                    }

                    Database.UpdateTask(task, e);  // Resources field is only updated when the Machine or Resource fields change., resources

                    if (e.Column.FieldName == "TaskName" || e.Column.FieldName == "Notes")
                    {
                        project.UpdateKanBan = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if(IsFormOpen("WaitForm1"))
                    SplashScreenManager.CloseForm();

                schedulerControl1.EndUpdate();
                schedulerControl1.RefreshData();
                gridView1.EndUpdate();
                gridView5.EndUpdate();
            }
        }
        private void RefreshProjectsButton_Click(object sender, EventArgs e)
        {
            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                RefreshProjectGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                SplashScreenManager.CloseForm();
            }
        }

        private void copyButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridControl3.MainView != gridView3)
                {
                    MessageBox.Show("Only projects in the Project View can be copied.");
                    return;
                }

                ProjectModel project = Database.GetProject((int)gridView3.GetFocusedRowCellValue("ProjectNumber")); // (string)gridView3.GetFocusedRowCellValue("JobNumber"),
                ProjectModel copiedProject;

                if (gridView3.SelectedRowsCount == 1)
                {
                    var result = XtraInputBox.Show("Change Project #", "Copy Project", "Enter a project number.");

                    if (result.Length > 0)
                    {
                        if (int.TryParse(result.ToString(), out int projectNumber))
                        {
                            project.AvailableResources = schedulerDataStorage1;
                            copiedProject = new ProjectModel(project, projectNumber);

                            if(Database.CreateProject(copiedProject))
                            {
                                RefreshProjectGrid();
                            }

                            int rowHandle = gridView3.LocateByValue("ProjectNumber", projectNumber);
                            if (rowHandle != GridControl.InvalidRowHandle)
                                gridView3.FocusedRowHandle = rowHandle;
                            gridView3.SetMasterRowExpanded(gridView3.FocusedRowHandle, true);
                        }
                        else
                        {
                            MessageBox.Show("Please enter a number for a Project Number.");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select a project.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void kanBanButton_Click(object sender, EventArgs e)
        {
            GenerateKanBan();
            //GenerateKanBanReport();
        }

        private void forwardDateButton_Click(object sender, EventArgs e)
        {
            if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeDates))
            {
                MessageBox.Show("This login is not authorized to make changes to dates.");
                return;
            }

            if (gridControl3.MainView != gridView3)
            {
                MessageBox.Show("Projects can only be scheduled in Project View.");
                return;
            }

            List<ComponentModel> selectedComponentList = GetListOfSelectedComponents();

            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                var compResult = from component in ComponentsList
                                 where selectedComponentList.Any(x => x.Component == component.Component && x.ProjectNumber == component.ProjectNumber)
                                 select component;

                if (selectedComponentList.Count == 0)
                {
                    MessageBox.Show("No components selected.");

                    return;
                }

                schedulerControl1.BeginUpdate();
                gridView1.BeginUpdate();
                gridView5.BeginUpdate();

                using (var form = new ForwardDateWindow("Forward Date", DateTime.Today))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        Database.ForwardDateTasks(compResult.ToList(), form.ForwardDate);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                Console.WriteLine(ex.ToString());
                RefreshProjectGrid();
            }
            finally
            {
                SplashScreenManager.CloseForm();
                schedulerControl1.EndUpdate();
                schedulerControl1.RefreshData();
                gridView1.EndUpdate();
                gridView5.EndUpdate();
            }
        }

        private void backDateButton_Click(object sender, EventArgs e)
        {
            if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeDates))
            {
                MessageBox.Show("This login is not authorized to make changes to dates.");
                return;
            }

            if (gridControl3.MainView != gridView3)
            {
                MessageBox.Show("Projects can only be scheduled in Project View.");
                return;
            }

            ProjectModel project = gridView3.GetFocusedRow() as ProjectModel;
            List<ComponentModel> selectedComponentList = GetListOfSelectedComponents();

            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                var compResult = from component in ComponentsList
                                 where selectedComponentList.Any(x => x.Component == component.Component && x.ProjectNumber == component.ProjectNumber)
                                 select component;

                if (selectedComponentList.Count == 0)
                {
                    XtraMessageBox.Show("No components selected.");
                    return;
                }

                schedulerControl1.BeginUpdate();
                gridView1.BeginUpdate();
                gridView5.BeginUpdate();

                using (var form = new ForwardDateWindow("Back Date", project.DueDate))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        Database.BackDateTasks(compResult.ToList(), form.ForwardDate);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                SplashScreenManager.CloseForm();
                schedulerControl1.EndUpdate();
                schedulerControl1.RefreshData();
                gridView1.EndUpdate();
                gridView5.EndUpdate();
            }
        }
        // Recreate Create Project button to reactivate.
        private void createProjectButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine("click");

            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                schedulerControl1.BeginUpdate();
                gridView1.BeginUpdate();
                gridView5.BeginUpdate();

                CreateProject();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
            finally
            {
                SplashScreenManager.CloseForm();
                schedulerControl1.EndUpdate();
                schedulerControl1.RefreshData();
                gridView1.EndUpdate();
                gridView5.EndUpdate();
            }
        }

        private void editProjectButton_Click(object sender, EventArgs e)
        {
            try
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));

                GridView view = gridControl3.MainView as GridView;

                ProjectModel selectedProject = view.GetFocusedRow() as ProjectModel;

                ProjectModel project = Database.GetProject(selectedProject.ProjectNumber);

                project.OldProjectNumber = project.ProjectNumber;

                schedulerControl1.BeginUpdate();
                gridView1.BeginUpdate();
                gridView5.BeginUpdate();

                EditProject(project);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
            finally
            {
                SplashScreenManager.CloseForm();
                schedulerControl1.EndUpdate();
                schedulerControl1.RefreshData();
                gridView1.EndUpdate();
                gridView5.EndUpdate();
            }
        }

        private List<ComponentModel> GetListOfSelectedComponents()
        {
            List<ComponentModel> componentList = new List<ComponentModel>();

            if (gridView3.GetMasterRowExpanded(gridView3.FocusedRowHandle))
            {
                var childView = gridView3.GetVisibleDetailView(gridView3.FocusedRowHandle) as GridView;

                foreach (int rowHandle in childView.GetSelectedRows())
                {
                    componentList.Add(new ComponentModel {ProjectNumber = (int)gridView3.GetRowCellValue(gridView3.FocusedRowHandle, "ProjectNumber"), Component = (string)childView.GetRowCellValue(rowHandle, "Component") });
                }
            }

            return componentList;
        }
        private void restoreProjectButton_Click(object sender, EventArgs e)
        {
            XtraInputBoxArgs args = new XtraInputBoxArgs();
            args.Caption = "Select Project to Restore";
            args.Prompt = $"{DeletedProjects.Count} Deleted Projects";
            args.DefaultButtonIndex = 0;

            LookUpEdit lookUpEdit = new LookUpEdit();
            lookUpEdit.Properties.Columns.Add(new LookUpColumnInfo("JobNumber", "Job #"));
            lookUpEdit.Properties.Columns.Add(new LookUpColumnInfo("ProjectNumber", "Project #"));
            lookUpEdit.Properties.Columns.Add(new LookUpColumnInfo("Project", "Project"));
            lookUpEdit.Properties.Columns.Add(new LookUpColumnInfo("Customer", "Customer"));
            lookUpEdit.Properties.DataSource = DeletedProjects;
            lookUpEdit.Properties.KeyMember = "ProjectNumber";
            lookUpEdit.Properties.DisplayMember = "ProjectNumber";

            args.Editor = lookUpEdit;

            ProjectModel result = (ProjectModel)XtraInputBox.Show(args);

            if (result != null)
            {
                Database.CreateProject(result);

                RefreshProjectGrid();

                DeletedProjects.Remove(result);
            }
        }
        private void resourceButton_Click(object sender, EventArgs e)
        {
            using (ManageResourcesForm form = new ManageResourcesForm())
            {
                form.ShowDialog(); // Code execution stops until user does something with the window.
                RoleTable = Database.GetRoleTable();
            }
        }
        private void adminButton_Click(object sender, EventArgs e)
        {
            using (AdminWindow adminWindow = new AdminWindow())
            {
                adminWindow.ShowDialog();
                UserList = Database.GetUsers();
            }
        }
        private void workLoadViewPrintButton_Click(object sender, EventArgs e)
        {
            // Check whether the GridControl can be previewed.
            if (!gridControl3.IsPrintingAvailable)
            {
                MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                return;
            }

            PrintOrientation = "Landscape";
            PaperSize = "Tabloid";
            FieldInfo fi = typeof(GridColumn).GetField("minWidth", BindingFlags.NonPublic | BindingFlags.Instance);
            fi.SetValue(projectBandedGridView.Columns.ColumnByFieldName("Stage"), 0);

            projectBandedGridView.Print();

            gridControl3.Print();
        }
        private void workLoadViewPrint2Button_Click(object sender, EventArgs e)
        {
            // Check whether the GridControl can be previewed.
            if (!gridControl3.IsPrintingAvailable)
            {
                MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                return;
            }

            PrintOrientation = "Portrait";
            PaperSize = "Letter";
            FieldInfo fi = typeof(GridColumn).GetField("minWidth", BindingFlags.NonPublic | BindingFlags.Instance);
            fi.SetValue(projectBandedGridView.Columns.ColumnByFieldName("Stage"), 0);

            projectBandedGridView.Print();
        }

        #endregion

        #region Chart View
        public List<Week> InitializeWeeksList(List<string> weekCategoryList)
        {
            List<Week> weekList = new List<Week>();
            DateTime wsDate;
            wsDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);

            for (int i = 1; i <= 20; i++)
            {
                foreach (string category in weekCategoryList)
                {
                    weekList.Add(new Week(i, wsDate.AddDays((i - 1) * 7), category));
                }
            }

            return weekList;
        }
        private List<Week> GetWeeks(List<TaskModel> tasks, string resourceType = "Department")
        {
            List<string> categories = new List<string>();
            List<Week> weekList;
            List<Week> deptWeekList = new List<Week>();
            Week weekTemp;
            int weekNum;

            if (resourceType == "Department")
            {
                categories = Database.GetDepartments();
            }
            else if (resourceType == "Machines")
            {
                categories = Database.GetMachines();
            }

            weekList = InitializeWeeksList(categories);

            foreach (TaskModel task in tasks)
            {
                if (task.StartDate == null || task.FinishDate == null)
                {
                    goto Skip;
                }
                // Selects the weeks with the correct category item.
                if (resourceType == "Department")
                {
                    var results = from wk in weekList
                                  where (task.TaskName.StartsWith(wk.Department) || (task.TaskName.Contains("Grind") && task.TaskName.Contains(wk.Department))) // && Convert.ToDateTime(rdr["StartDate"]) >= wk.WeekStart && Convert.ToDateTime(rdr["StartDate"]) <= wk.WeekEnd
                                  orderby wk.WeekNum ascending
                                  select wk;

                    deptWeekList = results.ToList();
                }
                else if (resourceType == "Personnel")
                {
                    var results = from wk in weekList
                                  where (task.Personnel.Contains(wk.Department)) // && Convert.ToDateTime(rdr["StartDate"]) >= wk.WeekStart && Convert.ToDateTime(rdr["StartDate"]) <= wk.WeekEnd
                                  orderby wk.WeekNum ascending
                                  select wk;

                    deptWeekList = results.ToList();
                }
                else if (resourceType == "Machines")
                {
                    var results = from wk in weekList
                                  where (task.Machine != null && task.Machine.Contains(wk.Department))
                                  orderby wk.WeekNum ascending
                                  select wk;

                    deptWeekList = results.ToList();
                }
                // Deposits hours into correct day buckets.
                if (deptWeekList.Any())
                {
                    weekTemp = deptWeekList.Find(x => (x.WeekStart <= task.StartDate && x.WeekEnd >= task.StartDate) || (x.WeekStart > task.StartDate && x.WeekNum == 1));
                    weekNum = weekTemp.WeekNum;
                    //weekTemp.AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));

                    double hours = task.Hours;
                    double days = (int)Database.GetBusinessDays(task.StartDate ?? new DateTime(1, 1, 0001), task.FinishDate ?? new DateTime(1, 1, 0001));
                    //double days = Database.BusinessDaysUntil(task.StartDate ?? new DateTime(1, 1, 0001), task.FinishDate ?? new DateTime(1, 1, 0001));
                    //double days = int.Parse(task.Duration.Split(' ')[0]);

                    decimal dailyAVG;

                    if (days == 0)
                    {
                        dailyAVG = (decimal)hours;
                    }
                    else
                    {
                        dailyAVG = (decimal)(hours / days);
                    }

                    DateTime date;

                    if (weekTemp.WeekStart > task.StartDate && weekNum == 1)
                    {
                        date = weekTemp.WeekStart.AddDays(1);
                        days = days - Database.GetBusinessDays(task.StartDate ?? new DateTime(1, 1, 0001), date);
                    }
                    else
                    {
                        date = task.StartDate ?? new DateTime(1, 1, 0001);
                    }

                    if (days >= 1)
                    {
                        while (days > 0)
                        {
                            if (date.DayOfWeek == DayOfWeek.Saturday)
                            {
                                date = date.AddDays(1);

                                weekNum++;

                                if (weekNum > 20)
                                {
                                    goto MyEnd;
                                }

                                //weekTemp = deptWeekList.Find(x => x.WeekNum == weekNum);
                                weekTemp = deptWeekList[weekNum - 1];
                                //weekTemp.AddHoursToDay((int)date.DayOfWeek, dailyAVG);
                                //Console.WriteLine($"{weekTemp.Department} {weekTemp.WeekStart.ToShortDateString()} {date.DayOfWeek} {dailyAVG} {days}");
                            }
                            else
                            {
                                weekTemp.AddHoursToDay((int)date.DayOfWeek, dailyAVG);
                                //if (weekTemp.Department == "CNC Finish")
                                //    Console.WriteLine($"{weekTemp.Department} {weekTemp.WeekStart.ToShortDateString()} {date.DayOfWeek} Daily AVG. {dailyAVG} Hrs {hours} Days {days}");
                                days -= 1;
                            }


                            date = date.AddDays(1);
                        }
                    }
                    else
                    {
                        weekTemp.AddHoursToDay((int)date.AddDays(days).DayOfWeek, dailyAVG);
                        //if (weekTemp.Department == "CNC Finish")
                        //    Console.WriteLine($"{weekTemp.Department} {weekTemp.WeekStart.ToShortDateString()} {date.AddDays(days).DayOfWeek} {dailyAVG} {days}");
                    }
                }

            Skip:;
            }

        MyEnd:;

            return weekList;
        }
        private void LoadGraph(List<Week> weekList, List<string> departmentList)
        {
            Series tempSeries;

            int i = 0;

            chartControl1.Series.Clear();

            List<string> weekTitleArr = new List<string>();
            DataTable dailyDeptCapacities = Database.GetDailyDepartmentCapacities();

            for (int n = 0; n < 20; n++)
            {
                weekTitleArr.Add(n.ToString());
            }

            if (TimeUnits == "Days")
            {
                var results = from wks in weekList
                                where wks.WeekStart == DateTime.Parse(timeFrameComboBoxEdit.Text.Split(' ')[0])
                                orderby wks.WeekStart
                                select wks;

                foreach (Week week in results.ToList())
                {                    
                    tempSeries = new Series(week.Department, ViewType.Bar);

                    foreach (ClassLibrary.Day day in week.DayList)
                    {
                        tempSeries.Points.Add(new SeriesPoint(day.DayName, Decimal.Round(day.Hours, 1)));
                    }

                    chartControl1.Series.Add(tempSeries);
                }
            }
            else if(TimeUnits == "Weeks")
            {
                foreach (string dept in departmentList)
                {
                    // Creates a new series for the current department.
                    tempSeries = new Series(dept, ViewType.Bar);

                    // Selects all weeks corresponding to the current dept.
                    var deptWeeks = from wks in weekList
                                    where wks.Department == dept
                                    orderby wks.WeekStart
                                    select wks;

                    // Cycles through each of a departments weeks.
                    foreach (Week week in deptWeeks)
                    {
                        // Adds a point for each week using the sum of the weeks hours as a data point.
                        tempSeries.Points.Add(new SeriesPoint("WK " + weekTitleArr[i++] + " " + week.WeekStart.ToShortDateString() , Decimal.Round(week.Hours, 1)));
                    }

                    chartControl1.Series.Add(tempSeries);

                    i = 0;
                }
            }
        }

        private void PopulateDeptForecastHours(List<Week> deptWeekList)
        {
            IWorkbook workbook = spreadsheetControl1.Document;
            Worksheet worksheet1 = workbook.Worksheets["Dept. Weekly Hrs"];
            Worksheet worksheet2 = workbook.Worksheets["Dept. Daily Hrs"];
            List<Week> currentWeekList1, currentWeekList2;
            DateTime currentDate1, currentDate2;
            DateTime wsDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
            DateTime date = DateTime.Today;

            spreadsheetControl1.BeginUpdate();

            for (int c = 2; c <= 20; c++)
            {
                worksheet1.Cells[2, c].Value = wsDate.AddDays((c - 2) * 7);
                currentDate1 = DateTime.Parse(worksheet1.Cells[2, c].Value.ToString());
                currentWeekList1 = deptWeekList.FindAll(x => x.WeekStart == currentDate1);

                worksheet2.Cells[2, c].Value = GeneralOperations.AddBusinessDays(date, c - 2);
                currentDate2 = DateTime.Parse(worksheet2.Cells[2, c].Value.ToString());
                currentWeekList2 = deptWeekList.FindAll(x => x.WeekStart <= currentDate2 && x.WeekEnd >= currentDate2);

                for (int r = 3; r <= 16; r++)
                {
                    worksheet1.Cells[r, c].Value = Decimal.Round(currentWeekList1.Find(x => x.Department.Contains(worksheet1.Cells[r, 1].Value.ToString())).Hours, 1);

                    worksheet2.Cells[r, c].Value = Decimal.Round(currentWeekList2.Find(x => x.Department.Contains(worksheet2.Cells[r, 1].Value.ToString())).DayList.Find(x => x.Date == currentDate2).Hours,1);
                }
            }

            spreadsheetControl1.EndUpdate();
        }

        private void PopulateMachineHours(List<Week> machineWeekList)
        {
            IWorkbook workbook = spreadsheetControl1.Document;
            Worksheet worksheet1 = workbook.Worksheets["Mach. Weekly Hrs"];
            Worksheet worksheet2 = workbook.Worksheets["Mach. Daily Hrs"];

            List<Week> currentWeekList1, currentWeekList2;
            DateTime currentDate1, currentDate2;
            DateTime wsDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
            DateTime date = DateTime.Today;
            List<string> machineList = machineWeekList.Select(x => x.Department).Distinct().ToList();

            spreadsheetControl1.BeginUpdate();

            int r1 = 3;

            foreach (var machine in machineList)
            {
                worksheet1.Cells[r1, 1].Value = machine;
                worksheet2.Cells[r1++, 1].Value = machine;
            }

            for (int c = 2; c <= 20; c++)
            {
                worksheet1.Cells[2, c].Value = wsDate.AddDays((c - 2) * 7);
                currentDate1 = DateTime.Parse(worksheet1.Cells[2, c].Value.ToString());
                currentWeekList1 = machineWeekList.FindAll(x => x.WeekStart == currentDate1);

                worksheet2.Cells[2, c].Value = GeneralOperations.AddBusinessDays(date, c - 2);
                currentDate2 = DateTime.Parse(worksheet2.Cells[2, c].Value.ToString());
                currentWeekList2 = machineWeekList.FindAll(x => x.WeekStart <= currentDate2 && x.WeekEnd >= currentDate2);

                for (int r2 = 3; r2 <= 14; r2++)
                {
                    worksheet1.Cells[r2, c].Value = Decimal.Round(currentWeekList1.Find(x => x.Department.Contains(worksheet1.Cells[r2, 1].Value.ToString())).Hours, 1);

                    worksheet2.Cells[r2, c].Value = Decimal.Round(currentWeekList2.Find(x => x.Department.Contains(worksheet2.Cells[r2, 1].Value.ToString())).DayList.Find(x => x.Date == currentDate2).Hours, 1);
                }
            }

            spreadsheetControl1.EndUpdate();
        }

        public string SaveExcelFile(string fileName)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = @"C:\Users\" + Environment.UserName + @"\Desktop";
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"; // Text files (*.txt)|*.txt|
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = false;
            saveFileDialog.FileName = fileName;
            saveFileDialog.Title = "Save Chart Data";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Save. The selected path can be got with saveFileDialog.FileName.ToString()
                return saveFileDialog.FileName.ToString();
            }
            else
            {
                return "";
            }
        }

        private void PopulateTimeFrameComboBox()
        {
            DateTime weekStart = new DateTime();

            weekStart = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);

            timeFrameComboBoxEdit.Properties.Items.Clear();

            if (TimeUnits == "Days")
            {
                for (int i = 0; i < 20; i++)
                {
                    // One week spans of time.
                    timeFrameComboBoxEdit.Properties.Items.Add($"{weekStart.AddDays(i * 7).ToShortDateString()} - {weekStart.AddDays(i * 7 + 6).ToShortDateString()}");
                }
            }
            else if(TimeUnits == "Weeks")
            {
                // 20 week spans of time.
                timeFrameComboBoxEdit.Properties.Items.Add($"{weekStart.ToShortDateString()} - {weekStart.AddDays((19 * 7 + 6)).ToShortDateString()}");
            }
        }

        private int GetDeptDailyCapacity(string department, DataTable deptCapacityDT)
        {
            //var result = from DataRow myRow in deptCapacityDT.Rows
            //             where myRow.Field<string>("Department") == department
            //             select new { dailyCapacity = myRow.Field<int>("DailyCapacity") };

            //int dailyCapacity = deptCapacityDT.AsEnumerable().Where(p => p.Field<string>("Department") == department).Select(p => p.Field<int>("DailyCapacity")).FirstOrDefault();

            return deptCapacityDT.AsEnumerable().Where(p => p.Field<string>("Department") == department).Select(p => p.Field<int>("DailyCapacity")).FirstOrDefault();
        }
        private void GetOverallToolRoomHours()
        {
            List<TaskModel> taskList = new List<TaskModel>();
            List<Week> deptWeeksList = new List<Week>();
            List<Week> machineWeekList = new List<Week>();
            List<string> departmentList = new List<string>();

            departmentList = Database.GetDepartments();

            if (timeFrameComboBoxEdit.Text != "")
            {
                string weekStart, weekEnd;

                weekStart = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek).ToShortDateString();
                weekEnd = DateTime.Parse(weekStart).AddDays((19 * 7 + 6)).ToShortDateString();

                taskList = Database.GetTasks(weekStart, weekEnd);

                deptWeeksList = GetWeeks(taskList);
                machineWeekList = GetWeeks(taskList, "Machines");

                LoadGraph(deptWeeksList, departmentList);

                PopulateDeptForecastHours(deptWeeksList);
                PopulateMachineHours(machineWeekList);
            }
        }
        private string GetResourceType()
        {
            return chartRadioGroup.Properties.Items[chartRadioGroup.SelectedIndex].Description.ToString();
        }

        private void RefreshChartButton_Click(object sender, EventArgs e)
        {
            try
            {
                GetOverallToolRoomHours();
                //GetDepartmentHours();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void timeFrameComboBoxEdit_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                GetOverallToolRoomHours();
                //GetDepartmentHours();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }

        private void TimeUnitsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (TimeUnitsComboBox.Text == "Days")
                {
                    this.TimeUnits = "Days";
                    PopulateTimeFrameComboBox();
                    timeFrameComboBoxEdit.SelectedIndex = 0;
                }
                else if (TimeUnitsComboBox.Text == "Weeks")
                {
                    this.TimeUnits = "Weeks";
                    PopulateTimeFrameComboBox();
                    timeFrameComboBoxEdit.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void chartRadioGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            chartViewNavigationFrame.SelectedPageIndex = chartRadioGroup.SelectedIndex;

            if (chartRadioGroup.SelectedIndex == 1)
            {
                exportButton.Visible = true;
            }
            else
            {
                exportButton.Visible = false;
            }
        }

        private void exportButton_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Today;
            string fileName = SaveExcelFile($"Forecast Hours {date.Month}-{date.Day}-{date.Year}");

            if (fileName.Length > 0)
            {
                spreadsheetControl1.Document.SaveDocument(fileName);
            }            
        }

        #endregion

        #region Gantt View

        private void InitializeResources(ProjectModel project)
        {
            int i = 0;
            int ParentID = 0;

            //CustomResourceCollection.Clear();
            CustomResourceCollection = new BindingList<CustomResource>();

            foreach (ComponentModel component in project.Components)
            {
                ParentID = i;
                CustomResourceCollection.Add(CreateCustomResource(i++, -1, component.Component));

                foreach (TaskModel task in component.Tasks)
                {
                    CustomResourceCollection.Add(CreateCustomResource(i++, ParentID, task.TaskName));
                }
            }

            ResourceMappingInfo mappings = this.schedulerStorage2.Resources.Mappings;

            mappings.Id = "ResID";
            //mappings.Color = "ResColor";
            mappings.Caption = "Name";
            mappings.ParentId = "ParentID";

            schedulerStorage2.Resources.Clear();

            schedulerStorage2.Resources.DataSource = CustomResourceCollection;

            Console.WriteLine("Initialize Resources");

            for (int id = 0; id < schedulerStorage2.Resources.Count; id++)
            {
                Resource resource = schedulerStorage2.Resources[id];

                Console.WriteLine($"{resource.Id} {resource.Caption} {resource.ParentId}");
            }
        }

        private void GenerateEventList(BindingList<CustomAppointment> eventList, ProjectModel project)
        {
            int i = 0;
            int baseCount;

            foreach (ComponentModel component in project.Components)
            {
                baseCount = i++;

                foreach (TaskModel task in component.Tasks)
                {
                    Resource resource = schedulerStorage2.Resources[i++];
                    eventList.Add(CreateEvent(task.TaskID + baseCount, project.JobNumber + " #" + project.ProjectNumber + " " + component.Component, resource.Id, task.TaskID, task.TaskName, task.StartDate, task.FinishDate));
                }
            }

            Console.WriteLine("Initialize Appointments");

            foreach (CustomAppointment apt in eventList)
            {
                //Resource resource = schedulerStorage2.Resources[Convert.ToInt16(apt.OwnerId)];
                Console.WriteLine($"{apt.AppointmentId} {apt.Subject} {apt.OwnerId} {apt.Subject}");
            }
        }

        private CustomResource CreateCustomResource(int res_id, int parent_Id, string caption)
        {
            CustomResource cr = new CustomResource();
            cr.ResID = res_id;
            cr.ParentID = parent_Id;
            cr.Name = caption;
            return cr;
        }

        private CustomAppointment CreateEvent(int appointmentId, string subject, object resourceId, int taskId, string location, DateTime? startDate, DateTime? finishDate)
        {
            CustomAppointment apt = new CustomAppointment();

            apt.AppointmentId = appointmentId;
            apt.Subject = subject;
            apt.Location = location;
            apt.OwnerId = resourceId;
            apt.StartDate = Convert.ToDateTime(startDate);
            apt.FinishDate = Convert.ToDateTime(finishDate);
            apt.TaskId = taskId;

            return apt;
        }

        private CustomDependency CreateCustomDependency(int dep_id, int par_id)
        {
            CustomDependency cd = new CustomDependency();
            cd.DepID = dep_id;
            cd.ParentID = par_id;

            return cd;
        }

        private void InitializeDependencies(ProjectModel project)
        {
            int aID = 1;
            int baseCount;

            foreach (ComponentModel component in project.Components)
            {
                baseCount = aID - 1;

                foreach (TaskModel task in component.Tasks)
                {
                    task.ChangeIDs(baseCount);

                    if (task.Predecessors.Contains(","))
                    {
                        foreach (string predID in task.Predecessors.Split(','))
                        {
                            CustomDependencyList.Add(CreateCustomDependency(aID, Convert.ToInt32(predID)));
                        }

                        aID++;
                    }
                    else if (task.Predecessors != "")
                    {
                        CustomDependencyList.Add(CreateCustomDependency(aID++, Convert.ToInt32(task.Predecessors)));
                    }
                    else
                    {
                        aID++;
                    }
                }
            }
        }

        private void PopulateProjectComboBox()
        {
            projectComboBox.Properties.Items.Clear();

            foreach (string item in Database.GetJobNumberComboList())
            {
                projectComboBox.Properties.Items.Add(item);
            }
        }

        private void LoadProject(int projectNumber)
        {
            Project = Database.GetProject(projectNumber);

            //Project = ProjectsList.Find(x => x.ProjectNumber == projectNumber);

            Project.HasTasksWithNullDates();

            ResourceMappingInfo resourceMappings = this.schedulerStorage2.Resources.Mappings;

            resourceMappings.Id = "AptID";
            resourceMappings.ParentId = "ParentID"; // Need this for hierarchy in resource tree.
            resourceMappings.Caption = "TaskName"; // In the Resource tree designer the field name has to match the field that is mapped to caption.

            schedulerStorage2.Resources.Clear();

            BindingList<AptResourceModel> aptResources = new BindingList<AptResourceModel>(AptResourceModel.GetProjectResourceData(Project));

            schedulerStorage2.Resources.DataSource = aptResources; // Woohoo!! This finally works!

            if (schedulerStorage2.Appointments.Count > 0)
            {
                schedulerStorage2.Appointments.Clear();
            }

            schedulerStorage2.Appointments.CustomFieldMappings.Clear();

            schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("JobNumber", "JobNumber"));
            schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("ProjectNumber", "ProjectNumber"));
            schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskID", "TaskID"));
            schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskName", "TaskName"));
            schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Component", "Component"));
            schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Hours", "Hours"));
            schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Predecessors", "Predecessors"));

            AppointmentMappingInfo appointmentMappings = schedulerStorage2.Appointments.Mappings;

            appointmentMappings.AppointmentId = "AptID";
            appointmentMappings.Start = "StartDate";
            appointmentMappings.End = "FinishDate";
            appointmentMappings.Subject = "Subject";
            appointmentMappings.Location = "Location";
            appointmentMappings.Description = "Notes";
            appointmentMappings.PercentComplete = "PercentComplete";
            appointmentMappings.ResourceId = "AptID";

            BindingList<TaskModel> taskList = new BindingList<TaskModel>(Project.GetTaskList());

            schedulerStorage2.Appointments.DataSource = taskList;

            AppointmentDependencyMappingInfo appointmentDependencyMappingInfo = schedulerStorage2.AppointmentDependencies.Mappings;

            appointmentDependencyMappingInfo.DependentId = "DependentID";
            appointmentDependencyMappingInfo.ParentId = "ParentID";

            BindingList<AptDependencyModel> dependencyList = new BindingList<AptDependencyModel>(AptDependencyModel.GetDependencyData(Project));

            schedulerStorage2.AppointmentDependencies.DataSource = dependencyList;
        }

        private List<int> GetCollapsedNodes()
        {
            List<int> collapsedNodes = new List<int>();

            for (int i = 0; i < resourcesTree1.Nodes.Count; i++)
            {
                if (!resourcesTree1.Nodes[i].Expanded)
                {
                    collapsedNodes.Add(i);
                }
            }

            return collapsedNodes;
        }

        private void CollapseNodes(List<int> collapsedNodes)
        {
            foreach (int node in collapsedNodes)
            {
                resourcesTree1.Nodes[node].Collapse();
            }
        }

        // This is for the resource tree in the Gantt view and has nothing do with the resources table.
        private DataTable GetProjectResourceData(ProjectModel project)
        {
            DataTable dt = new DataTable();
            int i = 1;
            int parentID = 0;

            dt.Columns.Add("AptID", typeof(int));
            dt.Columns.Add("TaskName", typeof(string));
            dt.Columns.Add("ParentID", typeof(int));

            foreach (ComponentModel component in project.Components)
            {
                DataRow newRow1 = dt.NewRow();

                newRow1["AptID"] = i;
                newRow1["TaskName"] = component.Component;
                parentID = i++;

                dt.Rows.Add(newRow1);

                foreach (TaskModel task in component.Tasks)
                {
                    DataRow newRow2 = dt.NewRow();

                    newRow2["AptID"] = i;
                    newRow2["TaskName"] = task.TaskName;
                    newRow2["ParentID"] = parentID;

                    Console.WriteLine(newRow2["TaskName"].ToString());

                    dt.Rows.Add(newRow2);

                    i++;
                }
            }

            return dt;
        }
        private void LoadProjectGantt()
        {
            var number = GetProjectComboBoxInfo(projectComboBox);
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            //LoadProject();
            LoadProject(number.projectNumber);

            SplashScreenManager.CloseForm();
        }
        private bool UpdateTaskStorage2(TaskModel movedTask)
        {
            ComponentModel tempComponent;
            ProjectModel globalProject = null;

            //var number = GetProjectComboBoxInfo(projectComboBox);

            //Resource resource = schedulerStorage2.Resources[Convert.ToInt16(apt.ResourceId) - 1];

            //resource = schedulerStorage2.Resources[Convert.ToInt16(resource.ParentId) - 1];

            //globalProject = ProjectsList.Find(x => x.ProjectNumber == movedTask.ProjectNumber);

            tempComponent = Project.Components.Find(x => x.ProjectNumber == movedTask.ProjectNumber && x.Component == movedTask.Component);

            //movedTask = globalComponent.Tasks.Find(x => x.TaskID == (int)apt.CustomFields["TaskID"]);

            //movedTask.SetDates(apt.Start, apt.End);

            //schedulerControl1.BeginUpdate();
            //schedulerControl2.BeginUpdate();
            //gridView1.BeginUpdate();
            //gridView5.BeginUpdate();

            //try
            //{
                if (RightMouseButtonPressed)
                {
                    foreach (TaskModel task in tempComponent.Tasks)
                    {
                        Console.WriteLine($"Task: {task.TaskName,-13} Start Date: {((DateTime)task.StartDate).ToShortDateString(),-10} Finish Date: {GeneralOperations.AddBusinessDays((DateTime)task.StartDate, task.Duration).ToShortDateString()}");
                    }

                    Console.WriteLine();

                    if (MoveSubsequentTaskWithLockedSpacing)
                    {
                        if (((DateTime)movedTask.StartDate - OldTaskStartDate).Days > 0)
                        {
                            tempComponent.UpdateSuccessorTaskDates(movedTask, GeneralOperations.GetWorkingDays(OldTaskStartDate, (DateTime)movedTask.StartDate));
                        }
                        else
                        {
                            foreach (int taskID in movedTask.GetPredecessorList())
                            {
                                tempComponent.UpdatePredecessorTaskDates(taskID, GeneralOperations.GetWorkingDays((DateTime)movedTask.StartDate, OldTaskStartDate));
                            }
                        }

                        Database.UpdateTaskDates(tempComponent.Tasks);

                        MoveSubsequentTaskWithLockedSpacing = false;
                    }
                    else
                    {
                        tempComponent.UpdateTaskDates(movedTask, OldTaskStartDate);
                    }

                    return true;
                }
                else
                {
                    if (tempComponent.UpdateTaskDates(movedTask)) // Database.UpdateTask(movedTask, tempComponent), db.UpdateTask(number.jobNumber, number.projectNumber, component.Component, task.TaskID, apt.Start, apt.End, Project.OverlapAllowed)
                    {
                        return true;
                    }
                }
            //}
            //finally
            //{
            //    gridView5.EndUpdate();
            //    gridView1.EndUpdate();
            //    schedulerControl1.EndUpdate();
            //    schedulerControl1.RefreshData();
            //    schedulerControl2.EndUpdate();
            //    schedulerControl2.RefreshData();

            //    gridView3.BeginUpdate();
            //    globalProject.LatestFinishDate = globalProject.GetLatestFinishDate();
            //    gridView3.EndUpdate();
            //}

            return false;
        }
        private void projectComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                LoadProjectGantt();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void resourcesTree1_CustomDrawNodeCell(object sender, DevExpress.XtraTreeList.CustomDrawNodeCellEventArgs e)
        {
            //if (schedulerControl2.Storage != null)
            //{
            //    Resource resource = schedulerControl2.Storage.Resources.Items.Find(r => r.Caption == e.CellText);
            //    e.Appearance.BackColor = resource.GetColor();
            //}
        }
        private void schedulerControl2_DragOver(object sender, DragEventArgs e)
        {
            Point pos = schedulerControl2.PointToClient(Cursor.Position);
            SchedulerViewInfoBase viewInfo = schedulerControl2.ActiveView.ViewInfo;
            HitInfo = viewInfo.CalcHitInfo(pos, false);
        }
        private void schedulerControl2_MouseDown(object sender, MouseEventArgs e)
        {
            var scheduler = sender as DevExpress.XtraScheduler.SchedulerControl;
            var hitInfo = scheduler.ActiveView.CalcHitInfo(e.Location, false);

            if (e.Button == MouseButtons.Right)
            {
                RightMouseButtonPressed = true;

                if (hitInfo.HitTest == SchedulerHitTest.AppointmentContent)
                {
                    Appointment apt = ((AppointmentViewInfo)hitInfo.ViewInfo).Appointment;
                    DraggedAppointment = apt;
                }
            }
        }
        private void schedulerControl2_PopupMenuShowing(object sender, DevExpress.XtraScheduler.PopupMenuShowingEventArgs e)
        {
            e.Menu.Items.Remove(e.Menu.Items.FirstOrDefault(x => x.Caption == "&Copy"));

            if (e.Menu.Items.Count(x => x.Caption == "Mo&ve") > 0)
            {
                DXMenuItem menuItem = e.Menu.Items.FirstOrDefault(x => x.Caption == "Mo&ve");
                menuItem.Caption = "Move All Component Tasks with Locked Spacing";
            }

            if (e.Menu.Items.Count(x => x.Caption == "Move Subsequent Component Tasks with Lock Spacing") == 0)
            {
                e.Menu.Items.Insert(1, new SchedulerMenuItem("Move Subsequent Component Tasks with Lock Spacing", schedulerStorage2_MoveSubsequentComponentTasksWithLockedSpacing));
            }
        }
        private void schedulerStorage2_MoveSubsequentComponentTasksWithLockedSpacing(object sender, EventArgs e)
        {
            MoveSubsequentTaskWithLockedSpacing = true;

            //MessageBox.Show($"Final Start: {HitInfo.ViewInfo.Interval.Start.ToShortDateString()} Finish: {HitInfo.ViewInfo.Interval.End.ToShortDateString()}");

            DraggedAppointment.Start = HitInfo.ViewInfo.Interval.Start;
            DraggedAppointment.End = HitInfo.ViewInfo.Interval.End;
        }
        private void schedulerStorage2_AppointmentChanging(object sender, PersistentObjectCancelEventArgs e)
        {
            if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeDates))
            {
                MessageBox.Show("This login is not authorized to make changes to dates.");
                e.Cancel = true;
            }
            // Have this if statement handle changes to the database. But not yet.
            //if (true)
            //{

            //}
            //else
            //{
            //    e.Cancel = true;
            //}
            if (RightMouseButtonPressed)
            {
                TaskModel changingTask = ((Appointment)e.Object).GetSourceObject(schedulerStorage2) as TaskModel;

                OldTaskStartDate = (DateTime)changingTask.StartDate;  // (DateTime)changingTask.StartDate

                Console.WriteLine($"Old Date: {OldTaskStartDate}");
                Console.WriteLine();
            }
        }
        private void schedulerStorage2_AppointmentsChanged(object sender, PersistentObjectsEventArgs e)
        {
            List<int> collapsedNodes = new List<int>();

            TaskModel movedTask;

            try
            {
                foreach (Appointment apt in e.Objects)
                {
                    movedTask = apt.GetSourceObject(schedulerStorage2) as TaskModel;

                    if (UpdateTaskStorage2(movedTask))
                    {
                        var number = GetProjectComboBoxInfo(projectComboBox);
                        collapsedNodes = GetCollapsedNodes();
                        LoadProject(number.projectNumber);
                        schedulerControl2.RefreshData();
                        CollapseNodes(collapsedNodes);
                    }
                    else
                    {
                        var number = GetProjectComboBoxInfo(projectComboBox);
                        collapsedNodes = GetCollapsedNodes();
                        LoadProject(number.projectNumber);
                        schedulerControl2.RefreshData();
                        CollapseNodes(collapsedNodes);
                    }
                    //MessageBox.Show(apt.Subject);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message} \n\n {ex.StackTrace}");
            }
        }
        private void schedulerControl2_AppointmentFlyoutShowing(object sender, AppointmentFlyoutShowingEventArgs e)
        {
            TaskModel hoveredTask = new TaskModel();

            hoveredTask.JobNumber = e.FlyoutData.Appointment.CustomFields["JobNumber"].ToString();
            hoveredTask.ProjectNumber = (int)e.FlyoutData.Appointment.CustomFields["ProjectNumber"];
            hoveredTask.Component = e.FlyoutData.Appointment.CustomFields["Component"].ToString();
            hoveredTask.TaskName = e.FlyoutData.Appointment.CustomFields["TaskName"].ToString();
            hoveredTask.Hours = (int)e.FlyoutData.Appointment.CustomFields["Hours"];
            hoveredTask.Notes = e.FlyoutData.Appointment.Description;

            e.Control = CreateLabel(hoveredTask);
        }

        private void RefreshGanttButton_Click(object sender, EventArgs e)
        {
            LoadProjects();
            PopulateProjectComboBox();

            try
            {
                if (projectComboBox.Text != "")
                {
                    List<int> collapsedNodes = new List<int>();
                    collapsedNodes = GetCollapsedNodes();
                    LoadProjectGantt();
                    CollapseNodes(collapsedNodes);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        #endregion

        #region Calender

        private void PopulateProjectComboBox2()
        {
            projectComboBox2.Properties.Items.Clear();

            projectComboBox2.Properties.Items.AddRange(Database.GetJobNumberComboList());
        }

        private void PopulateComponentComboBox()
        {
            if (componentComboBox.Properties.Items.Count != 0 && componentComboBox.Properties.Items != null)
            {
                componentComboBox.Properties.Items.Clear();
            }

            componentComboBox.Properties.Items.AddRange(CalendarProject.Components.Select(x => x.Component).ToArray());

            componentComboBox.Properties.Items.Add("");
        }

        private void LoadProjectCalendar(string comboBoxName)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));

            calendarDataStorage.Appointments.CustomFieldMappings.Clear();

            calendarDataStorage.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("JobNumber", "JobNumber"));
            calendarDataStorage.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("ProjectNumber", "ProjectNumber"));
            calendarDataStorage.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskID", "TaskID"));
            calendarDataStorage.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskName", "TaskName"));
            calendarDataStorage.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Component", "Component"));
            calendarDataStorage.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Hours", "Hours"));
            calendarDataStorage.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Predecessors", "Predecessors"));

            // This sets the AptID for all the tasks in a project.
            AptResourceModel.GetProjectResourceData(CalendarProject);

            AppointmentMappingInfo appointmentMappings = calendarDataStorage.Appointments.Mappings;

            appointmentMappings.AppointmentId = "AptID";
            appointmentMappings.Start = "StartDate";
            appointmentMappings.End = "FinishDate";
            appointmentMappings.Subject = "Component";
            appointmentMappings.Location = "Location";
            appointmentMappings.Description = "Notes";
            appointmentMappings.PercentComplete = "PercentComplete";
            //appointmentMappings.ResourceId = "AptID";

            List<TaskModel> rawTaskList;

            if (comboBoxName == "projectComboBox2" || componentComboBox.Text == "")
            {
                rawTaskList = CalendarProject.GetTaskList();
            }
            else
            {
                rawTaskList = CalendarProject.Components.Find(x => x.Component == componentComboBox.Text).Tasks;
            }

            BindingList<TaskModel> taskList = new BindingList<TaskModel>(rawTaskList);

            calendarDataStorage.Appointments.DataSource = taskList;

            SplashScreenManager.CloseForm();
        }

        private void projectComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            var number = GetProjectComboBoxInfo(projectComboBox2);
            var control = sender as ComboBoxEdit;

            CalendarProject = Database.GetProject(number.projectNumber);

            PopulateComponentComboBox();

            LoadProjectCalendar(control.Name);
        }

        private void componentComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var control = sender as ComboBoxEdit;

            if (projectComboBox2.Text != "")
            {
                LoadProjectCalendar(control.Name);
            }
        }
        private void RefreshCalendarButton_Click(object sender, EventArgs e)
        {
            PopulateProjectComboBox2();

            try
            {
                if (projectComboBox2.Text != "")
                {
                    LoadProjectCalendar(sender.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void PrintCalendarButton_Click(object sender, EventArgs e)
        {
            ShowSchedulerPreview(schedulerControl3);
        }
        private void schedulerControl3_AppointmentFlyoutShowing(object sender, AppointmentFlyoutShowingEventArgs e)
        {
            TaskModel hoveredTask = new TaskModel();

            hoveredTask.JobNumber = e.FlyoutData.Appointment.CustomFields["JobNumber"].ToString();
            hoveredTask.ProjectNumber = (int)e.FlyoutData.Appointment.CustomFields["ProjectNumber"];
            hoveredTask.Component = e.FlyoutData.Appointment.CustomFields["Component"].ToString();
            hoveredTask.TaskName = e.FlyoutData.Appointment.CustomFields["TaskName"].ToString();
            hoveredTask.Hours = (int)e.FlyoutData.Appointment.CustomFields["Hours"];
            hoveredTask.Notes = e.FlyoutData.Appointment.Description;
            hoveredTask.ComponentPicture = ComponentsList.Find(x => x.ProjectNumber == hoveredTask.ProjectNumber && x.Component == hoveredTask.Component).picture;

            e.Control = CreateLabel(hoveredTask);
        }

        private void calendarDataStorage_AppointmentChanging(object sender, PersistentObjectCancelEventArgs e)
        {
            MessageBox.Show("Calendar View is read-only.  No changes permitted.");

            e.Cancel = true;
        }

        private void ShowSchedulerPreview(SchedulerControl scheduler)
        {
            // 
            // Check whether the SchedulerControl can be previewed.
            if (!scheduler.IsPrintingAvailable)
            {
                MessageBox.Show("The 'DevExpress.XtraPrinting.vX.Y.dll' is not found", "Error");
                return;
            }

            //MonthlyPrintStyle style = scheduler.PrintStyles[SchedulerPrintStyleKind.Monthly] as MonthlyPrintStyle;

            //style.CalendarHeaderVisible = false;
            //style.PageSettings.Landscape = true;
            //style.PageSettings.Margins.Top = 25;
            //style.PageSettings.Margins.Bottom = 25;
            //style.PageSettings.Margins.Right = 25;
            //style.PageSettings.Margins.Left = 25;
            //style.CompressWeekend = false;

            // Open the Preview window.
            //scheduler.ShowPrintPreview();

            MonthlyXtraSchedulerReport monthlyXtraSchedulerReport = new MonthlyXtraSchedulerReport(CalendarProject, componentComboBox.Text);

            SchedulerControlPrintAdapter scPrintAdapter =
                new SchedulerControlPrintAdapter(this.schedulerControl3);
            monthlyXtraSchedulerReport.SchedulerAdapter = scPrintAdapter;

            monthlyXtraSchedulerReport.CreateDocument(true);

            ReportPrintTool printTool = new ReportPrintTool(monthlyXtraSchedulerReport);
            printTool.Report.CreateDocument(true);
            printTool.ShowPreviewDialog();
        }

        #endregion

        private List<string> GetResourceList(string role, string resourceType)
        {
            List<string> resourceList = new List<string>();

            if (role != "")
            {
                var result = from roleTable in RoleTable.AsEnumerable()
                             where roleTable.Field<string>("Role") == role
                             select roleTable;

                foreach (var resource in result)
                {
                    resourceList.Add(resource.Field<string>("ResourceName"));
                }
            }
            else if (role == "")
            {
                var result2 = from roleTable in RoleTable.AsEnumerable()
                              where roleTable.Field<string>("ResourceType") == resourceType
                              group roleTable by roleTable.Field<string>("ResourceName") into grp
                              orderby grp.Key
                              select grp;

                foreach (var resource in result2)
                {
                    resourceList.Add(resource.Key);
                }
            }

            if (resourceType == "Person")
            {
                resourceList.Add("");
            }

            return resourceList;
        }
        private (string jobNumber, int projectNumber) GetProjectComboBoxInfo(ComboBoxEdit comboBox)
        {
            string[] jobNumberComboBoxText, jobNumberComboBoxText2;

            jobNumberComboBoxText = comboBox.Text.Split(' ');
            jobNumberComboBoxText2 = comboBox.Text.Split('#');

            return (jobNumberComboBoxText[0], Convert.ToInt32(jobNumberComboBoxText2[1]));
        }
        private Label CreateLabel(TaskModel task)
        {
            int extraRow = 0, pictureWidth = 0, pictureHeight = 0, maxTextPixelWidth, remainingWidth;
            Label myControl = new Label();
            List<string> textList = new List<string>();
            myControl.BackColor = Color.LightGreen;
            myControl.Size = new Size(400, 155);

            textList.Add(task.Component);
            textList.Add(task.Notes);

            myControl.Text = $"    Job#: {task.JobNumber}{Environment.NewLine}" +
                             $"   Proj#: {task.ProjectNumber}{Environment.NewLine}" +
                             $"    Comp: {task.Component}{Environment.NewLine}" +
                             $"    Task: {task.TaskName}{Environment.NewLine}" +
                             $"     Hrs: {task.Hours}{Environment.NewLine}" +
                             $"   Notes: {task.Notes}{Environment.NewLine}" +
                             $"Due Date: {task.DueDate:M-d-yy}{Environment.NewLine}";

            myControl.Font = new Font("Lucida Sans Typewriter", 12, FontStyle.Bold);

            if (task.ComponentPicture != null)
            {
                pictureWidth = task.ComponentPicture.Width;
                pictureHeight = task.ComponentPicture.Height;
                myControl.Image = task.ComponentPicture;
                myControl.ImageAlign = ContentAlignment.BottomLeft;
            }

            if (pictureWidth > myControl.Width)
            {
                myControl.Width = pictureWidth + 5;
            }

            maxTextPixelWidth = textList.Max(x => x.Length) * 10;

            if (maxTextPixelWidth + 99 > myControl.Width)
            {
                remainingWidth = maxTextPixelWidth + 99 - myControl.Width;
                extraRow = 19 * (remainingWidth / myControl.Width);

                if (remainingWidth % myControl.Width > 0)
                {
                    extraRow += 19;
                }
            }

            myControl.Height = (19 * 7) + extraRow + pictureHeight;

            return myControl;
        }
        private void PopulateDepartmentComboBoxes()
        {
            List<string> departmentList1 = new List<string> { "Design", "Programming", "Program Rough", "Program Finish", "Program Electrodes", "CNCs", "CNC People", "CNC Rough", "CNC Finish", "CNC Electrodes", "Grind", "Inspection", "EDM Sinker", "EDM Wire (In-House)", "Mold Service", "Polish", "Manual", "All" };
            List<string> departmentList2 = new List<string> { "Design", "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "CNC Electrodes", "Grind", "Inspection", "EDM Sinker", "EDM Wire (In-House)", "Mold Service", "Polish", "Manual", "All" };

            departmentComboBox.Properties.Items.AddRange(departmentList1);
            departmentComboBox2.Properties.Items.AddRange(departmentList2);
        }
    }
}
