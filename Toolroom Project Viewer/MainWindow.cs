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
using DevExpress.XtraRichEdit.API.Native;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using ClassLibrary;
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
using DevExpress.Data;

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
        private string printOrientation, paperSize;
        private string[] departmentArr = { "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "CNC Electrodes", "Grind", "Inspection", "EDM Sinker", "EDM Wire (In-House)", "Polish" };
        public DataTable RoleTable { get; set; }
        private string TimeUnits { get; set; }
        private ProjectModel Project { get; set; }
        private List<ProjectModel> ProjectInfoList { get; set; }

        private List<int> gridView3ExpandedRowsList = new List<int>();
        private List<int> gridView4ExpandedRowsList = new List<int>();
        private List<ExpandedProjectRows> expandedProjectRowsList = new List<ExpandedProjectRows>();
        private List<int> gridView3SelectedRows = new List<int>();
        private List<int> gridView4SelectedRows = new List<int>();
        private List<int> gridView5SelectedRows = new List<int>();
        private DataTable ResourceDataTable;
        private string Role, Tasks;
        Regex TaskRegExpression, RoleRegExpression;
        private bool AllProjectItemsChecked;
        //private ArrayList saveExpList;

        private RefreshHelper helper1, helper2, deptTaskViewHelper;

        public MainWindow()
        {
            try
            {
                this.TimeUnits = "Days";
                ResourceDataTable = Database.GetResourceData();
                InitializeComponent();
                SetRole();
                SetTasks();
                PopulateProjectCheckedComboBox();
                InitializeResources();
                InitializeAppointments();
                GroupByRadioGroup.SelectedIndex = 0;
                chartRadioGroup.SelectedIndex = 0;
                InitializePrintOptions();
                schedulerControl1.Start = DateTime.Today.AddDays(-7);
                schedulerControl1.OptionsCustomization.AllowAppointmentDelete = UsedAppointmentType.Custom;
                schedulerControl1.AllowAppointmentDelete += new AppointmentOperationEventHandler(schedulerControl1_AllowAppointmentDelete);

                PopulateDepartmentComboBoxes();
                PopulateProjectComboBox();
                PopulateTimeFrameComboBox();               

                schedulerStorage2.Appointments.CommitIdToDataSource = false;

                schedulerControl2.Start = DateTime.Today.AddDays(-7);
                schedulerControl2.Views.GanttView.ResourcesPerPage = 15;
                schedulerControl2.GroupType = SchedulerGroupType.Resource;
                schedulerControl2.ActiveViewType = SchedulerViewType.Gantt;
                //InitializeExample();
                gridView1.ActiveFilterCriteria = FilterTaskView(departmentComboBox2.Text, false, false);
                AddRepositoryItemToGrid();
                gridView3.DetailHeight = int.MaxValue;
                gridView4.DetailHeight = int.MaxValue;

                AddVersionNumber();

                CheckForUpdates();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
            }
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            try
            {
                helper1 = new RefreshHelper(gridView3, "JobNumber");
                helper2 = new RefreshHelper(gridView4, "Component");
                // TODO: This line of code loads data into the 'workload_Tracking_System_DBDataSet.Components' table. You can move, or remove it, as needed.
                this.componentsTableAdapter.Fill(this.workload_Tracking_System_DBDataSet.Components);
                // TODO: This line of code loads data into the 'workload_Tracking_System_DBDataSet.Projects' table. You can move, or remove it, as needed.
                this.projectsTableAdapter.Fill(this.workload_Tracking_System_DBDataSet.Projects);
                // TODO: This line of code loads data into the 'workload_Tracking_System_DBDataSet.Tasks' table. You can move, or remove it, as needed.
                this.tasksTableAdapter.Fill(this.workload_Tracking_System_DBDataSet.Tasks);
                // TODO: This line of code loads data into the 'workload_Tracking_System_DBDataSet.WorkLoad' table. You can move, or remove it, as needed.
                this.workLoadTableAdapter.Fill(this.workload_Tracking_System_DBDataSet.WorkLoad);

                gridControl2.DataSource = workLoadTableAdapter.GetData();
                footerDateTime = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

                schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskID", "TaskID"));
                schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Component", "Component"));
                RoleTable = Database.GetRoleTable();
                ProjectInfoList = Database.GetProjectInfoList();
                PopulateEmployeeComboBox();
                //gridView3.Columns["IncludeHours"].VisibleIndex = 14;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
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

        #region Department Schedule View

        private void InitializeExample()
        {

            //SchedulerStorage schedulerStorage = new SchedulerStorage();
            AppointmentStorage appointmentStorage = new AppointmentStorage(schedulerStorage1);

            //schedulerStorage.CreateAppointment(AppointmentType.Normal, DateTime.Today, DateTime.Today.AddDays(1), "First Appointment");
            // DataStorage can be substituted for Storage property.

            //schedulerControl1.DataStorage = schedulerStorage;
            //schedulerControl1.Refresh();

            //schedulerStorage1.CreateAppointment(AppointmentType.Normal, DateTime.Today, DateTime.Today.AddDays(1), "First Appointment");
            //schedulerControl1.Refresh();

            //schedulerControl1.DataStorage.Appointments.Add()

            //schedulerControl1.DataStorage.CreateAppointment(AppointmentType.Normal, DateTime.Today, DateTime.Today.AddDays(1));
            //schedulerControl1.Refresh();

            //Appointment apt = schedulerControl1.DataStorage.CreateAppointment(AppointmentType.Normal);
            //apt.Start = DateTime.Today;
            //apt.End = DateTime.Today.AddDays(1);
            //apt.Subject = "First appointment";

            //schedulerControl1.DataStorage.Appointments.Add(apt);

            //Appointment apt2 = schedulerControl1.DataStorage.CreateAppointment(AppointmentType.Normal, DateTime.Today, DateTime.Today.AddDays(2), "Second Appointment");
            //schedulerControl1.DataStorage.Appointments.Add(apt2);

            // How do I create a bunch of appointments from a datatable?

            //appointmentStorage.Add(schedulerControl1.DataStorage.CreateAppointment(AppointmentType.Normal, DateTime.Today, DateTime.Today.AddDays(2), "Second Appointment"));

            //schedulerControl1.DataStorage.Appointments.Add(schedulerControl1.DataStorage.CreateAppointment(AppointmentType.Normal, DateTime.Today, DateTime.Today.AddDays(2), "Second Appointment"));

            Database db = new Database();

            schedulerStorage1.Appointments.ResourceSharing = true;

            AppointmentMappingInfo appointmentMappings = schedulerStorage1.Appointments.Mappings;

            appointmentMappings.Start = "StartDate";
            appointmentMappings.End = "FinishDate";
            appointmentMappings.Subject = "Subject";
            appointmentMappings.Location = "Machine";

            //appointmentMappings.ResourceId = {"ToolMaker"};

            schedulerStorage1.Appointments.DataSource = db.GetAppointmentData();

            ResourceIdCollection resourceIdCollection = new ResourceIdCollection();

            resourceIdCollection.Add(schedulerStorage1.Resources[0].Id);
            resourceIdCollection.Add(schedulerStorage1.Resources[1].Id);

            foreach (Resource item in schedulerStorage1.Resources.Items)
            {
                // Displays whatever is specified as the resource id such as the database id or the combined first and last names.
                Console.WriteLine(item.Id);
            }

            //foreach (Resource item in resourceIdCollection)
            //{
            //    Console.WriteLine(item.Id);
            //}

            AppointmentResourceIdCollectionContextElement multi_resource = new AppointmentResourceIdCollectionContextElement(resourceIdCollection);
            appointmentMappings.ResourceId = multi_resource.ValueToString();

            // Displays xml code ids for resources.
            Console.WriteLine(multi_resource.ValueToString());

            foreach (Appointment apt in schedulerStorage1.Appointments.Items)
            {
                apt.ResourceId = multi_resource.ValueToString();

                //foreach (string id in apt.ResourceIds)
                //{
                //    // Display's xml code ids.
                //    Console.WriteLine(id);
                //}
            }

            schedulerControl1.GroupType = SchedulerGroupType.Resource;
        }

        private void InitializeResources()
        {
            Database db = new Database();
            DataTable dt = new DataTable();

            schedulerStorage1.Resources.Clear();

            ResourceStorage resourceStorage = new ResourceStorage(schedulerStorage1);
            ResourceMappingInfo resourceMappings = schedulerStorage1.Resources.Mappings;

            resourceMappings.Caption = "ResourceName";
            resourceMappings.Id = "ResourceName";

            //if (departmentComboBox.Text == "Programming")
            //{
            //    dt = db.GetResourceData().AsEnumerable().Where(x => x.Field<string>("Role").Contains("Programmer")).GroupBy(x => x.Field<string>("ResourceName")).Select(x => x.FirstOrDefault()).CopyToDataTable();
            //}
            //else if(departmentComboBox.Text == "CNC")
            //{
            //    dt = db.GetResourceData().AsEnumerable().Where(x => x.Field<string>("Role").Contains("Mill")).GroupBy(x => x.Field<string>("ResourceName")).Select(x => x.FirstOrDefault()).CopyToDataTable();
            //}
            //else
            //{
            //    dt = db.GetResourceData().AsEnumerable().Where(x => x.Field<string>("Department") == departmentComboBox.Text &&
            //                              x.Field<string>("Role").Contains(GetRoleFromDepartmentName(departmentComboBox.Text))).CopyToDataTable();
            //}

            //dt = db.GetResourceData();

            //foreach (DataRow nrow in dt.Rows)
            //{
            //    Console.WriteLine($"{nrow["ID"]} {nrow["ResourceName"]} {nrow["Role"]} {nrow["Department"]}");
            //}

            schedulerStorage1.Resources.DataSource = ResourceDataTable.AsEnumerable().GroupBy(x => x.Field<string>("ResourceName")).Select(x => x.FirstOrDefault()).CopyToDataTable();

            Console.WriteLine();
            Console.WriteLine("Resources");

            //for (int i = 0; i < schedulerStorage1.Resources.Items.Count; i++)
            //{
            //    Console.WriteLine($"{schedulerStorage1.Resources[i].Id} {schedulerStorage1.Resources[i].Caption}");
            //}
        }

        private string GetRoleFromDepartmentName(string department)
        {
            if (department.Contains("CNC"))
            {
                return "Mill";
            }
            else if (department.Contains("Program"))
            {
                return "Programmer";
            }
            else
            {
                return "";
            }
        }

        private void InitializeAppointments()
        {
            Database db = new Database();
            //bool grouped;

            schedulerStorage1.Appointments.Clear();
            schedulerStorage1.Appointments.CustomFieldMappings.Clear();

            schedulerStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("JobNumber", "JobNumber"));
            schedulerStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("ProjectNumber", "ProjectNumber"));
            schedulerStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskID", "TaskID"));
            schedulerStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskName", "TaskName"));
            schedulerStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Component", "Component"));
            schedulerStorage1.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("Hours", "Hours"));

            AppointmentMappingInfo appointmentMappings = schedulerStorage1.Appointments.Mappings;

            appointmentMappings.AppointmentId = "ID";
            appointmentMappings.Start = "StartDate";
            appointmentMappings.End = "FinishDate";
            appointmentMappings.Subject = "Subject";
            appointmentMappings.Location = "Location";
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

            schedulerStorage1.Appointments.DataSource = db.GetAppointmentData("All");

            //for (int i = 0; i < schedulerStorage1.Appointments.Items.Count; i++)
            //{
            //    Console.WriteLine($"{schedulerStorage1.Appointments[i].Id} {schedulerStorage1.Appointments[i].Subject} Resource: {schedulerStorage1.Appointments[i].ResourceId}");
            //}

            Console.WriteLine();
            Console.WriteLine($"{departmentComboBox.Text} Appointments");

            //if (GroupByRadioGroup.SelectedIndex == 1)
            //{
            //    this.schedulerStorage1.AppointmentsChanged -= new DevExpress.XtraScheduler.PersistentObjectsEventHandler(this.schedulerStorage1_AppointmentsChanged);

            //    for (int i = 0; i < schedulerStorage1.Appointments.Items.Count; i++)
            //    {
            //        schedulerStorage1.Appointments[i].ResourceId = null;
            //    }

            //    this.schedulerStorage1.AppointmentsChanged += new DevExpress.XtraScheduler.PersistentObjectsEventHandler(this.schedulerStorage1_AppointmentsChanged);
            //}
        }

        private void GetAppointments()
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("TaskName", typeof(string));
            dt.Columns.Add("StartDate", typeof(DateTime));
            dt.Columns.Add("FinishDate", typeof(DateTime));

            foreach (Appointment apt in schedulerStorage1.Appointments.Items)
            {
                dt.Rows.Add(apt.Id, apt.Subject, apt.Start, apt.End);
            }

            foreach (DataRow nrow in dt.Rows)
            {
                Console.WriteLine($"{nrow["ID"].ToString()} {nrow["TaskName"].ToString()} {nrow["StartDate"]} {nrow["FinishDate"]}");
            }
        }

        private bool UpdateTaskStorage1(Appointment apt)
        {
            Database db = new Database();
            int projectNumber, taskID;
            string jobNumber, component, taskName, resourceIDs, machine, resource;

            jobNumber = apt.CustomFields["JobNumber"].ToString();
            projectNumber = Convert.ToInt32(apt.CustomFields["ProjectNumber"]);
            component = apt.CustomFields["Component"].ToString();
            taskID = Convert.ToInt32(apt.CustomFields["TaskID"]);
            taskName = apt.CustomFields["TaskName"].ToString();
            resourceIDs = GenerateResourceIDsString(apt.ResourceIds);
            machine = GetMachineFromResourceIDs(apt.ResourceIds);
            resource = GetResourceFromResourceIDs(apt.ResourceIds);

            if (db.UpdateTask(jobNumber, projectNumber, component, taskID, apt.Start, apt.End, machine, resource, resourceIDs))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void SetAppointmentResources(int id, string machine, string resource)
        {
            Appointment apt = schedulerStorage1.Appointments.GetAppointmentById(id);
            Resource res;

            apt.ResourceIds.Clear();

            int machineCount = schedulerStorage1.Resources.Items.Where(x => x.Id.ToString() == machine).Count();
            int resourceCount = schedulerStorage1.Resources.Items.Where(x => x.Id.ToString() == resource).Count();

            if (machineCount == 0 && resourceCount == 0)
            {
                res = schedulerStorage1.Resources.Items.GetResourceById("None");
                apt.ResourceIds.Add(res.Id);
            }

            if (machine != "" && machineCount == 1)
            {
                res = schedulerStorage1.Resources.Items.GetResourceById(machine);
                apt.ResourceIds.Add(res.Id);
            }

            if (resource != "" && resourceCount == 1)
            {
                res = schedulerStorage1.Resources.Items.GetResourceById(resource);
                apt.ResourceIds.Add(res.Id);
            }
        }

        private void SetAppointmentResources(Appointment apt, string machine, string resource)
        {
            Resource res;

            apt.ResourceIds.Clear();

            int machineCount = schedulerStorage1.Resources.Items.Where(x => x.Id.ToString() == machine).Count();
            int resourceCount = schedulerStorage1.Resources.Items.Where(x => x.Id.ToString() == resource).Count();

            if (machineCount == 0 && resourceCount == 0)
            {
                res = schedulerStorage1.Resources.Items.GetResourceById("None");
                apt.ResourceIds.Add(res.Id);
            }

            if (machine != "" && machineCount == 1)
            {
                res = schedulerStorage1.Resources.Items.GetResourceById(machine);
                apt.ResourceIds.Add(res.Id);
            }

            if (resource != "" && resourceCount == 1)
            {
                res = schedulerStorage1.Resources.Items.GetResourceById(resource);
                apt.ResourceIds.Add(res.Id);
            }
        }

        private void SetAppointmentResources(object sender, CellValueChangedEventArgs e)
        {
            var grid = (sender as DevExpress.XtraGrid.Views.Grid.GridView);
            int projectNumber = (int)grid.GetRowCellValue(e.RowHandle, "ProjectNumber");
            string taskName;
            Database db = new Database();
            DataTable dt = new DataTable();

            taskName = "Program " + e.Column.FieldName.Remove(e.Column.FieldName.Length - 10, 10);

            if (taskName.Contains("Electrode"))
            {
                taskName = taskName + "s";
            }

            dt = db.GetTasksWithChangedResources(projectNumber, taskName);

            foreach (DataRow nrow in dt.Rows)
            {
                nrow["Resources"] = GenerateResourceIDsString(nrow["Machine"].ToString(), nrow["Resource"].ToString());
            }

            var apts = schedulerStorage1.Appointments.Items.Where(x => (int)x.CustomFields["ProjectNumber"] == projectNumber && x.CustomFields["TaskName"].ToString() == taskName);

            foreach (Appointment apt in apts.ToList())
            {
                Console.WriteLine($"{apt.CustomFields["JobNumber"]} {apt.CustomFields["ProjectNumber"]} {apt.CustomFields["TaskName"]} {apt.CustomFields["Hours"]}");
                DataRow nrow = dt.AsEnumerable().First(x => x.Field<int>("TaskID") == (int)apt.CustomFields["TaskID"]);
                SetAppointmentResources(apt, nrow.Field<string>("Machine"), nrow.Field<string>("Resource"));
            }
        }

        /// <summary>
        /// Gets the last selected machine in resource list.
        /// </summary>
        private string GetMachineFromResourceIDs(AppointmentResourceIdCollection appointmentResourceIdCollection)
        {
            string id = "";
            foreach (var item in appointmentResourceIdCollection)
            {
                Console.WriteLine($"Resource: {item.ToString()}");
                // This just validates that the selected resource is a machine and not a person.  It assumes that the resource list is comprised of both people and machines.
                if (ResourceDataTable.AsEnumerable().Where(x => x.Field<string>("ResourceName") == item.ToString() && x.Field<string>("ResourceType") == "Machine").Count() >= 1)
                {
                    id = item.ToString(); 
                }
            }

            return id;
        }
        /// <summary>
        /// Gets the last selected person in resource list.
        /// </summary>
        private string GetResourceFromResourceIDs(AppointmentResourceIdCollection appointmentResourceIdCollection)
        {
            string id = "";
            foreach (var item in appointmentResourceIdCollection)
            {
                // This just validates that the selected resource is a person and not a machine.  It assumes that the resource list is comprised of both people and machines.
                if (ResourceDataTable.AsEnumerable().Where(x => x.Field<string>("ResourceName") == item.ToString() && x.Field<string>("ResourceType") == "Person").Count() >= 1)
                {
                    id = item.ToString();
                }
            }

            return id;
        }

        private string GenerateResourceIDsString(string machine, string resource)
        {
            AppointmentResourceIdCollection appointmentResourceIdCollection = new AppointmentResourceIdCollection();
            Resource res;
            int machineCount = schedulerStorage1.Resources.Items.Where(x => x.Id.ToString() == machine).Count();
            int resourceCount = schedulerStorage1.Resources.Items.Where(x => x.Id.ToString() == resource).Count();

            if (machineCount == 0 && resourceCount == 0)
            {
                res = schedulerStorage1.Resources.Items.GetResourceById("None");
                appointmentResourceIdCollection.Add(res.Id);
            }

            if (machine != "" && machineCount == 1)
            {
                res = schedulerStorage1.Resources.Items.GetResourceById(machine);
                appointmentResourceIdCollection.Add(res.Id);
            }

            if (resource != "" && resourceCount == 1)
            {
                res = schedulerStorage1.Resources.Items.GetResourceById(resource);
                appointmentResourceIdCollection.Add(res.Id);
            }
            
            AppointmentResourceIdCollectionXmlPersistenceHelper helper = new AppointmentResourceIdCollectionXmlPersistenceHelper(appointmentResourceIdCollection);
            return helper.ToXml();
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

        private void SetTasks()
        {
            string department = departmentComboBox.Text;

            if (department == "Program Rough")
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
                Tasks = @"^(CNC Rough)\w*";
            }
            else if (department == "Rough")
            {
                Tasks = @"Rough";
            }
            else if (department == "CNC Finish")
            {
                Tasks = @"^(CNC Finish)\w*";
            }
            else if (department == "CNC Electrodes")
            {
                Tasks = @"^(CNC Electrodes)";
            }
            else if (department == "CNC")
            {
                Tasks = @"^CNC\w*";
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
                Tasks = @"\w*Grind\w*";
            }
            else if (department == "Polish")
            {
                Tasks = @"Polish \(In-House\)\w*";
            }
            else if (department == "Inspection")
            {
                Tasks = @"^Inspection\w*";
            }
            else if (department == "All")
            {
                Tasks = "All";
            }

            TaskRegExpression = new Regex(Tasks);
        }

        private void SetRole()
        {
            string department = departmentComboBox.Text;
            schedulerControl1.ActiveView.ResourcesPerPage = 0;

            if (department == "Program Rough")
            {
                Role = "Rough Programmer";
            }
            else if (department == "Program Finish")
            {
                Role = "Finish Programmer";
            }
            else if (department == "Program Electrodes")
            {
                Role = "Electrode Programmer";
            }
            else if (department == "Programming")
            {
                Role = "Programmer";
            }
            else if (department == "CNC Rough")
            {
                Role = "Rough Mill";
            }
            else if (department == "Rough")
            {
                Role = "Rough";
            }
            else if (department == "CNC Finish")
            {
                Role = "Finish Mill";
            }
            else if (department == "CNC Electrodes")
            {
                Role = "Graphite Mill";
            }
            else if (department == "CNCs")
            {
                Role = "Mill";
            }
            else if (department == "CNC People")
            {
                Role = "CNC Operator";
            }
            else if (department == "EDM Sinker")
            {
                Role = @"^(EDM Sinker)$";
            }
            else if (department == "EDM Wire")
            {
                Role = @"^EDM Wire$";
            }
            else if (department == "Grind")
            {
                Role = "Tool Maker";
            }
            else if (department == "Polish")
            {
                Role = "Tool Maker";
            }
            else if (department == "Inspection")
            {
                Role = "CMM Operator";
            }
            else if (department == "All")
            {
                Role = "All";
                schedulerControl1.ActiveView.ResourcesPerPage = 8;
                //return;
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
            InitializeAppointments();
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
                RefreshDepartmentScheduleView();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
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

        private void schedulerControl1_DragDrop(object sender, DragEventArgs e)
        {
            //MessageBox.Show("DragDrop");
        }

        private void schedulerControl1_AppointmentResized(object sender, AppointmentResizeEventArgs e)
        {
            //MessageBox.Show("Resize");
        }

        private void schedulerControl1_AllowAppointmentDelete(object sender, AppointmentOperationEventArgs e)
        {
            e.Allow = false;
        }

        private void schedulerStorage1_AppointmentsChanged(object sender, PersistentObjectsEventArgs e)
        {
            foreach (Appointment apt in e.Objects)
            {
                if (UpdateTaskStorage1(apt))
                {

                }
                else
                {
                    InitializeAppointments();
                    schedulerControl1.RefreshData();
                }
                //MessageBox.Show(apt.Subject);
            }

            //updateAppointment((Appointment)e.Objects);
            //getAppointments();
            //MessageBox.Show("AppointmentChanged");
        }

        private void schedulerStorage1_FilterAppointment(object sender, PersistentObjectCancelEventArgs e)
        {
            Appointment apt = (Appointment)e.Object;

            if (Tasks != "All")
            {
                if (AllProjectItemsChecked == false)
                {
                    e.Cancel = !projectCheckedComboBoxEdit.Properties.Items.Where(x => x.Value.ToString().Contains($"#{apt.CustomFields["ProjectNumber"]}") && x.CheckState == CheckState.Checked && TaskRegExpression.IsMatch(apt.Location)).Any(); // 
                }
                else
                {
                    e.Cancel = !TaskRegExpression.IsMatch(apt.Location);
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
                    if (apt.CustomFields["JobNumber"].ToString().Contains("quote"))
                    {
                        e.Cancel = true;
                    }
                }

                //e.Cancel = !apt.Location.Contains(Tasks);
            }
        }

        private void schedulerStorage1_FilterResource(object sender, PersistentObjectCancelEventArgs e)
        {
            try
            {
                SchedulerStorage storage = (SchedulerStorage)sender;
                Resource res = (Resource)e.Object;

                if (Role != "All")
                {
                    if (GroupByRadioGroup.SelectedIndex == 0)
                    {
                        e.Cancel = ResourceDataTable.AsEnumerable().Where(x => x.Field<string>("ResourceName") == res.Id.ToString() && (RoleRegExpression.IsMatch(x.Field<string>("Role")) || x.Field<string>("Role") == "None")).Count() < 1;
                    }
                    else if (GroupByRadioGroup.SelectedIndex == 1)
                    {
                        Console.WriteLine($"Resource: {res.Id.ToString()}");
                        e.Cancel = ResourceDataTable.AsEnumerable().Where(x => x.Field<string>("ResourceName") == res.Id.ToString() && (x.Field<string>("Department").Contains(departmentComboBox.Text) || x.Field<string>("Role") == "None")).Count() < 1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        #endregion

        #region Department Task View

        private void InitializePrintOptions()
        {
            // string[] departmentArr = { "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "CNC Electrodes", "Grind", "Inspection", "EDM Sinker", "EDM Wire (In-House)", "Polish" };

            foreach (string item in departmentArr)
            {
                PrintDeptsCheckedComboBoxEdit.Properties.Items.Add(item, CheckState.Unchecked, true);
            }

            PrintDeptsCheckedComboBoxEdit.Properties.SeparatorChar = ',';
            PrintDeptsCheckedComboBoxEdit.SetEditValue("Program Rough, Program Finish, Program Electrodes, CNC Rough, CNC Finish, CNC Electrodes, Grind, Inspection, EDM Sinker, EDM Wire (In-House)");
        }

        private CriteriaOperator FilterTaskView(string department, bool includeQuotes, bool includeCompleteTasks)
        {
            List<CriteriaOperator> criteriaOperators = new List<CriteriaOperator>();

            if (includeQuotes == false)
            {
                criteriaOperators.Add(new NotOperator(new FunctionOperator(FunctionOperatorType.Contains, new OperandProperty("JobNumber"), "Quote"))); // Excludes tasks with quote in jobnumber. 
            }

            if (includeCompleteTasks == false)
            {
                criteriaOperators.Add(new NullOperator("Status"));  // Excludes tasks with Status set to null. 
            }

            if (department == "Program")
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
            else if (department == "All")
            {
                //gridView1.ActiveFilterString = String.Empty;
                //gridView1.ClearColumnsFilter();
            }

            footerDateTime = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

            return GroupOperator.And(criteriaOperators);
        }

        // This filter is used by the print resources function.
        private void FilterTaskView2(string resource)
        {
            List<CriteriaOperator> criteriaOperators = new List<CriteriaOperator>();

            criteriaOperators.Add(new NotOperator(new FunctionOperator(FunctionOperatorType.Contains, new OperandProperty("JobNumber"), "Quote")));
            criteriaOperators.Add(new NullOperator("Status"));

            criteriaOperators.Add(new BinaryOperator("Resource", resource, BinaryOperatorType.Equal));


            gridView1.ActiveFilterCriteria = GroupOperator.And(criteriaOperators);

            footerDateTime = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
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

        private void OpenKanBanWorkbook(int rowIndex)
        {
            string jobNumber, component;
            int projectNumber;
            //string column = gridView1.FocusedColumn.FieldName;
            //int rowIndex = gridView1.FocusedRowHandle;
            Database db = new Database();
            ExcelInteractions ei = new ExcelInteractions();

            if (rowIndex >= 0)
            {
                component = gridView1.GetRowCellValue(rowIndex, gridView1.Columns["Component"]).ToString();
                jobNumber = gridView1.GetRowCellValue(rowIndex, gridView1.Columns["JobNumber"]).ToString();
                projectNumber = Convert.ToInt32(gridView1.GetRowCellValue(rowIndex, gridView1.Columns["ProjectNumber"]));
                //MessageBox.Show("Component");
                ei.OpenKanBanWorkbook(db.GetKanBanWorkbookPath(jobNumber, projectNumber), component);
            }
        }

        private DateTime GetDueDate(GridView view, int listSourceRowIndex)
        {
            return ProjectInfoList.Find(x => x.ProjectNumber == (int)view.GetListSourceRowCellValue(listSourceRowIndex, "ProjectNumber")).DueDate;
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
            gridView1.ActiveFilterCriteria = FilterTaskView(departmentComboBox2.Text, false, false);
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

        private void printTaskViewButton_Click(object sender, EventArgs e)
        {
            // Check whether the GridControl can be previewed.
            if (!gridControl1.IsPrintingAvailable)
            {
                MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                return;
            }

            gridView1.Columns["Resource"].OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.True;
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

            gridView1.Columns["Resource"].OptionsColumn.Printable = DevExpress.Utils.DefaultBoolean.False;
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
                        gridView1.OptionsPrint.RtfPageHeader = @"{\rtf1\deff0{\fonttbl{\f0 Calibri;}{\f1 Microsoft Sans Serif;}}{\colortbl ;\red0\green0\blue255 ;}{\*\defchp \b\f1\fs22}{\stylesheet {\ql\b\f1\fs22 Normal;}{\*\cs1\b\f1\fs22 Default Paragraph Font;}{\*\cs2\sbasedon1\b\f1\fs22 Line Number;}{\*\cs3\b\ul\f1\fs22\cf1 Hyperlink;}{\*\ts4\tsrowd\b\f1\fs22\ql\tscellpaddfl3\tscellpaddl108\tscellpaddfb3\tscellpaddfr3\tscellpaddr108\tscellpaddft3\tsvertalt\cltxlrtb Normal Table;}{\*\ts5\tsrowd\sbasedon4\b\f1\fs22\ql\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\tscellpaddfl3\tscellpaddl108\tscellpaddfr3\tscellpaddr108\tsvertalt\cltxlrtb Table Simple 1;}}{\*\listoverridetable}{\info{\creatim\yr2018\mo1\dy10\hr10\min20}{\version1}}\nouicompat\splytwnine\htmautsp\sectd\trowd\irow0\irowband-1\lastrow\ts5\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\trleft-108\trautofit1\trpaddfl3\trpaddl108\trpaddfr3\trpaddr108\tbllkhdrcols\tbllkhdrrows\tbllknocolband\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx3810\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx7710\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx11610\pard\plain\ql\intbl\yts5{\b\f1\fs22\cf0 Job Number: All}\b\f1\fs22\cell\pard\plain\qc\intbl\yts5{\b\f1\fs22\cf0 Resource: " + PrintEmployeeWorkCheckedComboBoxEdit.Properties.Items[i].Value + @"}\b\f1\fs22\cell\pard\plain\qr\intbl\yts5{\b\f1\fs22\cf0 Component: All}\b\f1\fs22\cell\trowd\irow0\irowband-1\lastrow\ts5\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\trleft-108\trautofit1\trpaddfl3\trpaddl108\trpaddfr3\trpaddr108\tbllkhdrcols\tbllkhdrrows\tbllknocolband\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx3810\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx7710\clvertalt\clbrdrt\brdrnone\brdrw10\clbrdrl\brdrnone\brdrw10\clbrdrb\brdrnone\brdrw10\clbrdrr\brdrnone\brdrw10\cltxlrtb\clftsWidth3\clwWidth3888\clpadfr3\clpadr108\clpadft3\clpadt108\cellx11610\row\pard\plain\ql\b\f1\fs22\par}";
                        //gridView1.OptionsPrint.RtfPageHeader = richEditControl1.RtfText;
                        gridView1.OptionsPrint.RtfPageFooter = @"{\rtf1\ansi {\fonttbl\f0\ Microsoft Sans Serif;} \f0\pard \fs18 \qr \b Report Date: " + footerDateTime + @"\b0 \par}";
                        gridView1.OptionsPrint.AutoWidth = false;
                        //gridView1.GridControl.ShowPrintPreview();
                        gridView1.GridControl.Print();
                    }
                }
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

        private void gridView1_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            try
            {
                //MessageBox.Show("gridView1_CellValueChanged");
                GridView view = sender as GridView;
                Database db = new Database();

                int projectNumber, taskID, id;
                string jobNumber, component, predecessors, duration, machine, resource;

                id = (int)view.GetFocusedRowCellValue("ID");
                jobNumber = (string)view.GetFocusedRowCellValue("JobNumber");
                component = (string)view.GetFocusedRowCellValue("Component");
                duration = (string)view.GetFocusedRowCellValue("Duration");
                predecessors = (string)view.GetFocusedRowCellValue("Predecessors");
                projectNumber = (int)view.GetFocusedRowCellValue("ProjectNumber");
                taskID = (int)view.GetFocusedRowCellValue("TaskID");
                machine = (string)view.GetFocusedRowCellValue("Machine");
                resource = (string)view.GetFocusedRowCellValue("Resource");

                deptTaskViewHelper = new RefreshHelper(gridView1, "ProjectNumber");

                //ProjectInfo pi = ProjectInfoList.Find(x => x.ProjectNumber == (int)view.GetListSourceRowCellValue(e.RowHandle, "ProjectNumber"));

                if (e.Column.FieldName == "StartDate")
                {
                    db.ChangeTaskStartDate(jobNumber, projectNumber, component, (DateTime)e.Value, duration, taskID);
                }
                else if (e.Column.FieldName == "FinishDate")
                {
                    db.ChangeTaskFinishDate(jobNumber, projectNumber, component, (DateTime)e.Value, taskID);
                }
                else if (e.Column.FieldName == "Machine" || e.Column.FieldName == "Resource")
                {
                    SetAppointmentResources(id, machine, resource);
                    db.UpdateTasksTable(sender, e);
                }
                else
                {
                    db.UpdateTasksTable(sender, e);
                }

                deptTaskViewHelper.SaveViewInfo();
                this.tasksTableAdapter.Fill(this.workload_Tracking_System_DBDataSet.Tasks);
                deptTaskViewHelper.LoadViewInfo();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void gridView1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("gridView1_Click");
        }

        private void RefreshTasksButton_Click(object sender, EventArgs e)
        {
            try
            {
                ProjectInfoList = Database.GetProjectInfoList();
                deptTaskViewHelper = new RefreshHelper(gridView1, "ProjectNumber");
                RoleTable = Database.GetRoleTable();
                deptTaskViewHelper.SaveViewInfo();
                this.tasksTableAdapter.Fill(this.workload_Tracking_System_DBDataSet.Tasks);
                PopulateEmployeeComboBox();
                deptTaskViewHelper.LoadViewInfo();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void gridView1_CustomUnboundColumnData(object sender, CustomColumnDataEventArgs e)
        {
            GridView view = sender as GridView;

            ProjectModel pi = ProjectInfoList.Find(x => x.ProjectNumber == (int)view.GetListSourceRowCellValue(e.ListSourceRowIndex, "ProjectNumber"));

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
            string machineType = "";
            Database db = new Database();

            if (task == "CNC Rough")
            {
                machineType = "Rough Mill";
            }
            else if (task == "CNC Finish")
            {
                machineType = "Finish Mill";
            }
            else if(task == "CNC Electrodes")
            {
                machineType = "Graphite Mill";
            }

            if (machineType != "")
            {
                repositoryItemCheckedComboBoxEdit1.DataSource = GetResourceList(machineType);
            }

            //MessageBox.Show("Test1");
        }

        private void gridView1_ShownEditor(object sender, EventArgs e)
        {
            ComboBoxEdit comboBoxEdit = null;

            if (gridView1.ActiveEditor.EditorTypeName == "ComboBoxEdit")
            {
                comboBoxEdit = gridView1.ActiveEditor as ComboBoxEdit;
            }

            if (comboBoxEdit != null)
            {
                string department = departmentComboBox2.Text;
                string role = "";

                comboBoxEdit.Properties.Items.Clear();

                if (department == "Program Rough")
                {
                    role = "Rough Programmer";
                }
                else if (department == "Program Finish")
                {
                    role = "Finish Programmer";
                }
                else if (department == "Program Electrodes")
                {
                    role = "Electrode Programmer";
                }
                else if (department.EndsWith("Grind") || department == "Polish")
                {
                    role = "Tool Maker";
                }
                else if (department == "CNC Rough")
                {
                    role = "Rough CNC Operator";
                }
                else if (department == "CNC Finish")
                {
                    role = "Finish CNC Operator";
                }
                else if( department == "CNC Electrodes")
                {
                    role = "Electrode CNC Operator";
                }
                else if (department == "EDM Wire (In-House)")
                {
                    role = "EDM Wire Operator";
                }
                else if (department == "EDM Sinker")
                {
                    role = "EDM Sinker Operator";
                }
                else if (department == "Hole Pop")
                {
                    role = "Hole Popper Operator";
                }
                else if (department.StartsWith("Inspection"))
                {
                    role = "CMM Operator";
                }
                else if(department == "All")
                {
                    role = "";
                }

                comboBoxEdit.Properties.Items.Add("");
                comboBoxEdit.Properties.Items.AddRange(GetResourceList(role).ToArray());

                //if (role != "")
                //{
                //    comboBoxEdit.Properties.Items.Add("");
                //    comboBoxEdit.Properties.Items.AddRange(GetResourceList(role).ToArray());
                //}
            }
        }

        #endregion

        #region Project View

        private void InitializeComponentGrid()
        {
            Database db = new Database();

            try
            {
                //gridControl3.DataSource = db.GetComponentCompletionPercents();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
            }
        }

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

            this.projectsTableAdapter.Fill(this.workload_Tracking_System_DBDataSet.Projects);
            this.componentsTableAdapter.Fill(this.workload_Tracking_System_DBDataSet.Components);
            this.tasksTableAdapter.Fill(this.workload_Tracking_System_DBDataSet.Tasks);

            RecursiveExpand();
            SelectRows();
        }

        private int GetRowHandleByColumnValue(GridView view, string ColumnFieldName, object value)
        {
            int result = GridControl.InvalidRowHandle;
            for (int i = 0; i < view.RowCount; i++)
                if (view.GetDataRow(i)[ColumnFieldName].Equals(value))
                    return i;
            return result;
        }

        private bool KanBanExists(string jobNumber, int projectNumber)
        {
            Database db = new Database();

            string kanBanWorkbookPath = db.GetKanBanWorkbookPath(jobNumber, projectNumber);

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
            using (var form = new ProjectCreationForm(schedulerStorage1))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    if (form.DataValidated)
                    {
                        RefreshProjectGrid();
                        gridView3.FocusedRowHandle = GetRowHandleByColumnValue(gridView3, "ProjectNumber", form.Project.ProjectNumber);
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
            using (var form = new ProjectCreationForm(project, schedulerStorage1))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    if (form.DataValidated)
                    {
                        if (gridView3.GetFocusedRowCellValue("KanBanWorkbookPath").ToString().Length > 0)
                        {
                            MessageBox.Show("A project has changed.  Need to regenerate and reprint Kan Ban.");
                            gridView3.Appearance.FocusedRow.BackColor = Color.Red;
                        }

                        RefreshProjectGrid();
                    }
                }
                else if (result == DialogResult.Cancel)
                {

                }
            }
        }

        // TODO: Wait and see how selecting rows from grid control work for this.
        //private List<string> GetComponentListFromUser(string textString = "")
        //{
        //    Database db = new Database();
        //    var number = GetComboBoxInfo();
        //    List<string> componentList = db.GetComponentList(number.jobNumber, number.projectNumber);

        //    using (var form = new SelectComponentsWindow(componentList, textString))
        //    {
        //        var result = form.ShowDialog();

        //        if (result == DialogResult.OK)
        //        {
        //            return form.ComponentList;
        //        }
        //        else if (result == DialogResult.Cancel)
        //        {

        //        }

        //        return null;
        //    }
        //}

        private void gridView3_KeyDown(object sender, KeyEventArgs e)
        { 
            if (e.KeyCode == Keys.Delete && e.Modifiers == Keys.Control)
            {
                if (MessageBox.Show("Delete Project?", "Confirmation", MessageBoxButtons.YesNo) != DialogResult.Yes)
                    return;

                GridView view = sender as GridView;
                Database db = new Database();

                try
                {
                    if (db.RemoveProject((string)view.GetFocusedRowCellValue("JobNumber"), (int)view.GetFocusedRowCellValue("ProjectNumber")))
                    {
                        view.DeleteRow(view.FocusedRowHandle);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                }
            }
        }
        private void GridView3_ValidatingEditor(object sender, BaseContainerValidateEditorEventArgs e)
        {
            ColumnView view = sender as ColumnView;
            GridColumn column = (e as EditFormValidateEditorEventArgs)?.Column ?? view.FocusedColumn;

            if (column.FieldName == "ProjectNumber")
            {
                if (int.TryParse(e.Value.ToString(), out int result) == true)
                {
                    if (Database.ProjectExists(result))
                    {
                        e.ErrorText = "There is already a project with that number.";
                        e.Valid = false;
                    }
                }
                else
                {
                    e.ErrorText = "Project number must be a number.";
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
                Database db = new Database();
                GridView view = sender as GridView;
                //string resource, jobNumber, department;
                //int projectNumber;

                if (!db.UpdateProjectsTable(sender, e))
                {
                    RefreshProjectGrid();
                }
                else
                {
                    if (e.Column.FieldName.Contains("Programmer"))
                    {
                        db.SetTaskResources(sender, e, schedulerStorage1);
                        //SetAppointmentResources(sender, e);
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

        private void gridView4_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            Database db = new Database();

            try
            {
                db.UpdateComponentsTable(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
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
        private void repositoryItemImageEdit2_Popup(object sender, EventArgs e)
        {
            //MessageBox.Show("Popup");
        }

        private void repositoryItemImageEdit2_ImageChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Image changed.");
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
                Console.WriteLine(ex.Message + "\n\n" + ex.StackTrace);
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }
        private void gridView4_CustomRowCellEditForEditing(object sender, CustomRowCellEditEventArgs e)
        {
            if (e.Column.FieldName == "Pictures")
            {
                e.RepositoryItem = repositoryItemImageEdit2;
            }
        }
        private void GridView5_ValidatingEditor(object sender, BaseContainerValidateEditorEventArgs e)
        {
            ColumnView view = sender as ColumnView;
            GridColumn column = (e as EditFormValidateEditorEventArgs)?.Column ?? view.FocusedColumn;
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
        // Not sure how this event works.
        private void gridView5_CellValueChanging(object sender, CellValueChangedEventArgs e)
        {


        }
        private void gridView5_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            try
            {
                GridView view = sender as GridView;
                Database db = new Database();

                int projectNumber, taskID, id;
                string jobNumber, component, predecessors, duration;

                id = (int)view.GetFocusedRowCellValue("ID");
                jobNumber = (string)view.GetFocusedRowCellValue("JobNumber");
                component = (string)view.GetFocusedRowCellValue("Component");
                duration = (string)view.GetFocusedRowCellValue("Duration");
                predecessors = (string)view.GetFocusedRowCellValue("Predecessors");
                projectNumber = (int)view.GetFocusedRowCellValue("ProjectNumber");
                taskID = (int)view.GetFocusedRowCellValue("TaskID");

                //ProjectInfo pi = ProjectInfoList.Find(x => x.ProjectNumber == (int)view.GetListSourceRowCellValue(e.RowHandle, "ProjectNumber"));

                if (e.Column.FieldName == "StartDate" && e.Value != DBNull.Value)
                {
                    db.ChangeTaskStartDate(jobNumber, projectNumber, component, (DateTime)e.Value, duration, taskID);

                    RefreshProjectGrid();
                }
                else if (e.Column.FieldName == "FinishDate" && e.Value != DBNull.Value)
                {
                    db.ChangeTaskFinishDate(jobNumber, projectNumber, component, (DateTime)e.Value, taskID);

                    RefreshProjectGrid();
                }
                else
                {
                    if (e.Column.FieldName == "Machine" || e.Column.FieldName == "Resource")
                    {
                        SetAppointmentResources(id, view.GetRowCellValue(e.RowHandle, "Machine").ToString(), view.GetRowCellValue(e.RowHandle, "Resource").ToString());
                    }

                    db.UpdateTasksTable(sender, e);  // Resources field is only updated when the Machine or Resource fields change., resources
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void RefreshProjectsButton_Click(object sender, EventArgs e)
        {
            try
            {
                RefreshProjectGrid();
                //helper1.LoadViewInfo();
                //helper2.LoadViewInfo();

                //ExpandStoredRows();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void copyButton_Click(object sender, EventArgs e)
        {
            Database db = new Database();

            try
            {
                ProjectModel pi = db.GetProject((int)gridView3.GetFocusedRowCellValue("ProjectNumber")); // (string)gridView3.GetFocusedRowCellValue("JobNumber"), 

                if (gridView3.SelectedRowsCount == 1)
                {
                    var result = XtraInputBox.Show("Change Project #", "Copy Project", "");

                    if (result != null)
                    {
                        if (int.TryParse(result.ToString(), out int projectNumber))
                        {
                            pi.SetDefaultCopiedProjectInfo(projectNumber);
                            db.LoadProjectToDB(pi);
                            RefreshProjectGrid();
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
            try
            {
                if (gridView3.SelectedRowsCount != 1)
                {
                    MessageBox.Show("Please select a project.");
                    return;
                }
                else
                {
                    Database db = new Database();
                    ExcelInteractions ei = new ExcelInteractions();
                    string jobNumber = gridView3.GetFocusedRowCellValue("JobNumber").ToString();
                    int projectNumber = (int)gridView3.GetFocusedRowCellValue("ProjectNumber");
                    string path;
                    List<string> componentList = new List<string>();

                    ProjectModel pi = db.GetProject(projectNumber);

                    //if (BlankStartFinishDateExists(pi))
                    //{
                    //    MessageBox.Show("A blank start or finish date exists. Please fill in all dates.");
                    //    return;
                    //}

                    if (KanBanExists(jobNumber, projectNumber))
                    {
                        DialogResult result = XtraMessageBox.Show("A Kan Ban for this project already exists.\n\nDo you want to create a new one?\n\n" +
                            "(Click 'Yes' to create new one (All info preserved).  Click 'No' to cancel.)", "Warning",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                        if (result == DialogResult.Yes)
                        {
                            path = ei.GenerateKanBanWorkbook2(pi);

                            gridView3.Appearance.FocusedRow.BackColor = Color.White;

                            if (path != "")
                            {
                                db.SetKanBanWorkbookPath(path, pi.JobNumber, pi.ProjectNumber);
                            }
                        }
                        else if (result == DialogResult.No)
                        {
                            //componentList = GetListOfSelectedComponents();

                            //if (componentList.Count == 0)
                            //{
                            //    XtraMessageBox.Show("No components selected.");
                            //    return;
                            //}

                            //ei.EditKanBanWorkbook(pi, db.GetKanBanWorkbookPath(jobNumber, projectNumber), componentList);

                            return;
                        }
                        else if (result == DialogResult.Cancel)
                        {
                            return;
                        }

                    }
                    else
                    {
                        path = ei.GenerateKanBanWorkbook2(pi);

                        if (path != "")
                        {
                            db.SetKanBanWorkbookPath(path, pi.JobNumber, pi.ProjectNumber);
                        }
                    }

                    RefreshProjectGrid();
                }
            }
            catch (Exception ex1)
            {
                MessageBox.Show(ex1.Message + "\n\n" + ex1.StackTrace);
            }
        }

        private void forwardDateButton_Click(object sender, EventArgs e)
        {
            List<string> componentList = new List<string>();
            Database db = new Database();

            try
            {
                string jobNumber = gridView3.GetFocusedRowCellValue("JobNumber").ToString();
                int projectNumber = (int)gridView3.GetFocusedRowCellValue("ProjectNumber");
                //DateTime preselectDate = (DateTime)
                componentList = GetListOfSelectedComponents();

                if (componentList.Count == 0)
                {
                    XtraMessageBox.Show("No components selected.");
                    return;
                }

                using (var form = new ForwardDateWindow("Forward Date", DateTime.Today))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        db.ForwardDateProjectTasks(jobNumber, projectNumber, componentList, form.ForwardDate);
                        RefreshProjectGrid();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void backDateButton_Click(object sender, EventArgs e)
        {
            List<string> componentList = new List<string>();
            Database db = new Database();

            try
            {
                string jobNumber = gridView3.GetFocusedRowCellValue("JobNumber").ToString();
                int projectNumber = (int)gridView3.GetFocusedRowCellValue("ProjectNumber");
                DateTime preselectDate = (DateTime)gridView3.GetFocusedRowCellValue("DueDate");
                componentList = GetListOfSelectedComponents();

                if (componentList.Count == 0)
                {
                    XtraMessageBox.Show("No components selected.");
                    return;
                }

                using (var form = new ForwardDateWindow("Back Date", preselectDate))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        db.BackDateProjectTasks(jobNumber, projectNumber, componentList, form.ForwardDate);
                        RefreshProjectGrid();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void createProjectButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine("click");

            try
            {
                CreateProject();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void editProjectButton_Click(object sender, EventArgs e)
        {
            try
            {
                Database db = new Database();
                int projectNumber = (int)gridView3.GetFocusedRowCellValue("ProjectNumber");                    
                ProjectModel project = db.GetProject(projectNumber);
                EditProject(project);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private List<string> GetListOfSelectedComponents()
        {
            List<string> componentList = new List<string>();

            if (gridView3.GetMasterRowExpanded(gridView3.GetSelectedRows()[0]))
            {
                var childView = gridView3.GetVisibleDetailView(gridView3.GetSelectedRows()[0]) as GridView;

                foreach (int rowHandle in childView.GetSelectedRows())
                {
                    componentList.Add((string)childView.GetRowCellValue(rowHandle, "Component"));
                }
            }

            return componentList;
        }

        private void resourceButton_Click(object sender, EventArgs e)
        {
            ManageResourcesForm form = new ManageResourcesForm();
            form.Show();
        }

        #endregion

        #region Chart View

        private void LoadGraph(List<Week> weekList, List<string> departmentList)
        {
            Database db = new Database();
            Series tempSeries;

            int i = 0;

            chartControl1.Series.Clear();

            List<string> weekTitleArr = new List<string>();
            DataTable dailyDeptCapacities = db.GetDailyDepartmentCapacities();
            int dailyCapacity;

            for (int n = 0; n < 20; n++)
            {
                weekTitleArr.Add(n.ToString());
            }

            //SideBySideBarSeries series = new SideBySideBarSeries();
            //Series series1 = new Series("Program Rough Hours", ViewType.Bar);
            //Series series2 = new Series("Program Finish Hours", ViewType.Bar);
            //Series series3 = new Series("Program Electrode Hours", ViewType.Bar);

            if (TimeUnits == "Days")
            {
                foreach(Week week in weekList)
                {
                    dailyCapacity = dailyDeptCapacities.AsEnumerable().Where(p => p.Field<string>("Department").ToString().Contains(week.Department)).Select(p => p.Field<int>("DailyCapacity")).FirstOrDefault();
                    tempSeries = new Series(week.Department, ViewType.Bar); //  + " Hours (Cap. " + dailyCapacity + ")"

                    foreach (Day day in week.DayList)
                    {
                        tempSeries.Points.Add(new SeriesPoint(day.DayName, (int)day.Hours));
                    }

                    chartControl1.Series.Add(tempSeries);
                }
            }
            else if(TimeUnits == "Weeks")
            {
                foreach (string dept in departmentList)
                {
                    dailyCapacity = dailyDeptCapacities.AsEnumerable().Where(p => p.Field<string>("Department").ToString().Contains(dept)).Select(p => p.Field<int>("DailyCapacity")).FirstOrDefault();
                    tempSeries = new Series(dept, ViewType.Bar); //  + " Hours (Cap." + dailyCapacity * 5 + ")"

                    var deptWeeks = from wks in weekList
                                    where wks.Department == dept
                                    orderby wks.WeekStart
                                    select wks;

                    foreach (Week week in deptWeeks)
                    {
                        // weekTitleArr[i++]
                        tempSeries.Points.Add(new SeriesPoint("WK " + weekTitleArr[i++] + " " + week.WeekStart.ToShortDateString() , (int)week.GetWeekHours()));
                    }

                    chartControl1.Series.Add(tempSeries);

                    i = 0;
                }
            }
        }

        private void LoadGraph(Week week)
        {
            chartControl2.Series.Clear();

            //SideBySideBarSeries series = new SideBySideBarSeries();
            Series series1 = new Series(week.Department + " Hours", ViewType.Bar);
            //Series series2 = new Series("Program Finish Hours", ViewType.Bar);
            //Series series3 = new Series("Program Electrode Hours", ViewType.Bar);

            foreach (Day day in week.DayList)
            {
                series1.Points.Add(new SeriesPoint(day.DayName, (int)day.Hours));
            }

            chartControl2.Series.Add(series1);
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
            Database db = new Database();
            List<Week> weeksList = new List<Week>();
            List<string> departmentList = new List<string>();
            string resourceType = GetResourceType();

            if (resourceType == "Department")
            {
                departmentList = Database.GetDepartments();
            }
            else if (resourceType == "Personnel")
            {
                departmentList = Database.GetAllResourcesOfType("Person");
            }

            if (timeFrameComboBoxEdit.Text != "")
            {
                string weekStart, weekEnd;

                weekStart = timeFrameComboBoxEdit.Text.Split(' ')[0];
                weekEnd = timeFrameComboBoxEdit.Text.Split(' ')[2];

                if (TimeUnits == "Days")
                {
                    weeksList = db.GetDayHours(weekStart, weekEnd);
                }
                else if (TimeUnits == "Weeks")
                {
                    weeksList = db.GetWeekHours(weekStart, weekEnd, departmentList, resourceType);
                }

                LoadGraph(weeksList, departmentList);
            }
        }
        private string GetResourceType()
        {
            return chartRadioGroup.Properties.Items[chartRadioGroup.SelectedIndex].Description.ToString();
        }
        private void GetDepartmentHours()
        {
            Database db = new Database();
            Week week;

            if (timeFrameComboBoxEdit.Text != "" && departmentComboBox3.Text != "")
            {
                string weekStart, weekEnd;

                weekStart = timeFrameComboBoxEdit.Text.Split(' ')[0];
                weekEnd = timeFrameComboBoxEdit.Text.Split(' ')[2];

                week = db.GetDayHours(weekStart, weekEnd).Find(x => x.Department == departmentComboBox3.Text);

                LoadGraph(week);
            }
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

        private void departmentComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetDepartmentHours();
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
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
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

        #endregion

        #region Gantt View

        private void LoadProject()
        {
            Database db = new Database();
            var number = GetComboBoxInfo();
            ProjectModel project = db.GetProject(number.projectNumber); // number.jobNumber, 

            schedulerStorage2.Appointments.ResourceSharing = true;

            InitializeResources(project);

            GenerateEventList(CustomEventList, project);

            schedulerStorage2.Appointments.CustomFieldMappings.Clear();
            schedulerStorage2.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("TaskId", "TaskId"));

            AppointmentMappingInfo appointmentMappings = schedulerStorage2.Appointments.Mappings;

            appointmentMappings.AppointmentId = "AppointmentID";

            appointmentMappings.Subject = "Subject";
            appointmentMappings.Location = "Location";
            //appointmentMappings.Description = "Notes";
            appointmentMappings.ResourceId = "OwnerId";
            appointmentMappings.Start = "StartDate";
            appointmentMappings.End = "FinishDate";

            schedulerStorage2.Appointments.DataSource = CustomEventList;

            Console.WriteLine("Check Appointments");

            for (int i = 0; i < schedulerStorage2.Appointments.Count; i++)
            {
                Appointment appointment = schedulerStorage2.Appointments[i];
                Console.WriteLine($"{appointment.Subject} {appointment.Location} {appointment.ResourceId} {appointment.Start} {appointment.End}");
            }

            InitializeDependencies(project);

            AppointmentDependencyMappingInfo appointmentDependencyMappingInfo = schedulerStorage2.AppointmentDependencies.Mappings;

            appointmentDependencyMappingInfo.DependentId = "DepID";
            appointmentDependencyMappingInfo.ParentId = "ParentID";

            schedulerStorage2.AppointmentDependencies.DataSource = CustomDependencyList;

            Console.WriteLine("Check Appointment Dependencies");

            for (int i = 0; i < schedulerStorage2.AppointmentDependencies.Count; i++)
            {
                AppointmentDependency appointmentDependency = schedulerStorage2.AppointmentDependencies[i];
                Console.WriteLine($"{appointmentDependency.DependentId} {appointmentDependency.ParentId}");
            }
        }

        private void InitializeResources(ProjectModel project)
        {
            int i = 0;
            int ParentID = 0;

            //CustomResourceCollection.Clear();
            CustomResourceCollection = new BindingList<CustomResource>();

            foreach (ComponentModel component in project.ComponentList)
            {
                ParentID = i;
                CustomResourceCollection.Add(CreateCustomResource(i++, -1, component.Name));

                foreach (TaskModel task in component.TaskList)
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

            foreach (ComponentModel component in project.ComponentList)
            {
                baseCount = i++;

                foreach (TaskModel task in component.TaskList)
                {
                    Resource resource = schedulerStorage2.Resources[i++];
                    eventList.Add(CreateEvent(task.ID + baseCount, project.JobNumber + " #" + project.ProjectNumber + " " + component.Name, resource.Id, task.ID, task.TaskName, task.StartDate, task.FinishDate));
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

            foreach (ComponentModel component in project.ComponentList)
            {
                baseCount = aID - 1;

                foreach (TaskModel task in component.TaskList)
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

        private void LoadProject(string jobNumber, int projectNumber)
        {
            Database db = new Database();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();

            Project = db.GetProject(projectNumber); // jobNumber, 

            //dt = db.getProjectData(jobNumber, projectNumber);

            //foreach (ClassLibrary.Component compi in Project.ComponentList)
            //{
            //    foreach (TaskInfo taski in compi.TaskList)
            //    {
            //        Console.WriteLine(compi.Name + " " + taski.ID + " " + taski.TaskName);
            //    }
            //}

            ResourceMappingInfo resourceMappings = this.schedulerStorage2.Resources.Mappings;

            resourceMappings.Id = "AptID";
            resourceMappings.ParentId = "ParentID"; // Need this for hierarchy in resource tree.
            resourceMappings.Caption = "TaskName"; // In the Resource tree designer the field name has to match the field that is mapped to caption.

            dt2 = GetProjectResourceData(Project);

            schedulerStorage2.Resources.Clear();

            Stopwatch sw = new Stopwatch();
            sw.Start();

            schedulerStorage2.Resources.DataSource = dt2; // Woohoo!! This finally works!

            //int i = 1;

            //THIS WORKS FOR SOME REASON.

            //foreach (DataRow nrow in dt2.Rows)
            //{
            //    Resource resource = schedulerStorage2.CreateResource(i, nrow["TaskName"].ToString());
            //    resource.ParentId = nrow["ParentID"];

            //    schedulerStorage2.Resources.Add(resource);
            //    i++;
            //}

            //THIS DOES NOT.

            //foreach (DataRow nrow in dt2.Rows)
            //{
            //    if (nrow["ParentID"] == DBNull.Value)
            //    {
            //        CustomResourceCollection.Add(CreateCustomResource(i++, -1, nrow["TaskName"].ToString()));
            //    }
            //    else
            //    {
            //        //Console.WriteLine($"{i}, {nrow["ParentID"]}, {nrow["TaskName"].ToString()}");
            //        CustomResourceCollection.Add(CreateCustomResource(i++, Convert.ToInt32(nrow["ParentID"]), nrow["TaskName"].ToString()));
            //    }
            //}

            //this.schedulerStorage2.Resources.DataSource = CustomResourceCollection;

            Console.WriteLine(sw.Elapsed);

            if (schedulerStorage2.Appointments.Count > 0)
            {
                schedulerStorage2.Appointments.Clear();
            }
            
            //schedulerStorage2.Appointments.CustomFieldMappings.Clear();  // Added appointment mappings when mainwindow form loads.


            AppointmentMappingInfo appointmentMappings = schedulerStorage2.Appointments.Mappings;

            appointmentMappings.AppointmentId = "AptID";
            appointmentMappings.Start = "StartDate";
            appointmentMappings.End = "FinishDate";
            appointmentMappings.Subject = "Subject";
            appointmentMappings.Location = "Location";
            appointmentMappings.Description = "Notes";
            appointmentMappings.PercentComplete = "PercentComplete";
            appointmentMappings.ResourceId = "AptID";
            Console.WriteLine(sw.Elapsed);
            dt1 = db.LoadProjectToDataTable(Project);
            Console.WriteLine(sw.Elapsed);
            var result = from taskTable in dt1.AsEnumerable()
                         where taskTable.IsNull("StartDate") || taskTable.IsNull("FinishDate")
                         select taskTable;

            int count = result.ToList().Count;

            if (count > 0)
            {
                MessageBox.Show("Project contains " + count + " task(s) with missing date(s).");
            }

            schedulerStorage2.Appointments.DataSource = dt1;

            AppointmentDependencyMappingInfo appointmentDependencyMappingInfo = schedulerStorage2.AppointmentDependencies.Mappings;

            appointmentDependencyMappingInfo.DependentId = "DependentId";
            appointmentDependencyMappingInfo.ParentId = "ParentId";

            schedulerStorage2.AppointmentDependencies.DataSource = db.GetDependencyData(dt1);
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

            foreach (ComponentModel component in project.ComponentList)
            {
                DataRow newRow1 = dt.NewRow();

                newRow1["AptID"] = i;
                newRow1["TaskName"] = component.Name;
                parentID = i++;

                dt.Rows.Add(newRow1);

                foreach (TaskModel task in component.TaskList)
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

        private void projectComboBox_BeforePopup(object sender, EventArgs e)
        {

        }

        private void projectComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                var number = GetComboBoxInfo();
                SplashScreenManager.ShowForm(typeof(WaitForm1));
                //LoadProject();
                LoadProject(number.jobNumber, number.projectNumber);

                SplashScreenManager.CloseForm();
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

        private void schedulerStorage2_AppointmentsChanged(object sender, PersistentObjectsEventArgs e)
        {
            List<int> collapsedNodes = new List<int>();

            foreach (Appointment apt in e.Objects)
            {
                if (UpdateTaskStorage2(apt))
                {
                    var number = GetComboBoxInfo();
                    collapsedNodes = GetCollapsedNodes();
                    LoadProject(number.jobNumber, number.projectNumber);
                    schedulerControl2.RefreshData();
                    CollapseNodes(collapsedNodes);
                }
                else
                {
                    var number = GetComboBoxInfo();
                    collapsedNodes = GetCollapsedNodes();
                    LoadProject(number.jobNumber, number.projectNumber);
                    schedulerControl2.RefreshData();
                    CollapseNodes(collapsedNodes);
                }
                //MessageBox.Show(apt.Subject);
            }
        }

        private void RefreshGanttButton_Click(object sender, EventArgs e)
        {
            PopulateProjectComboBox();

            try
            {
                if (projectComboBox.Text != "")
                {
                    var number = GetComboBoxInfo();

                    SplashScreenManager.ShowForm(typeof(WaitForm1));
                    //LoadProject();
                    LoadProject(number.jobNumber, number.projectNumber);

                    SplashScreenManager.CloseForm();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        #endregion

        #region WorkLoad

        private string[] validEditorArr = { "michell.willey", "mikeh", "Mikeh", "marks", "Marks", "joshua.meservey" };
        private void CollapseCompletedGroup()
        {
            int count = 0;
            for (int i = 0; i < bandedGridView1.RowCount; i++)
            {
                int rowHandle = bandedGridView1.GetVisibleRowHandle(i);
                if (bandedGridView1.IsGroupRow(rowHandle))
                {
                    count++;
                    //MessageBox.Show(count.ToString() + " " + rowHandle + " " + bandedGridView1.GetGroupRowDisplayText(rowHandle));
                    if (bandedGridView1.GetGroupRowDisplayText(rowHandle).Contains("Completed"))
                    {
                        bandedGridView1.CollapseGroupRow(rowHandle);
                    }
                }
            }
        }

        private void gridControl2_Load(object sender, EventArgs e)
        {
            footerDateTime = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
            ColorList = Database.GetColorEntries();
            CollapseCompletedGroup();
        }

        private void bandedGridView1_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            try
            {
                Console.WriteLine("bandedGridView1 Cell Value Changed Event");
                //Console.WriteLine("Changed Cell Value: " + e.Value.ToString());
                if (!validEditorArr.ToList<string>().Exists(x => x == Environment.UserName.ToString()))
                {
                    MessageBox.Show("This login is not authorized to make changes to work load tab.");
                    gridControl2.DataSource = workLoadTableAdapter.GetData();
                    CollapseCompletedGroup();
                    return;
                }

                if (e.Column.FieldName == "MWONumber" || e.Column.FieldName == "ProjectNumber")
                {
                    if (int.TryParse(e.Value.ToString(), out int number))
                    {
                       
                    }
                    else if (e.Value.ToString() == "")
                    {

                    }
                    else
                    {
                        MessageBox.Show("Please enter a number.");

                        if (bandedGridView1.GetFocusedRowCellValue("ID").ToString() == "-1")
                        {
                            return;
                        }
                        else
                        {
                            RefreshWorkloadGrid();
                            return;
                        }
                    }
                }

                Database db = new Database();

                if (bandedGridView1.GetFocusedRowCellValue("ID").ToString() != "-1" && !db.UpdateWorkloadTable(sender, e))
                {
                    RefreshWorkloadGrid();
                    return;
                }

                if (e.Column.FieldName == "DeliveryInWeeks" && bandedGridView1.GetFocusedRowCellValue("StartDate").ToString() != "")
                {
                    bandedGridView1.SetFocusedRowCellValue("FinishDate", Convert.ToDateTime(bandedGridView1.GetFocusedRowCellValue("StartDate")).AddDays(Convert.ToDouble(e.Value) * 7));
                }
                else if (e.Column.FieldName == "StartDate" && bandedGridView1.GetFocusedRowCellValue("DeliveryInWeeks").ToString() != "0")
                {
                    if (int.TryParse(bandedGridView1.GetFocusedRowCellValue("DeliveryInWeeks").ToString(), out int result))
                    {
                        bandedGridView1.SetFocusedRowCellValue("FinishDate", Convert.ToDateTime(e.Value).AddDays(Convert.ToDouble(bandedGridView1.GetFocusedRowCellValue("DeliveryInWeeks")) * 7));
                    }
                }
                else if (e.Column.FieldName == "RoughProgrammer" || e.Column.FieldName == "ElectrodeProgrammer" || e.Column.FieldName == "FinishProgrammer")
                {

                    db.UpdateProjectsTable(sender, e);
                    db.SetTaskResources(sender, e, schedulerStorage1);
                    RefreshProjectGrid();
                    RefreshDepartmentScheduleView();

                    //int rowHandle = -1;
                    //string projectNumber = bandedGridView1.GetRowCellValue(e.RowHandle, "ProjectNumber").ToString();
                    //Int32 mwoNumber = Int32.Parse(bandedGridView1.GetRowCellValue(e.RowHandle, "MWONumber").ToString());
                    //Int64 id = Int64.Parse(bandedGridView1.GetRowCellValue(e.RowHandle, "ID").ToString());

                    //if (projectNumber != "" && mwoNumber.ToString() == "")
                    //{
                    //    rowHandle = gridView3.LocateByValue("ProjectNumber", int.Parse(projectNumber));
                    //}
                    //else if (projectNumber == "" && mwoNumber.ToString() != "")
                    //{
                    //    //MessageBox.Show(mwoNumber);
                    //    // Entering a hard number for value causes this to work.  How do you pass a value with a variable?
                    //    rowHandle = gridView3.LocateByValue("ProjectNumber", bandedGridView1.GetRowCellValue(e.RowHandle, bandedGridView1.Columns["MWONumber"]));
                    //}
                    //else if (projectNumber != "" && mwoNumber.ToString() != "")
                    //{
                    //    rowHandle = gridView3.LocateByValue("ProjectNumber", bandedGridView1.GetRowCellValue(e.RowHandle, "MWONumber"));
                    //}


                    //if (gridView3.IsValidRowHandle(rowHandle))
                    //{
                    //    gridView3.SetRowCellValue(rowHandle, e.Column.FieldName, e.Value);
                    //    RefreshProjectGrid();
                    //    RefreshDepartmentScheduleView();
                    //}
                    //else
                    //{
                    //    MessageBox.Show("No matching row in Project View.");
                    //}
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }

            //bandedGridView1.RefreshData();

            //MessageBox.Show("Change");
        }

        private void bandedGridView1_RowUpdated(object sender, RowObjectEventArgs e)
        {
            GridView view = sender as GridView;
            DataRowView dataRowViewObj = e.Row as DataRowView;
            WorkLoadModel wli = new WorkLoadModel();
            //MessageBox.Show(bandedGridView1.FocusedRowHandle.ToString());

            try
            {
                if (view.IsNewItemRow(e.RowHandle))
                {
                    if (!validEditorArr.ToList<string>().Exists(x => x == Environment.UserName.ToString()))
                    {
                        MessageBox.Show("This login is not authorized to make changes to work load tab.");
                        return;
                    }

                    string checkDate;

                    //MessageBox.Show(((DataRowView)obj)["ProjectNumber"].ToString());
                    wli.ToolNumber = dataRowViewObj["ToolNumber"].ToString();

                    if(int.TryParse(dataRowViewObj["MWONumber"].ToString(), out int mwoNumber))
                    {
                        wli.MWONumber = mwoNumber;
                    }
                    else if (dataRowViewObj["MWONumber"].ToString() == "")
                    {
                        wli.MWONumber = -1;
                    }
                    else
                    {
                        MessageBox.Show("Please enter a number for Tool Number.");
                        RefreshWorkloadGrid();
                        return;
                    }

                    if(int.TryParse(dataRowViewObj["ProjectNumber"].ToString(), out int projectNumber))
                    {
                        wli.ProjectNumber = projectNumber;
                    }
                    else if (dataRowViewObj["ProjectNumber"].ToString() == "")
                    {
                        wli.ProjectNumber = -1;
                    }
                    else
                    {
                        MessageBox.Show("Please enter a number for the Project Number");
                        RefreshWorkloadGrid();
                        return;
                    }

                    wli.Stage = dataRowViewObj["Stage"].ToString();
                    wli.Customer = dataRowViewObj["Customer"].ToString();
                    wli.PartName = dataRowViewObj["PartName"].ToString();

                    int.TryParse(dataRowViewObj["DeliveryInWeeks"].ToString(), out int numberOfWeeks);

                    wli.DeliveryInWeeks = numberOfWeeks;

                    checkDate = dataRowViewObj["StartDate"].ToString();

                    if (checkDate != "")
                    {
                        wli.StartDate = Convert.ToDateTime(checkDate);
                    }

                    checkDate = dataRowViewObj["FinishDate"].ToString();

                    if (checkDate != "")
                    {
                        wli.FinishDate = Convert.ToDateTime(checkDate);
                    }

                    checkDate = dataRowViewObj["AdjustedDeliveryDate"].ToString();

                    if (checkDate != "")
                    {
                        wli.FinishDate = Convert.ToDateTime(checkDate);
                    }

                    int.TryParse(dataRowViewObj["MoldCost"].ToString(), out int moldCost);

                    wli.MoldCost = moldCost;
                    wli.Engineer = dataRowViewObj["Engineer"].ToString();
                    wli.Designer = dataRowViewObj["Designer"].ToString();
                    wli.ToolMaker = dataRowViewObj["ToolMaker"].ToString();
                    wli.RoughProgrammer = dataRowViewObj["RoughProgrammer"].ToString();
                    wli.FinisherProgrammer = dataRowViewObj["FinishProgrammer"].ToString();
                    wli.ElectrodeProgrammer = dataRowViewObj["ElectrodeProgrammer"].ToString();
                    wli.Apprentice = dataRowViewObj["Apprentice"].ToString();
                    wli.Manifold = dataRowViewObj["Manifold"].ToString();
                    wli.MoldBase = dataRowViewObj["MoldBase"].ToString();
                    wli.GeneralNotes = dataRowViewObj["GeneralNotes"].ToString();

                    Database.AddWorkLoadEntry(wli);

                    gridControl2.DataSource = workLoadTableAdapter.GetData();
                    CollapseCompletedGroup();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }

            //MessageBox.Show("Row Updated.");
        }

        private void gridView2_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            MessageBox.Show("Change");
        }

        private void gridView2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Click");
        }

        private void deleteButton_Click(object sender, EventArgs e)
        {
            if (!validEditorArr.ToList<string>().Exists(x => x == Environment.UserName.ToString()))
            {
                MessageBox.Show("This login is not authorized to make changes to work load tab.");
                gridControl2.DataSource = workLoadTableAdapter.GetData();
                CollapseCompletedGroup();
                return;
            }

            Database db = new Database();

            Console.WriteLine(bandedGridView1.FocusedRowHandle);

            if(bandedGridView1.FocusedRowHandle >= 0)
            {
                if(db.DeleteWorkLoadEntry(Convert.ToInt32(bandedGridView1.GetFocusedRowCellValue("ID"))))
                {
                    db.DeleteColorEntries(Convert.ToInt32(bandedGridView1.GetFocusedRowCellValue("ID")));
                    //bandedGridView1.DeleteRow(bandedGridView1.FocusedRowHandle);
                    gridControl2.DataSource = workLoadTableAdapter.GetData();
                    CollapseCompletedGroup();
                }
                else
                {
                    MessageBox.Show("Unable to delete row.");
                }
            }
        }

        private void bandedGridView1_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
        {
            Console.WriteLine(bandedGridView1.FocusedRowHandle);
        }

        private void bandedGridView1_InitNewRow(object sender, InitNewRowEventArgs e)
        {
            // This method fires when a user BEGINS to enter data in the new item row.
            //MessageBox.Show(bandedGridView1.GetRowCellValue(e.RowHandle, "ToolNumber").ToString());
        }

        private void gridView2_RowUpdated(object sender, RowObjectEventArgs e)
        {
            MessageBox.Show("Row Updated.");
        }

        private void gridControl2_ImageChanged(object sender, System.EventArgs e)
        {
            //repositoryItemImageEdit1
            MessageBox.Show("Image changed.");
        }

        private void bandedGridView1_PrintInitialize(object sender, PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
            
            pb.PageSettings.TopMargin = 25;
            pb.PageSettings.RightMargin = 25;
            pb.PageSettings.BottomMargin = 25;
            pb.PageSettings.LeftMargin = 25;
            pb.Document.AutoFitToPagesWidth = 1;

            if (paperSize == "Tabloid")
            {
                pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.Tabloid;
                pb.PageSettings.PrinterName = @"\\S-PS1-SMDRV\P-1336 HP CP5225 - Color";
            }
            else if (paperSize == "Letter")
            {
                pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.Letter;
            }

            if (printOrientation == "Landscape")
            {
                pb.PageSettings.Landscape = true;
            }
            else if (printOrientation == "Portrait")
            {
                pb.PageSettings.Landscape = false;
            }

            bandedGridView1.OptionsPrint.RtfPageFooter = @"{\rtf1\ansi {\fonttbl\f0\ Microsoft Sans Serif;} \f0\pard \fs18 \qr \b Report Date: " + footerDateTime + @"\b0 \par}";
        }

        //private void gridView2_PrintInitialize(object sender, PrintInitializeEventArgs e)
        //{
        //    PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
        //    pb.PageSettings.Landscape = true;
        //    pb.PageMargins.Top = 50;
        //    pb.PageMargins.Right = 50;
        //    pb.PageMargins.Bottom = 50;
        //    pb.PageMargins.Left = 50;
        //}

        private void RefreshWorkloadGrid()
        {
            gridControl2.DataSource = workLoadTableAdapter.GetData();
            RoleTable = Database.GetRoleTable();
            CollapseCompletedGroup();
            footerDateTime = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
        }

        private void workLoadRefreshButton_Click(object sender, EventArgs e)
        {
            RefreshWorkloadGrid();
        }

        private void workLoadTabPrintButton_Click(object sender, EventArgs e)
        {
            // Check whether the GridControl can be previewed.
            if (!gridControl2.IsPrintingAvailable)
            {
                MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                return;
            }

            printOrientation = "Landscape";
            paperSize = "Tabloid";
            FieldInfo fi = typeof(GridColumn).GetField("minWidth", BindingFlags.NonPublic | BindingFlags.Instance);
            fi.SetValue(bandedGridView1.Columns.ColumnByFieldName("Stage"), 0);
            //bandedGridView1.Columns.ColumnByFieldName("Stage").Width = 0;
            //bandedGridView1.ShowPrintPreview();
            //bandedGridView1.Columns.ColumnByFieldName("Stage").Width = 70;
            bandedGridView1.Print();
            
            //gridView2.PrintDialog(); // Page Orientation cannot be changed.

            //gridControl2.ShowPrintPreview();
            //gridControl2.PrintDialog(); // Cannot print in landscape orientation.
            //gridControl2.Print(); // Columns are all scrunched up.
        }

        private void workLoadTabPrint2Button_Click(object sender, EventArgs e)
        {
            // Check whether the GridControl can be previewed.
            if (!gridControl2.IsPrintingAvailable)
            {
                MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                return;
            }

            printOrientation = "Portrait";
            paperSize = "Letter";
            FieldInfo fi = typeof(GridColumn).GetField("minWidth", BindingFlags.NonPublic | BindingFlags.Instance);
            fi.SetValue(bandedGridView1.Columns.ColumnByFieldName("Stage"), 0);
            //bandedGridView1.Columns.ColumnByFieldName("Stage").Width = 0;
            //bandedGridView1.ShowPrintPreview();
            //bandedGridView1.Columns.ColumnByFieldName("Stage").Width = 70;
            bandedGridView1.Print();
        }

        private void printPreviewButton_Click(object sender, EventArgs e)
        {
            printOrientation = "Landscape";
            paperSize = "Tabloid";
            bandedGridView1.ShowPrintPreview();
        }

        private void bandedGridView1_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
            //MessageBox.Show(e.Column + " " + e.RowHandle);

            //if (e.Column.FieldName == "Customer" && e.RowHandle == 0)
            //e.Appearance.ForeColor = System.Drawing.Color.Red;
        }

        private void setStatusButton_Click(object sender, EventArgs e)
        {
            //SelectStatusWindow ssw = new SelectStatusWindow();
        }

        private void bandedGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            Console.WriteLine("bandedGridView1 Mouse down event.");
            List<string> PersonnelColumns = new List<string> { "Engineer", "Designer", "ToolMaker", "RoughProgrammer", "FinishProgrammer", "ElectrodeProgrammer" };
            List<string> OtherColumns = new List<string> { "AdjustDeliveryDate", "StartDate", "FinishDate", "GeneralNotes"};
            var hitInfo = bandedGridView1.CalcHitInfo(e.Location);
            Color? color;
            Color rowColor;

            if (hitInfo.InRowCell)
            {
                int rowHandle = hitInfo.RowHandle;
                GridColumn column = hitInfo.Column;

                if (e.Button == MouseButtons.Right)
                {
                    if (!validEditorArr.ToList<string>().Exists(x => x == Environment.UserName.ToString()))
                    {
                        MessageBox.Show("This login is not authorized to make changes to work load tab.");
                        return;
                    }

                    var cells = bandedGridView1.GetSelectedCells();

                    foreach (var cell in cells)
                    {
                        bandedGridView1.UnselectCell(cell);
                    }

                    bandedGridView1.SelectCell(rowHandle, column);

                    if (rowHandle % 2 == 0)
                    {
                        rowColor = bandedGridView1.Appearance.OddRow.BackColor;
                    }
                    else
                    {
                        rowColor = bandedGridView1.Appearance.EvenRow.BackColor;
                    }

                    //MessageBox.Show("Row Handle: " + rowHandle + " Column: " + column);
                    //MessageBox.Show("Left: " + e.Location.X + MainWindow.ActiveForm.Location.X + " Top: " + e.Location.Y + MainWindow.ActiveForm.Location.Y);

                    if (PersonnelColumns.Exists(x => x == column.FieldName))
                    {
                        color = GetColorFromUser("Personnel", e.Location, rowColor);

                        cells = bandedGridView1.GetSelectedCells();

                        SetSelectedCellColor(color, cells);
                    }
                    else if (OtherColumns.Exists(x => x == column.FieldName))
                    {
                        color = GetColorFromUser("Other", e.Location, rowColor);

                        cells = bandedGridView1.GetSelectedCells();

                        SetSelectedCellColor(color, cells);
                    }
                }
                else if (e.Button == MouseButtons.Left)
                {


                }

            }
        }

        private Color? GetColorFromUser(string columnType, Point clickLocation, Color rowColor)
        {
            using (var ssw = new SelectStatusWindow(columnType, rowColor))
            {
                Point windowLocation = new Point(clickLocation.X + ssw.Width, clickLocation.Y + (int)(ssw.Height * .5));

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

        private void SetSelectedCellColor(Color? color, GridCell[] cells)
        {
            if (color == null)
            {
                return;
            }

            int projectID;

            //ColorList.Clear();

            foreach (var cell in cells)
            {
                projectID = Convert.ToInt32(bandedGridView1.GetRowCellValue(cell.RowHandle, "ID"));

                ColorStruct colorItem = ColorList.Find(r => r.Column == cell.Column.FieldName && r.ProjectID == projectID); // Somehow the same color was added twice for the same roll for the same project.

                if (colorItem == null)
                {
                    ColorList.Add(new ColorStruct {ProjectID = projectID, Column = cell.Column.FieldName, Color = (Color)color, ColorARGB = ((Color)color).ToArgb() });
                    Database.AddColorEntry(projectID, cell.Column.FieldName, ((Color)color).ToArgb());

                }
                else
                {
                    colorItem.Color = (Color)color;
                    colorItem.ColorARGB = colorItem.Color.ToArgb();

                    Database.UpdateColorEntry(projectID, cell.Column.FieldName, ((Color)color).ToArgb());
                }
            }

            bandedGridView1.LayoutChanged();
        }

        private void bandedGridView1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if(e.RowHandle >= 0 && int.TryParse(bandedGridView1.GetRowCellValue(e.RowHandle, "ID").ToString(), out int projectID))
            {
                var data = ColorList.FirstOrDefault(p => p.Column == e.Column.FieldName && p.ProjectID == projectID);

                if (data != null)
                {
                    //Console.WriteLine(e.Column + " " + e.RowHandle + " " + data.Color);
                    
                    e.Appearance.BackColor = data.Color;
                }
            }

        }

        private void AddRepositoryItemToGrid()
        {
            RichTextBox richTextBox = new RichTextBox();
            richTextBox.Dock = DockStyle.Top;

            SimpleButton fontButton = new SimpleButton();
            fontButton.Appearance.Font = new Font(fontButton.Font.FontFamily, fontButton.Font.Size, FontStyle.Regular);
            fontButton.Text = "Font";
            fontButton.Left = 40;
            fontButton.Top = 40;
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
            editorOKButton.Top = 70;
            editorOKButton.Width = 50;
            editorOKButton.Height = 30;
            editorOKButton.Dock = DockStyle.None;
            editorOKButton.Click += new EventHandler(editorOKButton_Clicked);

            SimpleButton editorCancelButton = new SimpleButton();
            editorCancelButton.Text = "Cancel";
            editorCancelButton.Left = 110;
            editorCancelButton.Top = 70;
            editorCancelButton.Width = 50;
            editorCancelButton.Height = 30;
            editorCancelButton.Dock = DockStyle.None;
            editorCancelButton.Click += new EventHandler(editorCancelButton_Clicked);

            Panel panel = new Panel();
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
            popupContainerControl.Height = 170;

            PopupContainerEdit popupContainerEdit = new PopupContainerEdit();
            popupContainerEdit.Properties.PopupControl = popupContainerControl;

            // The initialization of this instance of repositoryItemPopupContainer edit is at the top of this class.
            repositoryItemPopupContainerEdit.PopupControl = popupContainerControl;

            gridControl2.RepositoryItems.Add(repositoryItemPopupContainerEdit);
            bandedGridView1.Columns["GeneralNotes"].ColumnEdit = repositoryItemRichTextEdit;
        }

        private void repositoryItemComboBox1_BeforePopup(object sender, EventArgs e)
        {
            Console.WriteLine("repositoryItemComboBox1_BeforePopup");
        }

        private void repositoryItemComboBox1_QueryPopUp(object sender, CancelEventArgs e)
        {

        }

        private void PersonnelRepositoryComboBox_BeforePopup(object sender, EventArgs e)
        {
            Console.WriteLine("PersonnelRepositoryComboBox_BeforePopup");
        }

        private void editorOKButton_Clicked(object sender, EventArgs e)
        {
            Console.WriteLine("okButton_Clicked");

            PopupContainerEdit popupContainerEdit = bandedGridView1.ActiveEditor as PopupContainerEdit;
            RichTextBox richTextBox = popupContainerEdit.Properties.PopupControl.Controls[0] as RichTextBox;


            popupContainerEdit.EditValue = richTextBox.Rtf;
            popupContainerEdit.ClosePopup();

            //Control button = sender as Control;
            ////Close the dropdown accepting the user's choice 
            //(button.Parent.Parent as PopupContainerControl).OwnerEdit.ClosePopup();
        }

        private void editorCancelButton_Clicked(object sender, EventArgs e)
        {
            PopupContainerEdit popupContainerEdit = bandedGridView1.ActiveEditor as PopupContainerEdit;

            popupContainerEdit.CancelPopup();
        }

        private void editorColorPickerControl_ColorChanged(object sender, EventArgs e)
        {
            PopupContainerEdit popupContainerEdit = bandedGridView1.ActiveEditor as PopupContainerEdit;

            RichTextBox richTextBox = (RichTextBox)popupContainerEdit.Properties.PopupControl.Controls[0];

            ColorEdit colorEditControl = (ColorEdit)sender;

            richTextBox.SelectionColor = colorEditControl.Color;
        }

        private void boldButton_Clicked(object sender, EventArgs e)
        {
            PopupContainerEdit popupContainerEdit = bandedGridView1.ActiveEditor as PopupContainerEdit;

            RichTextBox richTextBox = (RichTextBox)popupContainerEdit.Properties.PopupControl.Controls[0];

            richTextBox.SelectionFont = new Font(richTextBox.Font.FontFamily, richTextBox.Font.Size, FontStyle.Bold);
        }

        private void underlineButton_Clicked(object sender, EventArgs e)
        {
            PopupContainerEdit popupContainerEdit = bandedGridView1.ActiveEditor as PopupContainerEdit;

            RichTextBox richTextBox = (RichTextBox)popupContainerEdit.Properties.PopupControl.Controls[0];

            richTextBox.SelectionFont = new Font(richTextBox.Font.FontFamily, richTextBox.Font.Size, FontStyle.Underline);
        }

        private void plainButton_Clicked(object sender, EventArgs e)
        {
            PopupContainerEdit popupContainerEdit = bandedGridView1.ActiveEditor as PopupContainerEdit;

            RichTextBox richTextBox = (RichTextBox)popupContainerEdit.Properties.PopupControl.Controls[0];

            richTextBox.SelectionFont = new Font(richTextBox.Font.FontFamily, richTextBox.Font.Size, FontStyle.Regular);
        }

        private void fontButton_Clicked(object sender, EventArgs e)
        {
            PopupContainerEdit popupContainerEdit = bandedGridView1.ActiveEditor as PopupContainerEdit;

            RichTextBox richTextBox = (RichTextBox)popupContainerEdit.Properties.PopupControl.Controls[0];

            FontDialog fontDialog = new FontDialog();
            fontDialog.ShowColor = true;

            if (fontDialog.ShowDialog() != DialogResult.Cancel)
            {
                richTextBox.SelectionFont = fontDialog.Font;
                richTextBox.SelectionColor = fontDialog.Color;
            }
        }

        private void bandedGridView1_ShownEditor(object sender, EventArgs e)
        {
            try
            {
                Console.WriteLine("bandedGridView1_ShownEditor entered.");
                BandedGridView bandedGridView = sender as BandedGridView;
                
                Console.WriteLine(bandedGridView.ActiveEditor.EditorTypeName);
                Database db = new Database();

                ColumnView columnView = sender as ColumnView;
                PopupContainerEdit popupContainerEdit = null;
                ComboBoxEdit comboBoxEdit = null;

                if (bandedGridView != null)
                {

                    if (bandedGridView.ActiveEditor.EditorTypeName == "PopupContainerEdit")
                    {
                        popupContainerEdit = bandedGridView.ActiveEditor as PopupContainerEdit;
                    }
                    else if (bandedGridView.ActiveEditor.EditorTypeName == "ComboBoxEdit")
                    {
                        comboBoxEdit = bandedGridView.ActiveEditor as ComboBoxEdit;
                    }

                    if (popupContainerEdit != null)
                    {

                        //RichEditControl richEditControl = (RichEditControl)activeEditor.Properties.PopupControl.Controls[0];

                        //richEditControl.ActiveViewType = RichEditViewType.PrintLayout;
                        //richEditControl.ActiveView.ZoomFactor = 2f;
                        //richEditControl.Document.Sections[0].Margins.Left = 50;
                        //richEditControl.Document.Sections[0].Margins.Top = 50;

                        RichTextBox richTextBox = (RichTextBox)popupContainerEdit.Properties.PopupControl.Controls[0];

                        richTextBox.Rtf = bandedGridView1.GetFocusedRowCellValue("GeneralNotes").ToString();

                        //activeEditor.QueryResultValue += new QueryResultValueEventHandler(this.popupContainerEdit_QueryResultValue);
                    }

                    if (comboBoxEdit != null)
                    {
                        string column = bandedGridView1.FocusedColumn.FieldName;
                        string personnelType = "";

                        comboBoxEdit.Properties.Items.Clear();

                        if (column == "RoughProgrammer")
                        {
                            personnelType = "Rough Programmer";
                        }
                        else if (column == "FinishProgrammer")
                        {
                            personnelType = "Finish Programmer";
                        }
                        else if (column == "ElectrodeProgrammer")
                        {
                            personnelType = "Electrode Programmer";
                        }
                        else if (column == "ToolMaker")
                        {
                            personnelType = "Tool Maker";
                        }
                        else if (column == "Stage")
                        {
                            comboBoxEdit.Properties.Items.Add("1 - In-Design");
                            comboBoxEdit.Properties.Items.Add("2 - In-Programming");
                            comboBoxEdit.Properties.Items.Add("3 - In-Shop");
                            comboBoxEdit.Properties.Items.Add("4 - In-Mold Check-In or Outside Vendors");
                            comboBoxEdit.Properties.Items.Add("5 - Rework");
                            comboBoxEdit.Properties.Items.Add("6 - In-Repair / Development");
                            comboBoxEdit.Properties.Items.Add("7 - Completed");
                            comboBoxEdit.Properties.Items.Add("");

                            return;
                        }

                        if (personnelType != "")
                        {
                            comboBoxEdit.Properties.Items.Add("");
                            comboBoxEdit.Properties.Items.AddRange(GetResourceList(personnelType).ToArray());
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void bandedGridView1_CustomRowCellEditForEditing(object sender, CustomRowCellEditEventArgs e)
        {
            if (e.Column.FieldName == "GeneralNotes")
            {
                e.RepositoryItem = repositoryItemPopupContainerEdit;
            }
        }

        private void popupContainerEdit_QueryResultValue(object sender, QueryResultValueEventArgs e)
        {
            PopupContainerEdit popupContainerEdit = bandedGridView1.ActiveEditor as PopupContainerEdit;
            RichTextBox richTextBox = popupContainerEdit.Properties.PopupControl.Controls[0] as RichTextBox;
            
            popupContainerEdit.EditValue = richTextBox.Rtf;
        }

        #endregion

        private List<string> GetResourceList(string role)
        {
            List<string> resourceList = new List<string>();

            var result = from roleTable in RoleTable.AsEnumerable()
                         where roleTable.Field<string>("Role") == role
                         select roleTable;

            //resourceList.Add("");

            foreach (var resource in result)
            {
                resourceList.Add(resource.Field<string>("ResourceName"));
            }

            if (role == "")
            {
                var result2 = from roleTable in RoleTable.AsEnumerable()
                              where roleTable.Field<string>("ResourceType") == "Person"
                              group roleTable by roleTable.Field<string>("ResourceName") into grp
                              orderby grp.Key
                              select grp;

                foreach (var resource in result2)
                {
                    resourceList.Add(resource.Key);
                }
            }

            return resourceList;
        }

        private void SchedulerControl1_Click(object sender, EventArgs e)
        {

        }

        private void RepositoryItemHyperLinkEdit2_ButtonClick(object sender, ButtonPressedEventArgs e)
        {
            //Excel.
        }

        private (string jobNumber, int projectNumber) GetComboBoxInfo()
        {
            string[] jobNumberComboBoxText, jobNumberComboBoxText2;

            jobNumberComboBoxText = projectComboBox.Text.Split(' ');
            jobNumberComboBoxText2 = projectComboBox.Text.Split('#');

            return (jobNumberComboBoxText[0], Convert.ToInt32(jobNumberComboBoxText2[1]));
        }

        private void appointmentsFilterControl_FilterChanged(object sender, FilterChangedEventArgs e)
        {
            //schedulerControl1.DataStorage.Appointments.Filter = appointmentsFilterControl.FilterString;
        }

        private void schedulerControl1_EditAppointmentFormShowing(object sender, AppointmentFormEventArgs e)
        {

        }

        private void PopulateDepartmentComboBoxes()
        {
            List<string> departmentList1 = new List<string> { "Programming", "Program Rough", "Program Finish", "Program Electrodes", "CNCs", "CNC People", "CNC Rough", "CNC Finish", "CNC Electrodes", "Grind", "Inspection", "EDM Sinker", "EDM Wire (In-House)", "Polish", "All" };
            List<string> departmentList2 = new List<string> {"Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "CNC Electrodes", "Grind", "Inspection", "EDM Sinker", "EDM Wire (In-House)", "Polish", "All" };

            departmentComboBox.Properties.Items.AddRange(departmentList1);
            departmentComboBox2.Properties.Items.AddRange(departmentList2);
        }

        private void PopulateEmployeeComboBox()
        {
            var result = (from empList in this.workload_Tracking_System_DBDataSet.Tasks.AsEnumerable()
                          where empList.Field<string>("Resource") != null && empList.Field<string>("Resource") != "" && empList.Field<string>("Status") != null
                          orderby empList.Field<string>("Resource")
                          select empList.Field<string>("Resource")).Distinct().ToList();

            PrintEmployeeWorkCheckedComboBoxEdit.Properties.Items.Clear();

            foreach (var item in result)
            {
                PrintEmployeeWorkCheckedComboBoxEdit.Properties.Items.Add(item);    
            }
        }

        private bool UpdateTaskStorage2(Appointment apt)
        {
            Database db = new Database();
            ComponentModel component;
            string componentName;
            TaskModel task;
            int taskID;

            var number = GetComboBoxInfo();

            //Resource resource = schedulerStorage2.Resources[Convert.ToInt16(apt.ResourceId) - 1];

            //resource = schedulerStorage2.Resources[Convert.ToInt16(resource.ParentId) - 1];

            
            componentName = apt.CustomFields["Component"].ToString();

            component = Project.ComponentList.Find(x => x.Name == componentName);

            taskID = Convert.ToInt16(apt.CustomFields["TaskID"]);

            task = component.TaskList[Convert.ToInt16(apt.CustomFields["TaskID"]) - 1];

            if (db.UpdateTask(number.jobNumber, number.projectNumber, component.Name, task.ID, apt.Start, apt.End, Project.OverlapAllowed))
            {
                task.SetDates(apt.Start, apt.End);
                return true;
            }
            else
            {
                return false;
            }
        }

        private void chartControl1_CustomDrawCrosshair(object sender, CustomDrawCrosshairEventArgs e)
        {
            foreach (CrosshairElementGroup group in e.CrosshairElementGroups)
            {
                foreach (CrosshairElement element in group.CrosshairElements)
                {
                    //SeriesPoint currentPoint = element.SeriesPoint;

                    //if (currentPoint.Tag.GetType() == typeof(DataRowView))
                    //{
                    //    DataRowView rowView = (DataRowView)currentPoint.Tag;
                    //    string s = "Test";

                    //    element.LabelElement.Text = s;

                    //}
                }
            }
        }
    }
}
