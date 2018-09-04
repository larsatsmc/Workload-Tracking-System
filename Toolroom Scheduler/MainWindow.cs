using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ClassLibrary;

namespace Toolroom_Scheduler
{
    public partial class MainWindow : Form
    {
        // The real database.
        private static string ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;
        Data Source=X:\TOOLROOM\Workload Tracking System\Database\Workload Tracking System DB.accdb";
		// Test database.  Idea: Copy real database over new database so that we're working off of an identical copy of the real database.
		private static string connectionStringt = @"Provider=Microsoft.ACE.OLEDB.12.0;
        Data Source=X:\TOOLROOM\Josh Meservey\Workload Tracking System\Workload Tracking System.accdb";

        private static string QueryString { get; set; }
        private static OleDbConnection Connection = new OleDbConnection(ConnectionString);
        private static DataTable DataTable = new DataTable();

		private bool FormLoading = false;
        DateTimePicker oDateTimePicker;

        public MainWindow()
        {
            FormLoading = true;
            InitializeComponent();

            ProcessTabs.SelectTab("All");

            SetToolTips();
            PopulateSelectedTab();
            PopulateJobNumberComboBox();
            //populateDesignerView();
            //populateProgrammerView();
            //populateRoughMachineView();
            //populateFinishMachineView();
            //populateElectrodeMachineView();
            //populateEDMMachineView();
            //populateInspectionMachineView();
            FormLoading = false;
            Console.WriteLine("MainWindow Complete");
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            SetQueryString();
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            adapter.SelectCommand = new OleDbCommand(QueryString, Connection);
            adapter.Fill(DataTable);
            DataGridView1.DataSource = DataTable;
            PopulateDataGridViewComboboxes();
            Console.WriteLine("Form_Load Complete");
        }

        public void SetToolTips()
        {
            InfoTip.SetToolTip(CreateKanBanButton, "Click this button to create Kan Ban sheet from selected project. NOTE: All project info must be first be entered.");
            InfoTip.SetToolTip(BulkAssignButton, "Click this button to bulk assign resources for a selected project.");
            InfoTip.SetToolTip(ForwardDateButton, "Click this button to check what tasks are complete and update database.");
            InfoTip.SetToolTip(JobNumberComboBox, "This Combo Box is used to selectively display tasks for a given project.");
            InfoTip.SetToolTip(ManageResourcesButton, "Click this button to manage what resources (people / machines) show up in drop down menus.");
            InfoTip.SetToolTip(removeProjectButton, "Click this button to remove the selected project from the database.");
            InfoTip.SetToolTip(EditProjectButton, "Click this button to load a completed MS Project file into the database.");
            InfoTip.SetToolTip(CreateProjectButton, "Click this button to begin entering data to setup a MS Project file.");
        }

        public string GetJobNumberComboBoxValue()
        {
            return JobNumberComboBox.Text;
        }

        private (string jobNumber, int projectNumber) GetComboBoxInfo()
        {
            string[] jobNumberComboBoxText, jobNumberComboBoxText2;

            jobNumberComboBoxText = JobNumberComboBox.Text.Split(' ');
            jobNumberComboBoxText2 = JobNumberComboBox.Text.Split('#');

            return (jobNumberComboBoxText[0], Convert.ToInt32(jobNumberComboBoxText2[1]));
        }

        private void SetQueryString()
        {
            if(ProcessTabs.SelectedTab.Text == "Design")
            {
                QueryString = "SELECT * FROM Tasks WHERE TaskName LIKE '%Design%' ORDER BY ID";
            }
            else if(ProcessTabs.SelectedTab.Text == "Programming")
            {
                QueryString = "SELECT * FROM Tasks WHERE TaskName LIKE '%Program%' ORDER BY ID";
            }
            else if(ProcessTabs.SelectedTab.Text == "Roughing")
            {
                QueryString = "SELECT * FROM Tasks WHERE TaskName = 'CNC Rough' ORDER BY ID";
            }
            else if(ProcessTabs.SelectedTab.Text == "Finishing")
            {
                QueryString = "SELECT * FROM Tasks WHERE TaskName LIKE 'CNC Finish' ORDER BY ID";
            }
            else if(ProcessTabs.SelectedTab.Text == "Electrodes")
            {
                QueryString = "SELECT * FROM Tasks WHERE TaskName LIKE 'CNC Electrodes' ORDER BY ID";
            }
            else if(ProcessTabs.SelectedTab.Text == "EDM")
            {
                QueryString = "SELECT * FROM Tasks WHERE TaskName = 'EDM Sinker' ORDER BY ID";
            }
            else if(ProcessTabs.SelectedTab.Text == "Inspection")
            {
                QueryString = "SELECT * FROM Tasks WHERE TaskName LIKE '%Inspection%' ORDER BY ID";
            }
            else if (ProcessTabs.SelectedTab.Text == "All")
            {
                QueryString = "SELECT * FROM Tasks ORDER BY ID";
            }
        }

        public void RefreshDataGridView([CallerMemberName]string CallerName = "")
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter(QueryString, Connection);

            int rowIndex = 0;
            int colIndex = 0;
            int scrollRowPosition = 0;
            int scrollColPosition = 0;

            try
            {
                //MessageBox.Show("Automation");
                if (DataGridView1.Rows.Count != 0)
                {
                    rowIndex = DataGridView1.CurrentRow.Index;
                    colIndex = DataGridView1.CurrentCellAddress.X;
                }
                else
                {
                    rowIndex = 0;
                    colIndex = 0;
                }

                scrollRowPosition = DataGridView1.FirstDisplayedScrollingRowIndex;
                scrollColPosition = DataGridView1.FirstDisplayedScrollingColumnIndex;
                DataTable.Clear();
                adapter.SelectCommand = new OleDbCommand(QueryString, Connection);
                //Console.WriteLine(adapter.SelectCommand.CommandText);
                adapter.Fill(DataTable);
                DataGridView1.DataSource = DataTable;
                //foreach (DataRow nrow in datatable1.Rows)
                //{
                //    Console.WriteLine(nrow["Resource"]);
                //}
                Console.WriteLine(CallerName);

				if (DataGridView1.Rows.Count != 0)
				{
					if (CallerName == "ResetButton_Click")
					{
						DataGridView1.CurrentCell = DataGridView1.Rows[0].Cells[1];
					}
					else if (CallerName == "ProcessTabs_Selected")
					{
						DataGridView1.CurrentCell = DataGridView1.Rows[0].Cells[1];
					}
					else
					{
						DataGridView1.CurrentCell = DataGridView1.Rows[rowIndex].Cells[colIndex];
						DataGridView1.FirstDisplayedScrollingRowIndex = scrollRowPosition;
						//DataGridView1.FirstDisplayedScrollingColumnIndex = scrollColPosition;
					}
				}

                if (scrollRowPosition < 0)
                {
                    scrollRowPosition = 0;
                }

                //showTalliedData();

            }
            catch (Exception er)
            {
                MessageBox.Show(er.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private DateTime AddBusinessDays(DateTime date, string durationSt)
        {
            int days;
            string[] duration = durationSt.Split(' ');
            days = Convert.ToInt16(duration[0]);

            if (days < 0)
            {
                throw new ArgumentException("days cannot be negative", "days");
            }

            if (days == 0) return date;

            if (date.DayOfWeek == DayOfWeek.Saturday)
            {
                date = date.AddDays(2);
                days -= 1;
            }
            else if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = date.AddDays(1);
                days -= 1;
            }

            date = date.AddDays(days / 5 * 7);
            int extraDays = days % 5;

            if ((int)date.DayOfWeek + extraDays > 5)
            {
                extraDays += 2;
            }

            return date.AddDays(extraDays);

        }

        private void RemoveProject()
        {
            Database db = new Database();
            int selectedIndex;

            if(JobNumberComboBox.Text == "All")
            {
                //MessageBox.Show("Just select a single job.");

                OpenFileDialog openProjectsReport = new OpenFileDialog();
                openProjectsReport.InitialDirectory = @"C:\User\" + Environment.UserName + @"\Downloads\";
                openProjectsReport.Filter = "Excel Files (*.xls)|*.xls";

                if (openProjectsReport.ShowDialog() == DialogResult.OK)
                {
                    ReadClosedProjectsReport(openProjectsReport?.FileName);
                }
                
                return;
            }

            DialogResult result = MessageBox.Show("Are you sure you want to remove this project?", "Warning.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if(result == DialogResult.Yes)
            {

            }
            else if(result == DialogResult.No)
            {
                return;
            }
            var number = GetComboBoxInfo();

            if (JobNumberComboBox.SelectedIndex < JobNumberComboBox.Items.Count - 1)
            {
                selectedIndex = JobNumberComboBox.SelectedIndex; // Indexes in comboboxes are base-zero.           
            }
            else
            {
                selectedIndex = JobNumberComboBox.SelectedIndex - 1;               
            }

            db.ClearAllProjectData(number.jobNumber, number.projectNumber);
            RefreshDataGridView();
            PopulateJobNumberComboBox();

            JobNumberComboBox.SelectedIndex = selectedIndex;  // Test out to make sure error is not thrown when last item in combobox is selected for deletion.
        }

        private List<int> ReadClosedProjectsReport(string filePath)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet closedProjectSheet = workbook.Sheets[1];
            Excel.Worksheet closedMWOSheet = workbook.Sheets[2];
            List<int> projectNums = new List<int>();

            for (int r = 5; r < closedProjectSheet.UsedRange.Rows.Count; r++)
            {
                int.TryParse(closedProjectSheet.Cells[r, 1].Text, out int n);

                if (n != 0) { projectNums.Add(n); }
            }

            for (int r = 11; r < closedMWOSheet.UsedRange.Rows.Count; r++)
            {
                int.TryParse(closedMWOSheet.Cells[r, 2].Text, out int n);

                if(n != 0) {projectNums.Add(n);}
            }

            //foreach (int item in projectNums)
            //{
            //    Console.WriteLine(item);
            //}
           
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(closedProjectSheet);
            Marshal.ReleaseComObject(closedMWOSheet);

            workbook.Close();
            Marshal.ReleaseComObject(workbook);

            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            return projectNums;
        }

        private void FlagCompletedProjects(List<int> completedProjects)
        {
            foreach (string item in JobNumberComboBox.Items)
            {

                Console.WriteLine(item.Split('#')[1]);
            }
        }

        private bool KanBanExists(string jobNumber, int projectNumber)
        {
            Database db = new Database();
            string kanBanWorkbookPath = db.GetKanBanWorkbookPath(jobNumber, projectNumber);

            if(kanBanWorkbookPath != "")
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

        private void PopulateDataGridViewComboboxes()
        {
            Database db = new Database();
            List<ResourceInfo> ResourceList = new List<ResourceInfo>();
            List<string> SelectedList = null;
            List<string> DesignerList = new List<string> { "Phil Morris", "Brian Yoder", "Lee Meservey", "Jim Schmidt", " " };
            List<string> ProgrammerList = new List<string> { "Josh Meservey", "Shawn Swiggum", "Alex Anderson", "Rod Shilts", "Derek Timm", "Micah Bruns", "Ed Mendez", "John Gruntner", " " };
            List<string> RoughMachineList = new List<string> { "Mazak 1", "Mazak 2", "Mazak 3", "Haas", " " }; 
            List<string> FinishMachineList = new List<string> { "950 Yasda", "Old 640 Yasda", "New 640 Yasda", "430 Yasda", "Mazak 1", "Mazak 2", "Mazak 3", " " };
            List<string> GraphiteMachineList = new List<string> { "Makino", "Sodick", " " };
            List<string> EDMMachineList = new List<string> { "Sodick 1", "Sodick 2", " " };
            List<string> InspectionMachineList = new List<string> { "Brown & Sharpe", " "};
            List<string> StatusList = new List<string> { "Waiting", "In Progress", "On Hold", "Completed", " " };
            List<string> AllList = null;

            AllList = DesignerList.ToList();
            AllList.AddRange(ProgrammerList);
            AllList.AddRange(RoughMachineList);
            AllList.AddRange(FinishMachineList);
            AllList.AddRange(GraphiteMachineList);
            AllList.AddRange(EDMMachineList);
            AllList.AddRange(InspectionMachineList);

            try
            {

                if (ProcessTabs.SelectedTab.Text == "Design")
                {
                    //SelectedList = DesignerList.ToList();
                    SelectedList = db.GetResourceList("Designer");
                }
                else if (ProcessTabs.SelectedTab.Text == "Programming")
                {
                    //SelectedList = ProgrammerList.ToList();
                    SelectedList = new List<string>();

                    //ResourceList = db.GetResourceList("Programmer");
                    //ResourceList = ResourceList.Where(r => r.Role.Contains("Programmer")).GroupBy(r => new { r.FirstName, r.LastName }).Select(r => r.First()).ToList();

                    //foreach (var resource in ResourceList)
                    //{
                    //    Console.WriteLine($" {resource.FirstName} {resource.LastName} {resource.Role} ");
                    //    SelectedList.Add($"{resource.FirstName} {resource.LastName}");
                    //}
                    
                }
                else if (ProcessTabs.SelectedTab.Text == "Roughing")
                {
                    SelectedList = RoughMachineList.ToList();
                }
                else if (ProcessTabs.SelectedTab.Text == "Finishing")
                {
                    SelectedList = FinishMachineList.ToList();
                }
                else if (ProcessTabs.SelectedTab.Text == "Electrodes")
                {
                    SelectedList = GraphiteMachineList.ToList();
                }
                else if (ProcessTabs.SelectedTab.Text == "EDM")
                {
                    SelectedList = EDMMachineList.ToList();
                }
                else if (ProcessTabs.SelectedTab.Text == "Inspection")
                {
                    SelectedList = InspectionMachineList.ToList();
                }
                else if (ProcessTabs.SelectedTab.Text == "All")
                {
                    SelectedList = AllList.ToList();
                }

                //foreach (string name in SelectedList)
                //{
                //    Console.WriteLine(name);
                //}
                
                //(FinishingDataGridView.Columns[10] as DataGridViewComboBoxColumn).DataSource = null;
                (DataGridView1.Columns["Resource"] as DataGridViewComboBoxColumn).DataSource = SelectedList;
                (DataGridView1.Columns["Status"] as DataGridViewComboBoxColumn).DataSource = StatusList;

                //    for (int i = 0; i < FinishingDataGridView.Rows.Count; i++)
                //    {
                //        //DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)(FinishingDataGridView.Rows[i]).Cells[10];
                //        if (FinishingDataGridView.Rows[i].Cells[1].Value.ToString() == "CNC Rough")
                //        {
                //            (FinishingDataGridView.Rows[i].Cells[10] as DataGridViewComboBoxCell).DataSource = RoughMachineList;
                //        }
                //        else if (FinishingDataGridView.Rows[i].Cells[1].Value.ToString() == "CNC Finish")
                //        {
                //            (FinishingDataGridView.Rows[i].Cells[10] as DataGridViewComboBoxCell).DataSource = FinishMachineList;
                //        }
                //        else if (FinishingDataGridView.Rows[i].Cells[1].Value.ToString() == "CNC Electrodes")
                //        {
                //            (FinishingDataGridView.Rows[i].Cells[10] as DataGridViewComboBoxCell).DataSource = GraphiteMachineList;
                //        }
                //        else if (FinishingDataGridView.Rows[i].Cells[1].Value.ToString() == "EDM Sinker")
                //        {
                //            (FinishingDataGridView.Rows[i].Cells[10] as DataGridViewComboBoxCell).DataSource = EDMMachineList;
                //        }

                //    }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void PopulateJobNumberComboBox()
        {
            Database db = new Database();

            JobNumberComboBox.DataSource = db.GetJobNumberComboList();
        }

        private void PopulateProgrammerView()
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;
            string programmer1 = "Josh Meservey";
            string programmer2 = "Shawn Swiggum";
            string programmer3 = "Alex Anderson";
            string programmer4 = "Rod Shilts";
            string programmer5 = "Ben Meservey";
            string programmer6 = "Derek Timm";
            string programmer7 = "Micah Bruins";

            //DateTime dt = new DateTime();

            queryString = "SELECT * FROM Tasks WHERE (StartDate <= @currentDate AND FinishDate >= @currentDate) AND TaskName CONTAINS '%@currentTask%' ORDER BY StartDate ASC, Priority ASC";
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            //dt = FinishingDateTimePicker.Value;
            adapter.SelectCommand.Parameters.AddWithValue("@currentDate", ProgrammerDateTimePicker.Value.ToString("d"));
            adapter.SelectCommand.Parameters.AddWithValue("@currentTask", "Program");
            adapter.Fill(datatable);

            //populateToolnumberComboBox(datatable);

            JobLabelProgrammer1.Text = "Job: ";
            ComponentLabelProgrammer1.Text = "Component: ";

            JobLabelProgrammer2.Text = "Job: ";
            ComponentLabelProgrammer2.Text = "Component: ";

            JobLabelProgrammer3.Text = "Job: ";
            ComponentLabelProgrammer3.Text = "Component: ";

            JobLabelProgrammer4.Text = "Job: ";
            ComponentLabelProgrammer4.Text = "Component: ";

            JobLabelProgrammer5.Text = "Job: ";
            ComponentLabelProgrammer5.Text = "Component: ";

            JobLabelProgrammer6.Text = "Job: ";
            ComponentLabelProgrammer6.Text = "Component: ";

            JobLabelProgrammer7.Text = "Job: ";
            ComponentLabelProgrammer7.Text = "Component: ";

            ListBoxProgrammer1.Items.Clear();
            ListBoxProgrammer2.Items.Clear();
            ListBoxProgrammer3.Items.Clear();
            ListBoxProgrammer4.Items.Clear();
            ListBoxProgrammer5.Items.Clear();
            ListBoxProgrammer6.Items.Clear();
            ListBoxProgrammer7.Items.Clear();
            ListBoxProgrammer8.Items.Clear();

            foreach (DataRow nrow in datatable.Rows)
            {
                if (nrow["Status"].ToString() != "In Progress" && nrow["Status"].ToString() != "On Hold")
                {
                    if (nrow["Resource"].ToString() == programmer1)
                    {
                        ListBoxProgrammer1.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == programmer2)
                    {
                        ListBoxProgrammer2.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == programmer3)
                    {
                        ListBoxProgrammer3.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == programmer4)
                    {
                        ListBoxProgrammer4.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == programmer5)
                    {
                        ListBoxProgrammer5.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == programmer6)
                    {
                        ListBoxProgrammer6.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == programmer7)
                    {
                        ListBoxProgrammer7.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                }
                else if (nrow["Status"].ToString() == "In Progress")
                {
                    if (nrow["Resource"].ToString() == programmer1)
                    {
                        JobLabelProgrammer1.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelProgrammer1.Text = "Component: " + nrow["Component"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == programmer2)
                    {
                        JobLabelProgrammer2.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelProgrammer2.Text = "Component: " + nrow["Component"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == programmer3)
                    {
                        JobLabelProgrammer3.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelProgrammer3.Text = "Component: " + nrow["Component"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == programmer4)
                    {
                        JobLabelProgrammer4.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelProgrammer4.Text = "Component: " + nrow["Component"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == programmer5)
                    {
                        JobLabelProgrammer5.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelProgrammer5.Text = "Component: " + nrow["Component"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == programmer6)
                    {
                        JobLabelProgrammer6.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelProgrammer6.Text = "Component: " + nrow["Component"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == programmer7)
                    {
                        JobLabelProgrammer7.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelProgrammer7.Text = "Component: " + nrow["Component"].ToString();
                    }
                }

                //Console.WriteLine(nrow["JobNumber"] + " " + nrow["Component"] + " " + nrow["Resource"]);
            }
        }

        private void PopulateDesignerView()
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;
            string designer1 = "Phil Morris";
            string designer2 = "Brian Yoder";
            string designer3 = "Lee Meservey";
            string designer4 = "Jim Schmidt";

            //DateTime dt = new DateTime();

            queryString = "SELECT * FROM Tasks WHERE (StartDate <= @currentDate AND FinishDate >= @currentDate) AND TaskName = @currentTask ORDER BY StartDate ASC, Priority ASC";

            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            //dt = FinishingDateTimePicker.Value;
            adapter.SelectCommand.Parameters.AddWithValue("@currentDate", DesignerDateTimePicker.Value.ToString("d"));
            adapter.SelectCommand.Parameters.AddWithValue("@currentTask", "Design / Make Drawings");
            adapter.Fill(datatable);

            //populateToolnumberComboBox(datatable);

            JobLabelDesigner1.Text = "Job: ";
            ComponentLabelDesigner1.Text = "Component: ";

            JobLabelDesigner2.Text = "Job: ";
            ComponentLabelDesigner2.Text = "Component: ";

            JobLabelDesigner3.Text = "Job: ";
            ComponentLabelDesigner3.Text = "Component: ";

            JobLabelDesigner4.Text = "Job: ";
            ComponentLabelDesigner4.Text = "Component: ";

            ListBoxDesigner1.Items.Clear();
            ListBoxDesigner2.Items.Clear();
            ListBoxDesigner3.Items.Clear();
            ListBoxDesigner4.Items.Clear();

            foreach (DataRow nrow in datatable.Rows)
            {
                if (nrow["Status"].ToString() != "In Progress" && nrow["Status"].ToString() != "On Hold")
                {
                    if (nrow["Resource"].ToString() == designer1)
                    {
                        ListBoxDesigner1.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == designer2)
                    {
                        ListBoxDesigner2.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == designer3)
                    {
                        ListBoxDesigner3.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == designer4)
                    {
                        ListBoxDesigner4.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                }
                else if (nrow["Status"].ToString() == "In Progress")
                {
                    if (nrow["Resource"].ToString() == designer1)
                    {
                        JobLabelDesigner1.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelDesigner1.Text = "Component: " + nrow["Component"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == designer2)
                    {
                        JobLabelDesigner2.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelDesigner2.Text = "Component: " + nrow["Component"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == designer3)
                    {
                        JobLabelDesigner3.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelDesigner3.Text = "Component: " + nrow["Component"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == designer4)
                    {
                        JobLabelDesigner4.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelDesigner4.Text = "Component: " + nrow["Component"].ToString();
                    }
                }

                //Console.WriteLine(nrow["JobNumber"] + " " + nrow["Component"] + " " + nrow["Resource"]);
            }
        }

        private void PopulateRoughMachineView()
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;
            string roughingMachine1 = "Mazak 1";
            string roughingMachine2 = "Mazak 2";
            string roughingMachine3 = "Haas";
            string roughingMachine4 = "Mazak 3";
            //DateTime dt = new DateTime();

            queryString = "SELECT * FROM Tasks WHERE ((StartDate <= @currentDate AND FinishDate >= @currentDate) AND TaskName = @currentTask)" +
                                                "  OR ((StartDate <= @currentDate AND FinishDate >= @currentDate) AND TaskName = @currentTask2) " +
                                                "ORDER BY StartDate ASC, Priority ASC";
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            //dt = FinishingDateTimePicker.Value;
            adapter.SelectCommand.Parameters.AddWithValue("@currentDate", RoughingDateTimePicker.Value.ToString("d"));
            adapter.SelectCommand.Parameters.AddWithValue("@currentTask", "CNC Rough");
            adapter.SelectCommand.Parameters.AddWithValue("@currentTask2", "CNC Finish");
            adapter.Fill(datatable);

            //populateToolnumberComboBox(datatable);

            JobLabelRoughingMachine1.Text = "Job: ";
            ComponentLabelRoughingMachine1.Text = "Component: ";
            OperatorLabelRoughingMachine1.Text = "Operator: ";

            JobLabelRoughingMachine2.Text = "Job: ";
            ComponentLabelRoughingMachine2.Text = "Component: ";
            OperatorLabelRoughingMachine2.Text = "Operator: ";

            JobLabelRoughingMachine3.Text = "Job: ";
            ComponentLabelRoughingMachine3.Text = "Component: ";
            OperatorLabelRoughingMachine3.Text = "Operator: ";

            JobLabelRoughingMachine4.Text = "Job: ";
            ComponentLabelRoughingMachine4.Text = "Component: ";
            OperatorLabelRoughingMachine4.Text = "Operator: ";

            ListBoxRoughMachine1.Items.Clear();
            ListBoxRoughMachine2.Items.Clear();
            ListBoxRoughMachine3.Items.Clear();
            ListBoxRoughMachine4.Items.Clear();

            foreach (DataRow nrow in datatable.Rows)
            {
                if (nrow["Status"].ToString() != "In Progress" && nrow["Status"].ToString() != "On Hold")
                {
                    if (nrow["Resource"].ToString() == roughingMachine1)
                    {
                        ListBoxRoughMachine1.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == roughingMachine2)
                    {
                        ListBoxRoughMachine2.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == roughingMachine3)
                    {
                        ListBoxRoughMachine3.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == roughingMachine4)
                    {
                        ListBoxRoughMachine4.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                }
                else if (nrow["Status"].ToString() == "In Progress")
                {
                    if (nrow["Resource"].ToString() == roughingMachine1)
                    {
                        JobLabelRoughingMachine1.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelRoughingMachine1.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelRoughingMachine1.Text = "Operator: " + nrow["Operator"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == roughingMachine2)
                    {
                        JobLabelRoughingMachine2.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelRoughingMachine2.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelRoughingMachine2.Text = "Operator: " + nrow["Operator"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == roughingMachine3)
                    {
                        JobLabelRoughingMachine3.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelRoughingMachine3.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelRoughingMachine3.Text = "Operator: " + nrow["Operator"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == roughingMachine4)
                    {
                        JobLabelRoughingMachine4.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelRoughingMachine4.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelRoughingMachine4.Text = "Operator: " + nrow["Operator"].ToString();
                    }
                }

                //Console.WriteLine(nrow["JobNumber"] + " " + nrow["Component"] + " " + nrow["Resource"]);
            }
        }

        private void PopulateFinishMachineView()
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;
            string finishingMachine1 = "950 Yasda";
            string finishingMachine2 = "Old 640 Yasda";
            string finishingMachine3 = "New 640 Yasda";
            string finishingMachine4 = "430 Yasda";
            //DateTime dt = new DateTime();

            queryString = "SELECT * FROM Tasks WHERE (StartDate <= @currentDate AND FinishDate >= @currentDate) AND TaskName = @currentTask ORDER BY StartDate ASC, Priority ASC";
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            //dt = FinishingDateTimePicker.Value;
            adapter.SelectCommand.Parameters.AddWithValue("@currentDate", FinishingDateTimePicker.Value.ToString("d"));
            adapter.SelectCommand.Parameters.AddWithValue("@currentTask", "CNC Finish");
            adapter.Fill(datatable);

            //populateToolnumberComboBox(datatable);

            JobLabelFinishingMachine4.Text = "Job: ";
            ComponentLabelFinishingMachine4.Text = "Component: ";
            OperatorLabelFinishingMachine4.Text = "Operator: ";

            JobLabelFinishingMachine3.Text = "Job: ";
            ComponentLabelFinishingMachine3.Text = "Component: ";
            OperatorLabelFinishingMachine3.Text = "Operator: ";

            JobLabelFinishingMachine2.Text = "Job: ";
            ComponentLabelFinishingMachine2.Text = "Component: ";
            OperatorLabelFinishingMachine2.Text = "Operator: ";

            JobLabelFinishingMachine1.Text = "Job: ";
            ComponentLabelFinishingMachine1.Text = "Component: ";
            OperatorLabelFinishingMachine1.Text = "Operator: ";

            ListBoxFinishMachine1.Items.Clear();
            ListBoxFinishMachine2.Items.Clear();
            ListBoxFinishMachine3.Items.Clear();
            ListBoxFinishMachine4.Items.Clear();

            foreach (DataRow nrow in datatable.Rows)
            {
                if (nrow["Status"].ToString() != "In Progress" && nrow["Status"].ToString() != "On Hold")
                {
                    if(nrow["Resource"].ToString() == finishingMachine1)
                    {
                        ListBoxFinishMachine1.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == finishingMachine2)
                    {
                        ListBoxFinishMachine2.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == finishingMachine3)
                    {
                        ListBoxFinishMachine3.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == finishingMachine4)
                    {
                        ListBoxFinishMachine4.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                }
                else if (nrow["Status"].ToString() == "In Progress")
                {
                    if (nrow["Resource"].ToString() == finishingMachine1)
                    {
                        JobLabelFinishingMachine1.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelFinishingMachine1.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelFinishingMachine1.Text = "Operator: " + nrow["Operator"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == finishingMachine2)
                    {
                        JobLabelFinishingMachine2.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelFinishingMachine2.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelFinishingMachine2.Text = "Operator: " + nrow["Operator"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == finishingMachine3)
                    {
                        JobLabelFinishingMachine3.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelFinishingMachine3.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelFinishingMachine3.Text = "Operator: " + nrow["Operator"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == finishingMachine4)
                    {
                        JobLabelFinishingMachine4.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelFinishingMachine4.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelFinishingMachine4.Text = "Operator: " + nrow["Operator"].ToString();
                    }
                }

                //Console.WriteLine(nrow["JobNumber"] + " " + nrow["Component"] + " " + nrow["Resource"]);
            }
        }

        private void PopulateElectrodeMachineView()
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;
            string electrodeMachine1 = "Makino";
            string electrodeMachine2 = "Sodick";

            //DateTime dt = new DateTime();

            queryString = "SELECT * FROM Tasks WHERE (StartDate <= @currentDate AND FinishDate >= @currentDate) AND TaskName = @currentTask ORDER BY StartDate ASC, Priority ASC";
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            //dt = FinishingDateTimePicker.Value;
            adapter.SelectCommand.Parameters.AddWithValue("@currentDate", ElectrodesDateTimePicker.Value.ToString("d"));
            adapter.SelectCommand.Parameters.AddWithValue("@currentTask", "CNC Electrodes");
            adapter.Fill(datatable);

            //populateToolnumberComboBox(datatable);

            JobLabelElectrodeMachine1.Text = "Job: ";
            ComponentLabelElectrodeMachine1.Text = "Component: ";
            OperatorLabelElectrodeMachine1.Text = "Operator: ";

            JobLabelElectrodeMachine2.Text = "Job: ";
            ComponentLabelElectrodeMachine2.Text = "Component: ";
            OperatorLabelElectrodeMachine2.Text = "Operator: ";

            ListBoxElectrodeMachine1.Items.Clear();
            ListBoxElectrodeMachine2.Items.Clear();

            foreach (DataRow nrow in datatable.Rows)
            {
                if (nrow["Status"].ToString() != "In Progress" && nrow["Status"].ToString() != "On Hold")
                {
                    if (nrow["Resource"].ToString() == electrodeMachine1)
                    {
                        ListBoxElectrodeMachine1.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == electrodeMachine2)
                    {
                        ListBoxElectrodeMachine2.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }

                }
                else if (nrow["Status"].ToString() == "In Progress")
                {
                    if (nrow["Resource"].ToString() == electrodeMachine1)
                    {
                        JobLabelElectrodeMachine1.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelElectrodeMachine1.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelElectrodeMachine1.Text = "Operator: " + nrow["Operator"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == electrodeMachine2)
                    {
                        JobLabelElectrodeMachine2.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelElectrodeMachine2.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelElectrodeMachine2.Text = "Operator: " + nrow["Operator"].ToString();
                    }

                }

                //Console.WriteLine(nrow["JobNumber"] + " " + nrow["Component"] + " " + nrow["Resource"]);
            }
        }

        private void PopulateEDMMachineView()
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;
            string edmMachine1 = "Sodick 1";
            string edmMachine2 = "Sodick 2";

            //DateTime dt = new DateTime();

            queryString = "SELECT * FROM Tasks WHERE (StartDate <= @currentDate AND FinishDate >= @currentDate) AND TaskName = @currentTask ORDER BY StartDate ASC, Priority ASC";
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            //dt = FinishingDateTimePicker.Value;
            adapter.SelectCommand.Parameters.AddWithValue("@currentDate", EDMDateTimePicker.Value.ToString("d"));
            adapter.SelectCommand.Parameters.AddWithValue("@currentTask", "EDM Sinker");
            adapter.Fill(datatable);

            //populateToolnumberComboBox(datatable);

            JobLabelEDMMachine1.Text = "Job: ";
            ComponentLabelEDMMachine1.Text = "Component: ";
            OperatorLabelEDMMachine1.Text = "Operator: ";

            JobLabelEDMMachine2.Text = "Job: ";
            ComponentLabelEDMMachine2.Text = "Component: ";
            OperatorLabelEDMMachine2.Text = "Operator: ";

            ListBoxEDMMachine1.Items.Clear();
            ListBoxEDMMachine2.Items.Clear();

            foreach (DataRow nrow in datatable.Rows)
            {
                if (nrow["Status"].ToString() != "In Progress" && nrow["Status"].ToString() != "On Hold")
                {
                    if (nrow["Resource"].ToString() == edmMachine1)
                    {
                        ListBoxEDMMachine1.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                    else if (nrow["Resource"].ToString() == edmMachine2)
                    {
                        ListBoxEDMMachine2.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }

                }
                else if (nrow["Status"].ToString() == "In Progress")
                {
                    if (nrow["Resource"].ToString() == edmMachine1)
                    {
                        JobLabelEDMMachine1.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelEDMMachine1.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelEDMMachine1.Text = "Operator: " + nrow["Operator"].ToString();
                    }
                    else if (nrow["Resource"].ToString() == edmMachine2)
                    {
                        JobLabelEDMMachine2.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelEDMMachine2.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelEDMMachine2.Text = "Operator: " + nrow["Operator"].ToString();
                    }

                }

                //Console.WriteLine(nrow["JobNumber"] + " " + nrow["Component"] + " " + nrow["Resource"]);
            }
        }

        private void PopulateInspectionMachineView()
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;
            string inspectionMachine1 = "Brown & Sharpe";
            //string inspectionMachine2 = "None";

            //DateTime dt = new DateTime();

            queryString = "SELECT * FROM Tasks WHERE (StartDate <= @currentDate AND FinishDate >= @currentDate) AND TaskName CONTAINS '%@currentTask%' ORDER BY StartDate ASC, Priority ASC";
            //queryString2 = "SELECT * FROM Tasks WHERE (StartDate <= @currentDate AND FinishDate >= @currentDate) AND TaskName CONTAINS '%@currentTask%' AND (Status <> 'Completed') ORDER BY StartDate ASC, Priority ASC";
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            //dt = FinishingDateTimePicker.Value;
            adapter.SelectCommand.Parameters.AddWithValue("@currentDate", InspectionDateTimePicker.Value.ToString("d"));
            adapter.SelectCommand.Parameters.AddWithValue("@currentTask", "Inspection");
            adapter.Fill(datatable);

            //populateToolnumberComboBox(datatable);

            JobLabelInspectionMachine1.Text = "Job: ";
            ComponentLabelInspectionMachine1.Text = "Component: ";
            OperatorLabelInspectionMachine1.Text = "Operator: ";
            TaskNameLabelInspectionMachine1.Text = "Task Name: ";

            ListBoxInspectionMachine1.Items.Clear();

            foreach (DataRow nrow in datatable.Rows)
            {
                if (nrow["Status"].ToString() != "In Progress" && nrow["Status"].ToString() != "On Hold")
                {
                    if (nrow["Resource"].ToString() == inspectionMachine1)
                    {
                        ListBoxInspectionMachine1.Items.Add(nrow["JobNumber"] + " " + nrow["Component"]);
                    }
                }
                else if (nrow["Status"].ToString() == "In Progress")
                {
                    if (nrow["Resource"].ToString() == inspectionMachine1)
                    {
                        JobLabelInspectionMachine1.Text = "Job: " + nrow["JobNumber"].ToString();
                        ComponentLabelInspectionMachine1.Text = "Component: " + nrow["Component"].ToString();
                        OperatorLabelInspectionMachine1.Text = "Operator: " + nrow["Operator"].ToString();
                        TaskNameLabelInspectionMachine1.Text = "Task Name: " + nrow["TaskName"].ToString();
                    }
                }

                //Console.WriteLine(nrow["JobNumber"] + " " + nrow["Component"] + " " + nrow["Resource"]);
            }
        }

        private void PopulateSelectedTab()
        {
            if(ProcessTabs.SelectedTab.Text == "Design")
            {
                PopulateDesignerView();
            }
            else if (ProcessTabs.SelectedTab.Text == "Programming")
            {
                PopulateProgrammerView();
            }
            else if (ProcessTabs.SelectedTab.Text == "Roughing")
            {
                PopulateRoughMachineView();
            }
            else if (ProcessTabs.SelectedTab.Text == "Finishing")
            {
                PopulateFinishMachineView();
            }
            else if (ProcessTabs.SelectedTab.Text == "Electrodes")
            {
                PopulateElectrodeMachineView();
            }
            else if (ProcessTabs.SelectedTab.Text == "EDM")
            {
                PopulateEDMMachineView();
            }
            else if (ProcessTabs.SelectedTab.Text == "Inspection")
            {
                PopulateInspectionMachineView();
            }
        }

        private int PercentTextBoxValidation(TextBox textbox)
        {
            int someValue = 0;

            bool isANumber = int.TryParse(textbox.Text, out someValue);

            if (isANumber == true)
            {
                if (someValue > 100 || someValue < 0)
                {
                    MessageBox.Show("Please enter a value between 0 and 100.");
                    textbox.Text = "";
                }
            }
            else if(textbox.Text == "")
            {

            }
            else
            {
                MessageBox.Show("Please enter a number.");
                textbox.Text = "";
            }

            return someValue;
        }

        private DateTime GetForwardDateFromUser()
        {
            using (var form = new ForwardDateWindow())
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    return form.ForwardDate;
                }
                else if (result == DialogResult.Cancel)
                {

                }

                return new DateTime(2000, 1, 1);
            }
        }

        private DateTime GetBackDateFromUser()
        {
            Database db = new Database();
            var number = GetComboBoxInfo();
            ProjectInfo pi = db.GetProjectInfo(number.jobNumber, number.projectNumber);

            using (var form = new BackDateWindow(pi.DueDate))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    return form.BackDate;
                }
                else if (result == DialogResult.Cancel)
                {

                }

                return new DateTime(2000, 1, 1);
            }
        }

        private List<string> GetComponentListFromUser(string textString = "")
        {
            Database db = new Database();
            var number = GetComboBoxInfo();
            List<string> componentList = db.GetComponentList(number.jobNumber, number.projectNumber);

            using (var form = new SelectComponentsWindow(componentList, textString))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    return form.ComponentList;
                }
                else if (result == DialogResult.Cancel)
                {

                }

                return null;
            }
        }

        private void CreateProject()
        {
            using (var form = new Project_Creation_Form())
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    if(form.DataValidated)
                    {
                        PopulateJobNumberComboBox();
                        RefreshDataGridView();
                        JobNumberComboBox.SelectedItem = form.Project.JobNumber + " - #" + form.Project.ProjectNumber;
                    }
                }
                else if (result == DialogResult.Cancel)
                {

                }
            }
        }

        private void EditProject(ProjectInfo project)
        {
            using (var form = new Project_Creation_Form(project))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    if (form.DataValidated)
                    {
                        PopulateJobNumberComboBox();
                        RefreshDataGridView();
                        JobNumberComboBox.SelectedItem = form.Project.JobNumber + " - #" + form.Project.ProjectNumber;
                    }
                }
                else if (result == DialogResult.Cancel)
                {

                }
            }
        }

        private void FinishPercentTextBox1_TextChanged(object sender, EventArgs e)
        {
            FinishProgressBar1.Value = PercentTextBoxValidation((TextBox)sender);
        }

        private void FinishPercentTextBox2_TextChanged(object sender, EventArgs e)
        {
            FinishProgressBar2.Value = PercentTextBoxValidation((TextBox)sender);
        }

        private void FinishPercentTextBox3_TextChanged(object sender, EventArgs e)
        {
            FinishProgressBar3.Value = PercentTextBoxValidation((TextBox)sender);
        }

        private void FinishPercentTextBox4_TextChanged(object sender, EventArgs e)
        {
            FinishProgressBar4.Value = PercentTextBoxValidation((TextBox)sender);
        }

        private void FinishingDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            PopulateFinishMachineView();
        }

        private void ProcessTabs_Selected(object sender, TabControlEventArgs e)
        {
            SetQueryString();
            PopulateDataGridViewComboboxes();
            RefreshDataGridView();
            PopulateSelectedTab();
        }

        private void DesignerDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            PopulateDesignerView();
        }

        private void ProgrammerDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            PopulateProgrammerView();
        }

        private void RoughingDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            PopulateRoughMachineView();
        }

        private void ElectrodesDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            PopulateElectrodeMachineView();
        }

        private void EDMDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            PopulateEDMMachineView();
        }

        private void InspectionDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            PopulateInspectionMachineView();
        }

        private void LoadProjectButton_Click(object sender, EventArgs e)
        {
            Database db = new Database();
            //db.LoadMSProjectToDatabase();
            PopulateJobNumberComboBox();
            RefreshDataGridView();
        }

        private void JobNumberComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!FormLoading)
                if (JobNumberComboBox.Text != "All")
                {
                    var number = GetComboBoxInfo();

                    DataTable.DefaultView.RowFilter = "[JobNumber] = '" + number.jobNumber + "' AND [ProjectNumber] = '" + number.projectNumber + "'";
                    DataGridView1.DataSource = DataTable;
                }
                else
                {
                    DataTable.DefaultView.RowFilter = string.Empty;
                    DataGridView1.DataSource = DataTable;
                }
        }

        private void CreateProjectButton_Click(object sender, EventArgs e)
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

        private void BulkAssignButton_Click(object sender, EventArgs e)
        {
            Database db = new Database();
            string[] jobNumberComboTextArr;
            
            using (var barf = new BulkAssignResourcesForm())
            {
                var result = barf.ShowDialog();
                if (result == DialogResult.Cancel)
                {
                    barf.Close();
                }
                else if (result == DialogResult.OK)
                {
                    jobNumberComboTextArr = JobNumberComboBox.Text.Split(' ');
                    db.BulkAssignRoles(jobNumberComboTextArr[0], barf.RoughProgrammer, barf.FinishProgrammer, barf.ElectrodeProgrammer);
                    RefreshDataGridView();
                    barf.Close();
                }
            }    
        }

        private DateTime GetLatestPredecessorFinishDate(string jn, int pn, string component, string predecessors)
        {
            Database db = new Database();
            DateTime? latestFinishDate = null;
            DateTime currentDate;
            string[] predecessorArr;
            string predecessor;

            predecessorArr = predecessors.Split(',');

            foreach(string currPredecessor in predecessorArr)
            {
                predecessor = currPredecessor.Trim();
                currentDate = db.GetFinishDate(jn, pn, component, Convert.ToInt16(predecessor));

                if(latestFinishDate == null || latestFinishDate < currentDate)
                {
                    latestFinishDate = currentDate;
                }
            }

            return (DateTime)latestFinishDate;
        }

        private bool BlankStartFinishDateExists(ProjectInfo pi)
        {
            foreach (ClassLibrary.Component component in pi.ComponentList)
            {
                foreach (TaskInfo task in component.TaskList)
                {
                    if (task.StartDate == null || task.FinishDate == null)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private void oDateTimePicker_CloseUp(object sender, EventArgs e)
        {
            try
            {
                Database db = new Database();
                // Hiding the control after use  
                //oDateTimePicker.Visible = false;
                DataGridView1.Controls.Remove(oDateTimePicker);
                DateTime startDate;
                int projectNumber, taskID;
                string jobNumber, component, predecessors, duration;

                jobNumber = DataGridView1.CurrentRow.Cells["JobNumber"].Value.ToString();
                component = DataGridView1.CurrentRow.Cells["Component"].Value.ToString();
                duration = DataGridView1.CurrentRow.Cells["Duration"].Value.ToString();
                predecessors = DataGridView1.CurrentRow.Cells["Predecessors"].Value.ToString();
                projectNumber = Convert.ToInt32(DataGridView1.CurrentRow.Cells["ProjectNumber"].Value);
                taskID = Convert.ToInt32(DataGridView1.CurrentRow.Cells["TaskID"].Value);

                if (DataGridView1.CurrentCell.OwningColumn.Name == "StartDate")
                {
                    startDate = oDateTimePicker.Value;

                    if (predecessors != "" && startDate < GetLatestPredecessorFinishDate(jobNumber, projectNumber, component, predecessors))
                    {
                        MessageBox.Show("You cannot put a task start date before its predecessor's finish date.");
                        return;
                    }

                    db.ChangeTaskStartDate(jobNumber, projectNumber, component, oDateTimePicker.Value, duration, taskID);

                }
                else if (DataGridView1.CurrentCell.OwningColumn.Name == "FinishDate")
                {
                    db.ChangeTaskFinishDate(jobNumber, projectNumber, component, oDateTimePicker.Value, taskID);
                }

                RefreshDataGridView();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void dateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            // Some code could go here to check a date before accepting it.
            //if (predecessorsViolated() == true  && FinishingDataGridView.CurrentCellAddress.X == 5)
            //{
            //    MessageBox.Show("Cannot have a start date of a task before the finish date of its predecessor.");
            //    return;
            //}
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string columnName = DataGridView1.Columns[e.ColumnIndex].Name;
            // If any cell is clicked on the fifth or sixth column which is our date Column  
            if ((columnName == "StartDate" || columnName == "FinishDate") && e.RowIndex != -1)
            {
                //Initialized a new DateTimePicker Control  
                oDateTimePicker = new DateTimePicker();

                //Adding DateTimePicker control into DataGridView   
                DataGridView1.Controls.Add(oDateTimePicker);

                // Setting the format (i.e. 2014-10-10)  
                oDateTimePicker.Format = DateTimePickerFormat.Short;

                // It returns the retangular area that represents the Display area for a cell  
                Rectangle oRectangle = DataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);

                // If the selected column is start date and the task has a predecessor set the date to the finish date of the predecessor.
                if(columnName == "StartDate" && DataGridView1.Rows[e.RowIndex].Cells["Predecessors"].Value.ToString() != "")
                {
                    oDateTimePicker.Value = GetLatestPredecessorFinishDate(
                        DataGridView1.Rows[e.RowIndex].Cells["JobNumber"].Value.ToString(),
                        Convert.ToInt32(DataGridView1.Rows[e.RowIndex].Cells["ProjectNumber"].Value),
                        DataGridView1.Rows[e.RowIndex].Cells["Component"].Value.ToString(),
                        DataGridView1.Rows[e.RowIndex].Cells["Predecessors"].Value.ToString());
                }
                // If the select column is a start date or a finish date and there is a date in the field select the date.
                else if((DataGridView1.Rows[e.RowIndex]).Cells[e.ColumnIndex].Value.ToString() != "")
                {
                    oDateTimePicker.Value = Convert.ToDateTime((DataGridView1.Rows[e.RowIndex]).Cells[e.ColumnIndex].Value);
                }

                //Setting area for DateTimePicker Control  
                oDateTimePicker.Size = new Size(oRectangle.Width, oRectangle.Height);

                // Setting Location  
                oDateTimePicker.Location = new Point(oRectangle.X, oRectangle.Y);

                // An event attached to dateTimePicker Control which is fired when DateTimeControl is closed  
                oDateTimePicker.CloseUp += new EventHandler(oDateTimePicker_CloseUp);

                // An event attached to dateTimePicker Control which is fired when any date is selected  
                oDateTimePicker.TextChanged += new EventHandler(dateTimePicker_OnTextChange);

                // Now make it visible  
                oDateTimePicker.Visible = true;

                // TODO: Make time picker move when the user scrolls through the datagridview.
            }

            // If any cell in the project column is clicked, open up the corresponding Kan Ban Workbook.
            if(columnName == "ProjectNumber")
            {
                Database db = new Database();

                // This line failed when I changed a project and had the Kan Ban for 180079 open.
                db.OpenKanBanWorkbook(db.GetKanBanWorkbookPath(DataGridView1.Rows[e.RowIndex].Cells["JobNumber"].Value.ToString(), Convert.ToInt32(DataGridView1.Rows[e.RowIndex].Cells["ProjectNumber"].Value)));
                //MessageBox.Show(db.getKanBanWorkbookPath(DataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString(), Convert.ToInt32(DataGridView1.Rows[e.RowIndex].Cells[16].Value)));
            }
        }

        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (FormLoading == false)
                {
                    Database db = new Database();

                    if (DataGridView1.Columns[e.ColumnIndex].Name == "FinishDate")
                    {
                        string jobNumber, component;
                        int projectNumber, currentTaskID;
                        DateTime currentTaskFinishDate;

                        jobNumber = DataGridView1.CurrentRow.Cells["JobNumber"].Value.ToString();
                        projectNumber = Convert.ToInt32(DataGridView1.CurrentRow.Cells["ProjectNumber"].Value);
                        component = DataGridView1.CurrentRow.Cells["Component"].Value.ToString();

                        currentTaskFinishDate = Convert.ToDateTime(DataGridView1.CurrentRow.Cells["FinishDate"].Value);
                        currentTaskID = Convert.ToInt16(DataGridView1.CurrentRow.Cells["TaskID"].Value);

                        db.MoveDescendents(jobNumber, projectNumber, component, currentTaskFinishDate, currentTaskID);
                    }

                    db.UpdateDatabase(sender, e);
                    RefreshDataGridView();
                    PopulateSelectedTab();
                    //populateDesignerView();
                    //populateProgrammerView();
                    //populateRoughMachineView();
                    //populateFinishMachineView();
                    //populateElectrodeMachineView();
                    //populateEDMMachineView();
                    //populateInspectionMachineView();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void DataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs anError)
        {
            //MessageBox.Show("Error happened " + anError.Context.ToString());

            if (anError.Context == DataGridViewDataErrorContexts.Commit)
            {
                MessageBox.Show("Commit error");
            }
            if (anError.Context == DataGridViewDataErrorContexts.CurrentCellChange)
            {
                MessageBox.Show("Cell change");
            }
            if (anError.Context == DataGridViewDataErrorContexts.Parsing)
            {
                MessageBox.Show("Parsing Error");
            }
            if (anError.Context == DataGridViewDataErrorContexts.LeaveControl)
            {
                MessageBox.Show("Leave control error");
            }

            if ((anError.Exception) is ConstraintException)
            {
                DataGridView view = (DataGridView)sender;
                view.Rows[anError.RowIndex].ErrorText = "an error";
                view.Rows[anError.RowIndex].Cells[anError.ColumnIndex].ErrorText = "an error";

                anError.ThrowException = false;
            }
        }

        private void DataGridView1_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
            //    try
            //    {
            //        DataTable changes = ((DataTable)FinishingDataGridView.DataSource).GetChanges();
            //        if (changes != null)
            //        {
            //            SqlCommandBuilder mcb = new MySqlCommandBuilder(mySqlDataAdapter);
            //            mySqlDataAdapter.UpdateCommand = mcb.GetUpdateCommand();
            //            mySqlDataAdapter.Update(changes);
            //            ((DataTable)FinishingDataGridView.DataSource).AcceptChanges();

            //            MessageBox.Show("Cell Updated");
            //            return;
            //        }


            //    }

            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message);
            //    }
        }

        private void CreateKanBanButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (JobNumberComboBox.Text == "All")
                {
                    MessageBox.Show("Just select a single project.");
                    return;
                }
                else
                {
                    Database db = new Database();
                    ExcelInteractions ei = new ExcelInteractions();
                    var number = GetComboBoxInfo();
                    string path;
                    List<string> componentList = new List<string>();

                    ProjectInfo pi = db.GetProject(number.jobNumber, number.projectNumber);

                    if (BlankStartFinishDateExists(pi))
                    {
                        MessageBox.Show("A blank start or finish date exists. Please fill in all dates.");
                        return;
                    }

                    if (KanBanExists(number.jobNumber, number.projectNumber))
                    {
                        DialogResult result = MessageBox.Show("A Kan Ban for this project already exists.\n\nDo you want to create a new one?\n\n" +
                            "(Click 'Yes' to create new one.  Click 'No' to add to or edit the existing one.)", "Warning",
                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                        if (result == DialogResult.Yes)
                        {
                            path = ei.GenerateKanBanWorkbook(pi);

                            if (path != "")
                            {
                                db.SetKanBanWorkbookPath(path, pi.JobNumber, pi.ProjectNumber);
                            }
                        }
                        else if (result == DialogResult.No)
                        {
                            componentList = GetComponentListFromUser("Kan Ban");

                            if (componentList == null)
                            {
                                return;
                            }

                            ei.EditKanBanWorkbook(pi, db.GetKanBanWorkbookPath(number.jobNumber, number.projectNumber), componentList);
                        }
                        else if (result == DialogResult.Cancel)
                        {
                            return;
                        }

                    }
                    else
                    {
                        path = ei.GenerateKanBanWorkbook(pi);

                        if (path != "")
                        {
                            db.SetKanBanWorkbookPath(path, pi.JobNumber, pi.ProjectNumber);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

		private void ManageResourcesButton_Click(object sender, EventArgs e)
		{
			ManageResourcesForm MRF = new ManageResourcesForm();
			MRF.Show();
		}

        private void removeProjectButton_Click(object sender, EventArgs e)
        {
            RemoveProject();
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            if (JobNumberComboBox.Text == "All")
            {
                MessageBox.Show("Just select one job.");
            }
            else
            {
                Database db = new Database();
                var number = GetComboBoxInfo();
                //MessageBox.Show(db.getHighestProjectTaskID(number.jobNumber, number.projectNumber).ToString());
            }
        }

        private void ForwardDateButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (JobNumberComboBox.Text == "All")
                {
                    MessageBox.Show("Please select a project.");
                    return;
                }

                Database db = new Database();
                var number = GetComboBoxInfo();
                List<string> componentList = null;


                DialogResult dialogResult = MessageBox.Show("Do you want to forward schedule all tasks? (Click \"No\" to selectively schedule component tasks.", "Forward Schedule All?",
                                                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (dialogResult == DialogResult.Yes)
                {

                }
                else if (dialogResult == DialogResult.No)
                {
                    componentList = GetComponentListFromUser("Forward Date");

                    if (componentList == null)
                    {
                        return;
                    }
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    return;
                }

                db.ForwardDateProjectTasks(number.jobNumber, number.projectNumber, componentList, GetForwardDateFromUser());
                RefreshDataGridView();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void BackDateButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (JobNumberComboBox.Text == "All")
                {
                    MessageBox.Show("Please select a project.");
                    return;
                }

                Database db = new Database();
                var number = GetComboBoxInfo();
                List<string> componentList = null;

                DialogResult dialogResult = MessageBox.Show("Do you want to back schedule all tasks? (Click \"No\" to selectively schedule component tasks.", "Back Schedule All?",
                                          MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (dialogResult == DialogResult.Yes)
                {

                }
                else if (dialogResult == DialogResult.No)
                {
                    componentList = GetComponentListFromUser("Back Date");

                    if (componentList == null)
                    {
                        return;
                    }
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    return;
                }

                db.BackDateProjectTasks(number.jobNumber, number.projectNumber, componentList, GetBackDateFromUser());
                RefreshDataGridView();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void EditProjectButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (JobNumberComboBox.Text == "All")
                {
                    MessageBox.Show("Please select a project to edit.");
                }
                else
                {
                    Database db = new Database();
                    var number = GetComboBoxInfo();
                    ProjectInfo project = db.GetProject(number.jobNumber, number.projectNumber);
                    EditProject(project);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void JobNumberComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            
        }

        private void DataGridView1_MouseEnter(object sender, EventArgs e)
        {
            DataGridView1.Focus();
        }

        private void JobNumberComboBox_MouseEnter(object sender, EventArgs e)
        {
            JobNumberComboBox.Focus();
        }

        private void OpenViewerButton_Click(object sender, EventArgs e)
        {
            string filepath = @"X:\TOOLROOM\Workload Tracking System\Debug 2\Toolroom Project Viewer.exe";
            FileInfo fi = new FileInfo(filepath);

            if (fi.Exists)
            {
                //var attributes = File.GetAttributes(fi.FullName);    

                var res = Process.Start(fi.FullName);
            }
            else
            {
                MessageBox.Show("Can't find Toolroom Project Viewer Executable File. " + filepath + ".");
            }
        }

        private void DataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if(DataGridView1.Controls.Contains(oDateTimePicker))
                DataGridView1.Controls.Remove(oDateTimePicker);
        }
    }
}
