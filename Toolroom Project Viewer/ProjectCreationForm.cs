﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using ClassLibrary;
using DevExpress.XtraScheduler;
using DevExpress.XtraEditors;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using NuGet;
using ClassLibrary.Models;

namespace Toolroom_Project_Viewer
{
    /// <summary>
    /// Class for the ProjectCreationForm.
    /// </summary> 
    public partial class ProjectCreationForm : DevExpress.XtraEditors.XtraForm
    {
        Excel.Application excelApp;
        string prefix;
        bool formLoad = false;
        bool missingTaskInfo = false;
        //bool quoteLoaded = false;

        private DataTable RoleTable { get; set; }
        private DataTable DeptRoleTable { get; set; }
        public ProjectModel Project { get; private set; }
        private ComponentModel SelectedComponent { get; set; }
        private TaskModel SelectedTask { get; set; }
        public SchedulerDataStorage SchedulerStorageProp { get; private set; }
        private ContextMenuStrip ComponentMenu, ProjectMenu, TaskMenu;
        public bool DataValidated { get; private set; }
        public List<UserModel> UserList { get; set; }

        public ProjectCreationForm()
        {

        }
        /// <summary>
        /// Initializes a new Project Form.
        /// </summary> 
        public ProjectCreationForm(SchedulerDataStorage schedulerStorage)
        {
            Console.WriteLine("ProjectCreationForm Constructor");

            formLoad = true;
            Project = new ProjectModel();
            SchedulerStorageProp = schedulerStorage;
            RoleTable = Database.GetRoleTable();
            DeptRoleTable = Database.GetDepartmentRoles();
            UserList = Database.GetUsers();

            InitializeComponent();

            if (GetDPI() == 120)
            {
                this.Width = 1330;
            }
            //MessageBox.Show($"{getScalingFactor()}");
        }
        /// <summary>
        /// Initializes a new Project Form and sets the project property of the form to an instance of a property.
        /// </summary> 
        public ProjectCreationForm(ProjectModel project, SchedulerDataStorage schedulerStorage)
        {
            Console.WriteLine("ProjectCreationForm Constructor");

            formLoad = true;
            this.Project = project;
            SchedulerStorageProp = schedulerStorage;
            RoleTable = Database.GetRoleTable();
            DeptRoleTable = Database.GetDepartmentRoles();
            UserList = Database.GetUsers();
            InitializeComponent();
        }
        private void ProjectCreationForm_Load(object sender, EventArgs e)
        {
            if (Project.HasProjectInfo)
            {
                this.Text = "Edit Project";
                LoadProjectToForm(Project);
                this.CreateProjectButton.Text = "Change";
            }

            ProjectMenu = new ContextMenuStrip();

            ToolStripMenuItem renameLabel1 = new ToolStripMenuItem();

            renameLabel1.Text = "Rename";

            ToolStripMenuItem componentLabel = new ToolStripMenuItem();

            componentLabel.Text = "Load Component";

            ProjectMenu.Items.AddRange(new ToolStripMenuItem[] { renameLabel1, componentLabel });
            ProjectMenu.Click += ProjectMenuStrip_Click;
            MoldBuildTreeView.Nodes[0].ContextMenuStrip = ProjectMenu;

            ComponentMenu = new ContextMenuStrip();
            // Create component menu items.
            ToolStripMenuItem renameLabel2 = new ToolStripMenuItem();

            renameLabel2.Text = "Rename";

            ToolStripMenuItem copyLabel = new ToolStripMenuItem();

            copyLabel.Text = "Copy";

            ToolStripMenuItem createTemplateLabel = new ToolStripMenuItem();

            createTemplateLabel.Text = "Create Template";

            ToolStripMenuItem loadTemplateLabel = new ToolStripMenuItem();

            loadTemplateLabel.Text = "Load Template";

            ComponentMenu.Items.AddRange(new ToolStripMenuItem[] { renameLabel2, copyLabel, createTemplateLabel, loadTemplateLabel });
            ComponentMenu.Click += ComponentMenuStrip_Click;

            TaskMenu = new ContextMenuStrip();

            ToolStripMenuItem renameLabel3 = new ToolStripMenuItem();

            renameLabel3.Text = "Rename";

            TaskMenu.Items.AddRange(new ToolStripMenuItem[] { renameLabel3 });
            TaskMenu.Click += TaskMenuStrip_Click;

            prefix = "A-";
        }

        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117,
            /// <summary>
            /// Logical pixels inch in X
            /// </summary>
            LOGPIXELSX = 88,
            /// <summary>
            /// Logical pixels inch in Y
            /// </summary>
            LOGPIXELSY = 90
            // http://pinvoke.net/default.aspx/gdi32/GetDeviceCaps.html
        }

        private float GetDPI()
        {
            Graphics g = Graphics.FromHwnd(IntPtr.Zero);
            IntPtr desktop = g.GetHdc();
            int LogicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.VERTRES);
            int PhysicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.DESKTOPVERTRES);

            int Xdpi = GetDeviceCaps(desktop, (int)DeviceCap.LOGPIXELSX);
            int Ydpi = GetDeviceCaps(desktop, (int)DeviceCap.LOGPIXELSY);

            //float ScreenScalingFactor = (float)PhysicalScreenHeight / (float)LogicalScreenHeight;

            //MessageBox.Show($"{Xdpi} {Ydpi} {PhysicalScreenHeight} {LogicalScreenHeight}");

            return Xdpi; // 1.25 = 125%
        }

        private void PopulateComboBox(System.Windows.Forms.ComboBox cb)
        {
            if (cb.Name == "ToolMakerComboBox")
            {
                cb.DataSource = GetResourceList("Tool Maker");
            }
            else if (cb.Name == "DesignerComboBox")
            {
                cb.DataSource = GetResourceList("Designer");
            }
            else if (cb.Name == "RoughProgrammerComboBox")
            {
                cb.DataSource = GetResourceList("Rough Programmer");
            }
            else if (cb.Name == "FinishProgrammerComboBox")
            {
                cb.DataSource = GetResourceList("Finish Programmer");
            }
            else if (cb.Name == "ElectrodeProgrammerComboBox")
            {
                cb.DataSource = GetResourceList("Electrode Programmer");
            }
            else if (cb.Name == "EDMSinkerOperatorComboBox")
            {
                cb.DataSource = GetResourceList("EDM Sinker Operator");
            }
            else if (cb.Name == "RoughCNCOperatorComboBox")
            {
                cb.DataSource = GetResourceList("Rough CNC Operator");
            }
            else if (cb.Name == "ElectrodeCNCOperatorComboBox")
            {
                cb.DataSource = GetResourceList("Electrode CNC Operator");
            }
            else if (cb.Name == "FinishCNCOperatorComboBox")
            {
                cb.DataSource = GetResourceList("Finish CNC Operator");
            }
            else if (cb.Name == "EDMWireOperatorComboBox")
            {
                cb.DataSource = GetResourceList("EDM Wire Operator");
            }
        }
        public string FindMatchingDepartment(string comboBoxName)
        {
            List<string> searchWords = new List<string>();

            foreach (var item in DeptRoleTable.AsEnumerable())
            {
                searchWords = item.Field<string>("Role").Split(' ').ToList();

                Console.WriteLine($"Word Count: {searchWords.Count}");

                if (searchWords.All(x => comboBoxName.Contains(x)))
                {
                    return item.Field<string>("Department");
                }
            }

            return $"";
        }
        public void SetPersonnnel(object sender)
        {
            var combo = sender as System.Windows.Forms.ComboBox;            

            List<string> searchWords = new List<string>();

            string taskName = "";

            taskName = GeneralOperations.FindMatchingDepartment(combo.Name, DeptRoleTable);

            Project[combo.Name.Replace("ComboBox", "")] = combo.Text;

            foreach (var component in Project.Components)
            {
                List<TaskModel> result = component.Tasks.FindAll(x => x.TaskName == taskName);

                foreach (TaskModel task in result)
                {
                    if (task.Personnel != null && task.Personnel.Length > 0)
                    {
                        DialogResult result2 = MessageBox.Show($"{task.TaskName} for {component.Component} already has personnel assigned to it. Do wish to overwrite?", "Overwrite?", MessageBoxButtons.YesNo);

                        if (result2 == DialogResult.Yes)
                        {
                            task.Personnel = combo.Text;
                            task.Resources = GeneralOperations.GenerateResourceIDsString(SchedulerStorageProp, task.Machine, task.Personnel);
                        }
                    }
                    else
                    {
                        task.Personnel = combo.Text;
                        task.Resources = GeneralOperations.GenerateResourceIDsString(SchedulerStorageProp, task.Machine, task.Personnel);
                    }
                }
            }
        }
        private List<string> GetResourceList(string role)
        {
            List<string> resourceList = new List<string>();

            var result = from roleTable in RoleTable.AsEnumerable()
                         where roleTable.Field<string>("Role") == role
                         select roleTable;

            resourceList.Add("");

            foreach (var resource in result)
            {
                resourceList.Add(resource.Field<string>("ResourceName"));
            }

            return resourceList;
        }
        private void RenameNode()
        {
            string input = Interaction.InputBox("Enter a new name:", "Change Name", MoldBuildTreeView.SelectedNode.Text, -1, -1);

            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;

            if (selectedNode.Level >= 0 && selectedNode.Level <= 2)
            {
                RenameNode(input);
            }

            if (selectedNode.Level == 2)
            {
                predecessorsListBox.SelectedIndexChanged -= new System.EventHandler(predecessorsListBox_SelectedIndexChanged);

                predecessorsListBox.DataSource = GetPredecessorList(selectedNode.Parent);

                predecessorsListBox.ClearSelected();

                SelectPredecessors(selectedNode);

                predecessorsListBox.SelectedIndexChanged += new System.EventHandler(predecessorsListBox_SelectedIndexChanged);
            }
        }
        private void RenameNode(string newName)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            bool isValidChange = false;
            if (selectedNode == null || newName == "") return;
            
            if (selectedNode.Level == 0)
            {
                if (selectedNode.BackColor == Color.Red)
                {
                    selectedNode.BackColor = Color.White;
                    selectedNode.ForeColor = Color.Black; 
                }

                Project.JobNumber = newName;

                isValidChange = true;
            }
            else if (selectedNode.Level == 1)
            {
                if (!Project.ComponentNameExists(newName))
                {
                    if (SelectedComponent.SetName(newName))
                    {
                        isValidChange = true;
                    }
                }
            }
            else if (selectedNode.Level == 2)
            {
                ComponentModel component = Project.Components.ElementAt(selectedNode.Parent.Index);
                TaskModel task = component.Tasks.ElementAt(selectedNode.Index);
                task.SetName(newName);
                isValidChange = true;
            }

            if (isValidChange)
            {
                MoldBuildTreeView.SelectedNode.Text = newName;
            }
        }

        private void AddComponentToTree(string newNodeName)
        {
            if (newNodeName == "" || !Project.AddComponent(newNodeName)) return;

            TreeNode newNode = new TreeNode(newNodeName);
            MoldBuildTreeView.Nodes[0].Nodes.Add(newNode);

            if (MoldBuildTreeView.Nodes[0].Nodes.Count == 1)
            {
                MoldBuildTreeView.Nodes[0].Expand();
            }

            MoldBuildTreeView.Focus();
            MoldBuildTreeView.SelectedNode = newNode;
        }
        private void AddCopiedComponentToTree(ComponentModel copiedComponent)
        {
            if (!Project.AddComponent(copiedComponent)) return;

            TreeNode newComponentNode = new TreeNode(copiedComponent.Component);
            TreeNode newTaskNode;

            foreach (var task in copiedComponent.Tasks)
            {
                newTaskNode = newComponentNode.Nodes.Add(task.TaskName);

                newTaskNode.Nodes.Add(task.Hours.ToString());
                newTaskNode.Nodes.Add(task.Duration);
                newTaskNode.Nodes.Add(task.Machine);
                newTaskNode.Nodes.Add(task.Personnel);
                newTaskNode.Nodes.Add(task.Predecessors);
                newTaskNode.Nodes.Add(task.Notes);
            }

            MoldBuildTreeView.Nodes[0].Nodes.Add(newComponentNode);

            if (MoldBuildTreeView.Nodes[0].Nodes.Count == 1)
            {
                MoldBuildTreeView.Nodes[0].Expand();
            }

            MoldBuildTreeView.Focus();
            MoldBuildTreeView.SelectedNode = newComponentNode;
        }
        private void AddSelectedTasksToSelectedComponent()
        {
            string taskName;
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            var item = TaskListBox.SelectedItem;

            if (selectedNode == null || item == null || selectedNode.Level != 1)
            {
                MessageBox.Show("Please select a component to add tasks to or select tasks to add to a component.");
                return;
            }

            foreach (int i in TaskListBox.SelectedIndices)
            {
                var component = Project.Components.Where(x => x.Component == selectedNode.Text).First();
                taskName = TaskListBox.Items[i].ToString();
                MoldBuildTreeView.SelectedNode.Nodes.Add(taskName);
                component.AddTask(taskName, component.Component, SchedulerStorageProp);
            }

        }
        private void ReplaceComponentInTree(ComponentModel component)
        {

        }
        private void SetTaskInfoForSelectedTask()
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            ComponentModel component;
            TaskModel task;
            //string predecessorString = getSelectedPredecessorText(predecessorsListBox); // Uncomment to use project.

            if (durationNumericUpDown.Value.ToString() == "")
            {
                MessageBox.Show("Duration cannot be blank.");
                return;
            }

            if (!int.TryParse(durationNumericUpDown.Value.ToString(), out int result))
            {
                MessageBox.Show("Please enter a whole number for duration.");
                return;
            }
            // Check if task is selected.
            if (selectedNode.Level != 2)
            {
                MessageBox.Show("Please select a task to add info to.");
                return;
            }

            string predecessorString = GetSelectedPredecessorIndices(predecessorsListBox); // countTasks(MoldBuildTreeView, selectedNode.Parent.Text)

            // Check if selected task is set as its own predecessor.
            foreach (int index in predecessorsListBox.SelectedIndices)
            {
                if (selectedNode.Index == index)
                {
                    //predecessorsListBox.SelectedIndices[index] = false;
                    MessageBox.Show("A task cannot be its own predecessor.");
                    return;
                }
            }

            // Check if no predecessors are selected.
            if (predecessorString.Length == 0 && !selectedNode.Text.Contains("Program") && selectedNode.Index != 0)
            {
                MessageBox.Show("Please select a predecessor or predecessors.");
                return;
            }

            component = Project.Components.ElementAt(selectedNode.Parent.Index);

            task = component.Tasks.ElementAt(selectedNode.Index);

            // Check if selected node contains nodes and if task info fields are empty.
            // If true remove all task info nodes from selected task.
            if (selectedNode.Nodes.Count != 0 && TaskInfoIsEmpty())
            {
                for (int i = selectedNode.Nodes.Count - 1; i >= 0; i--)
                {
                    selectedNode.Nodes[i].Remove();
                }

                task.HasInfo = false;
            }
            // Check if selected task node contains any task info nodes.
            // If true change existing task info nodes to reflect changes in field (if any).
            else if (selectedNode.Nodes.Count != 0 && !TaskInfoIsEmpty())
            {

                selectedNode.Nodes[0].Text = hoursNumericUpDown.Value.ToString() + " Hour(s)";

                selectedNode.Nodes[1].Text = durationNumericUpDown.Value.ToString() + " " + durationUnitsComboBox.Text;

                selectedNode.Nodes[2].Text = machineComboBox.Text;

                selectedNode.Nodes[3].Text = personnelComboBox.Text;

                selectedNode.Nodes[4].Text = predecessorString;

                selectedNode.Nodes[5].Text = taskNotesTextBox.Text;

                task.HasInfo = true;
            }
            // Check if selected task node contains any task info nodes.
            // If false add nodes with info from fields.
            else if (selectedNode.Nodes.Count == 0)
            {
                for (int i = 0; i <= 5; i++)
                {
                    selectedNode.Nodes.Add("");
                }

                selectedNode.Nodes[0].Text = hoursNumericUpDown.Value.ToString() + " Hour(s)";

                selectedNode.Nodes[1].Text = durationNumericUpDown.Value.ToString() + " " + durationUnitsComboBox.Text;

                selectedNode.Nodes[2].Text = machineComboBox.Text;

                selectedNode.Nodes[3].Text = personnelComboBox.Text;

                selectedNode.Nodes[4].Text = predecessorString;

                selectedNode.Nodes[5].Text = taskNotesTextBox.Text;

                task.HasInfo = true;
            }

            task.SetTaskInfo
            (
                hoursNumericUpDown.Value,
                durationNumericUpDown.Value.ToString() + " " + durationUnitsComboBox.Text,
                machineComboBox.Text,
                personnelComboBox.Text,
                predecessorString,
                taskNotesTextBox.Text,
                SchedulerStorageProp
            );

            SelectNextTask();
        }
        private void DeactivatePersonnelValueChangedEvent()
        {
            RoughProgrammerComboBox.SelectedValueChanged -= Personnel_ValueChanged;
            RoughCNCOperatorComboBox.SelectedValueChanged -= Personnel_ValueChanged;
            ElectrodeProgrammerComboBox.SelectedValueChanged -= Personnel_ValueChanged;
            ElectrodeCNCOperatorComboBox.SelectedValueChanged -= Personnel_ValueChanged;
            FinishProgrammerComboBox.SelectedValueChanged -= Personnel_ValueChanged;
            FinishCNCOperatorComboBox.SelectedValueChanged -= Personnel_ValueChanged;
            EDMSinkerOperatorComboBox.SelectedValueChanged -= Personnel_ValueChanged;
            EDMWireOperatorComboBox.SelectedValueChanged -= Personnel_ValueChanged;
        }
        private void ActivatePersonnelValueChangedEvent()
        {
            RoughProgrammerComboBox.SelectedValueChanged += Personnel_ValueChanged;
            RoughCNCOperatorComboBox.SelectedValueChanged += Personnel_ValueChanged;
            ElectrodeProgrammerComboBox.SelectedValueChanged += Personnel_ValueChanged;
            ElectrodeCNCOperatorComboBox.SelectedValueChanged += Personnel_ValueChanged;
            FinishProgrammerComboBox.SelectedValueChanged += Personnel_ValueChanged;
            FinishCNCOperatorComboBox.SelectedValueChanged += Personnel_ValueChanged;
            EDMSinkerOperatorComboBox.SelectedValueChanged += Personnel_ValueChanged;
            EDMWireOperatorComboBox.SelectedValueChanged += Personnel_ValueChanged;
        }
        private void LoadProjectInfoToForm(ProjectModel project)
        {
            if (project.HasProjectInfo)
            {
                DeactivatePersonnelValueChangedEvent();

                MoldBuildTreeView.Nodes[0].Text = project.JobNumber;
                overLapAllowedCheckEdit.Checked = project.OverlapAllowed;
                ProjectNumberTextBox.Text = project.ProjectNumber.ToString();
                dueDateTimePicker.Value = project.DueDate;
                ToolMakerComboBox.Text = project.ToolMaker;
                DesignerComboBox.Text = project.Designer;
                RoughProgrammerComboBox.Text = project.RoughProgrammer;
                ElectrodeProgrammerComboBox.Text = project.ElectrodeProgrammer;
                FinishProgrammerComboBox.Text = project.FinishProgrammer;
                EDMSinkerOperatorComboBox.Text = project.EDMSinkerOperator;
                RoughCNCOperatorComboBox.Text = project.RoughCNCOperator;
                ElectrodeCNCOperatorComboBox.Text = project.ElectrodeCNCOperator;
                FinishCNCOperatorComboBox.Text = project.FinishCNCOperator;
                EDMWireOperatorComboBox.Text = project.EDMWireOperator;

                ActivatePersonnelValueChangedEvent();
            }
        }
        private void LoadProjectToForm(ProjectModel project)
        {
            PrintObjectTree();

            LoadProjectInfoToForm(project);

            LoadComponentListToTreeView(project.Components);

            MoldBuildTreeView.SelectedNode = MoldBuildTreeView.Nodes[0];
            this.ActiveControl = MoldBuildTreeView;
            MoldBuildTreeView.Nodes[0].Expand();
        }

        private void LoadComponentListToTreeView(List<ComponentModel> components)
        {
            TreeNode currentComponentNode, currentTaskNode;

            foreach (ComponentModel component in components)
            {
                currentComponentNode = MoldBuildTreeView.Nodes[0].Nodes.Add(component.Component);

                foreach (TaskModel task in component.Tasks)
                {
                    currentTaskNode = currentComponentNode.Nodes.Add(task.TaskName);

                    if (task.HasInfo == true)
                    {
                        currentTaskNode.Nodes.Add(task.Hours + " Hour(s)");
                        currentTaskNode.Nodes.Add(task.Duration);
                        currentTaskNode.Nodes.Add(task.Machine);
                        currentTaskNode.Nodes.Add(task.Personnel);
                        currentTaskNode.Nodes.Add(task.Predecessors);
                        currentTaskNode.Nodes.Add(task.Notes);
                    }
                }
            }
        }

        private void AddTaskListToTreeView(TreeNode selectedComponentNode, List<TaskModel> tasks)
        {
            TreeNode currentTaskNode;

            foreach (TaskModel task in tasks)
            {
                currentTaskNode = selectedComponentNode.Nodes.Add(task.TaskName);

                if (task.HasInfo == true)
                {
                    currentTaskNode.Nodes.Add(task.Hours + " Hour(s)");
                    currentTaskNode.Nodes.Add(task.Duration);
                    currentTaskNode.Nodes.Add(task.Machine);
                    currentTaskNode.Nodes.Add(task.Personnel);
                    currentTaskNode.Nodes.Add(task.Predecessors);
                    currentTaskNode.Nodes.Add(task.Notes);
                }
            }
        }

        private ProjectModel ConvertQuoteToProject(ProjectModel project)
        {
            // Need to check if form already contains project data.
            if (project.Components.Count > 0)
            {
                MessageBox.Show("Can't add a quote to a work project tree with data in it.");
                return project;
            }

            project.JobNumber = project.QuoteInfo.Customer + "_" + project.QuoteInfo.PartName + "-Quote"; // What to do when these two pieces of information are missing?
            project.SetProjectDueDate(DateTime.Today);
            project.HasProjectInfo = true;
            project.AddComponent("Mold");

            // Task list is automatically generated inside the QuoteInfo class when quote is read.
            project.Components.First().AddTaskList(project.QuoteInfo.TaskList);

            project.Components.First().Tasks.ForEach(x => x.Personnel = GetTaskPersonnel(x.TaskName));

            return project;
        }

        private bool TaskInfoIsEmpty()
        {
            if (
                GetValue(hoursNumericUpDown.Value.ToString()) == 0 &&
                GetValue(durationNumericUpDown.Value.ToString()) == 0 &&
                machineComboBox.Text == "" &&
                personnelComboBox.Text == "" &&
                predecessorsListBox.Text == "" &&
                taskNotesTextBox.Text == "")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private int GetValue(string value)
        {
            if (value == "")
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(value);
            }
        }

        private decimal GetDecimal(string value)
        {
            if (value == "")
            {
                return 0;
            }
            else
            {
                return Convert.ToDecimal(value);
            }
        }

        private void UpdateDuration()
        {
            decimal duration;
            decimal days;

            if (matchHoursCheckBox.Checked == true)
            {
                days = GetDecimal(hoursNumericUpDown.Value.ToString()) / 8;
                duration = Convert.ToDecimal(Math.Round(days, 0));
                durationNumericUpDown.Value = duration;
                durationUnitsComboBox.SelectedIndex = 1;
            }
        }

        private List<string> GetMachineList(string taskName)
        {
            List<string> machineList;

            if (taskName == "Program Rough")
            {
                machineList = new List<string> { "" };
            }
            else if (taskName == "Program Finish")
            {
                machineList = new List<string> { "" };
            }
            else if (taskName == "Program Electrodes")
            {
                machineList = new List<string> { "" };
            }
            else if (taskName == "CNC Rough")
            {
                machineList = GetResourceList("Rough Mill");
            }
            else if (taskName == "CNC Finish")
            {
                machineList = GetResourceList("Finish Mill");
            }
            else if (taskName == "CNC Electrodes")
            {
                machineList = GetResourceList("Graphite Mill");
            }
            else if (taskName == "EDM Sinker")
            {
                machineList = GetResourceList("EDM Sinker");
            }
            else if (taskName == "EDM Wire (In-House)")
            {
                machineList = GetResourceList("EDM Wire");
            }
            else
            {
                machineList = new List<string> { "" };
            }

            return machineList;
        }
        private List<string> GetPersonnelList(string taskName)
        {
            List<string> personnelList;

            if (taskName == "Program Rough")
            {
                personnelList = GetResourceList("Rough Programmer");
            }
            else if (taskName == "Program Finish")
            {
                personnelList = GetResourceList("Finish Programmer");
            }
            else if (taskName == "Program Electrodes")
            {
                personnelList = GetResourceList("Electrode Programmer");
            }
            else if (taskName == "CNC Rough")
            {
                personnelList = GetResourceList("Rough CNC Operator");
            }
            else if (taskName == "CNC Finish")
            {
                personnelList = GetResourceList("Finish CNC Operator");
            }
            else if (taskName.Contains("Grind"))  // I'm commenting this because the term tool maker has mingled with the term lead and there are a number of 'tool makers' that aren't capable of grinding.
            {
                personnelList = GetResourceList("Tool Maker");
            }
            else if (taskName == "CNC Electrodes")
            {
                personnelList = GetResourceList("Electrode CNC Operator");
            }
            else if (taskName == "EDM Sinker")
            {
                personnelList = GetResourceList("EDM Sinker Operator");
            }
            else if (taskName == "EDM Wire (In-House)")
            {
                personnelList = GetResourceList("EDM Wire Operator");
            }
            else if(taskName == "Design" || taskName == "Design Change")
            {
                personnelList = GetResourceList("Designer");
            }
            else if (taskName == "Hole Pop")
            {
                personnelList = GetResourceList("Hole Popper Operator");
            }
            else if (taskName.Contains("Inspection"))
            {
                personnelList = GetResourceList("CMM Operator");
            }
            else if (taskName == "Mold Service")
            {
                personnelList = GetResourceList("Tool Maker");
            }
            else
            {
                personnelList = new List<string> { "" };
            }

            return personnelList;
        }

        private TaskModel GetDefaultTaskInfo(string taskName)
        {
            TaskModel ti;

            if (taskName == "Program Rough")
            {
                ti = new TaskModel(RoughProgrammerComboBox.Text, "0", "1"); 
            }
            else if (taskName == "Program Finish")
            {
                ti = new TaskModel(FinishProgrammerComboBox.Text, "0", "1");
            }
            else if (taskName == "Program Electrodes")
            {
                ti = new TaskModel(ElectrodeProgrammerComboBox.Text, "0", "1");
            }
            else if (taskName == "CNC Rough")
            {
                ti = new TaskModel(RoughCNCOperatorComboBox.Text, "0", "1");
            }
            else if (taskName == "CNC Finish")
            {
                ti = new TaskModel(FinishCNCOperatorComboBox.Text, "0", "1");
            }
            else if (taskName == "CNC Electrodes")
            {
                ti = new TaskModel(ElectrodeCNCOperatorComboBox.Text, "0", "1");
            }
            else if (taskName == "Design")
            {
                ti = new TaskModel(DesignerComboBox.Text, "0", "0");
            }
            else if (taskName == "Design Change")
            {
                ti = new TaskModel(DesignerComboBox.Text, "0", "0");
            }
            else if (taskName == "Heat Treat")
            {
                ti = new TaskModel("", "0", "5");
            }
            else if (taskName.Contains("Inspection"))
            {
                ti = new TaskModel("", "1", "0");
            }
            else if (taskName.Contains("Grind"))
            {
                ti = new TaskModel(ToolMakerComboBox.Text, "1", "0");
            }
            else if (taskName.Contains("EDM Sinker"))
            {
                ti = new TaskModel(EDMSinkerOperatorComboBox.Text, "0", "1");
            }
            else if (taskName.Contains("EDM Wire"))
            {
                ti = new TaskModel(EDMWireOperatorComboBox.Text, "0", "1");
            }
            else if (taskName == "Polish")
            {
                ti = new TaskModel("", "0", "8");
            }
            else if (taskName == "Texturing")
            {
                ti = new TaskModel("", "0", "5");
            }
            else
            {
                ti = new TaskModel("", "0", "0"); ;
            }

            return ti;
        }

        private string GetTaskPersonnel(string taskName)
        {
            string personnel = "";

            if (taskName == "Design")
            {
                personnel = DesignerComboBox.Text;
            }
            else if (taskName == "Program Rough")
            {
                personnel = RoughProgrammerComboBox.Text;
            }
            else if (taskName == "Program Finish")
            {
                personnel = FinishProgrammerComboBox.Text;
            }
            else if (taskName == "Program Electrodes")
            {
                personnel = ElectrodeProgrammerComboBox.Text;
            }
            else if (taskName == "CNC Rough")
            {
                personnel = RoughCNCOperatorComboBox.Text;
            }
            else if (taskName == "CNC Finish")
            {
                personnel = FinishCNCOperatorComboBox.Text;
            }
            else if (taskName == "CNC Electrodes")
            {
                personnel = ElectrodeCNCOperatorComboBox.Text;
            }
            else if (taskName == "Heat Treat")
            {
                // No personnel are assigned to Heat Treat.
            }
            else if (taskName.Contains("Inspection"))
            {
                // There's no global input for inspection.
            }
            else if (taskName.Contains("Grind"))
            {
                personnel = ToolMakerComboBox.Text;
            }
            else if (taskName.Contains("EDM Sinker"))
            {
                personnel = EDMSinkerOperatorComboBox.Text;
            }
            else if (taskName.Contains("EDM Wire"))
            {
                personnel = EDMWireOperatorComboBox.Text;
            }
            else if (taskName == "Polish")
            {
                // There is no global input for Polish.
            }
            else if (taskName == "Texturing")
            {
                // There is no global input for Texture.
            }

            return personnel;
        }

        private List<string> GetPredecessorList(TreeNode node)
        {
            List<string> predecessorList = new List<string>();

            for (int i = 0; i < node.Nodes.Count; i++)
            {
                predecessorList.Add(node.Nodes[i].Text);
            }

            return predecessorList;
        }

        private List<string> GetPredecessorList(ComponentModel component)
        {
            List<string> predecessorList = new List<string>();

            predecessorList = component.Tasks.Select(x => x.TaskName).ToList();

            return predecessorList;
        }

        private string GetSelectedPredecessorText(ListBox listBox)
        {
            StringBuilder predecessorString = new StringBuilder();

            foreach (string item in listBox.SelectedItems)
            {
                if (predecessorString.Length == 0)
                {
                    predecessorString.Append(item);
                }
                else
                {
                    predecessorString.Append("," + item);
                }

            }

            return predecessorString.ToString();
        }

        private string GetSelectedPredecessorIndices(ListBox listBox)
        {
            StringBuilder predecessorString = new StringBuilder();
            int index;

            foreach (int n in listBox.SelectedIndices)
            {
                if (predecessorString.Length == 0)
                {
                    index = n + 1;
                    predecessorString.Append(index);
                }
                else
                {
                    index = n + 1;
                    predecessorString.Append("," + index);
                }

            }

            return predecessorString.ToString();
        }

        private string GetSelectedPredecessorIndexText(ListBox listBox)
        {
            StringBuilder predecessorIndexString = new StringBuilder();

            foreach (string index in listBox.SelectedIndices)
            {
                if (predecessorIndexString.Length == 0)
                {
                    predecessorIndexString.Append(index);
                }
                else
                {
                    predecessorIndexString.Append("," + index);
                }

            }

            return predecessorIndexString.ToString();
        }

        private bool NodeExists(TreeNode task, string node)
        {
            foreach (TreeNode item in task.Parent.Nodes)
            {
                if (item.Text == node)
                {
                    return true;
                }
            }

            return false;
        }

        private List<string> PredecessorsToPreselectList(TreeNode task)
        {
            List<string> list = new List<string>();
            string[] taskNameArr;

            if (task.PrevNode == null)
                return list;

            taskNameArr = task.Text.Split(' ');

            if (task.Text.Contains("Rough") || task.Text.Contains("Finish") || task.Text.Contains("Electrodes"))
            {
                if (task.Text.Contains("Program"))
                {
                    list.Add("Design");
                }
                else if (task.Text == "CNC Rough" || task.Text == "CNC Electrodes")
                {
                    list.Add("Program " + taskNameArr[1].ToString());
                }
                else if (task.Text.Contains("CNC Finish"))
                {
                    list.Add("Program " + taskNameArr[1].ToString());
                    list.Add(task.PrevNode.Text);
                }
                else if (task.Text.Contains("Inspection"))
                {
                    list.Add("CNC " + taskNameArr[3].ToString());
                }
                else if (task.Text == "Finish Grind")
                {
                    list.Add(task.PrevNode.Text);
                }
            }
            else if (task.Text == "Heat Treat")
            {
                if (NodeExists(task, "Inspection Post CNC Rough"))
                {
                    list.Add("Inspection Post CNC Rough");
                }
                else
                {
                    list.Add("CNC Rough");
                }
            }
            else if (task.Text == "Prep Grind")
            {
                list.Add("Heat Treat");
            }
            else if (task.Text.Contains("EDM Wire"))
            {
                if (task.PrevNode != null)
                    list.Add(task.PrevNode.Text);
            }
            else if (task.Text.Contains("EDM Sinker"))
            {
                if (task.Text == "EDM Sinker")
                {
                    if (NodeExists(task, "Inspection Post CNC Electrodes"))
                    {
                        list.Add("Inspection Post CNC Electrodes");
                    }
                    else if (NodeExists(task, "CNC Electrodes"))
                    {
                        list.Add("CNC Electrodes");
                    }

                    list.Add(task.PrevNode.Text);
                }
                else if (task.Text.Contains("Inspection"))
                {
                    list.Add("EDM Sinker");
                }
            }
            else if (task.Text.Contains("Polish"))
            {
                if (task.Text == "Polish (In-House)" || task.Text == "Polish (Outsource)")
                {
                    list.Add(task.PrevNode.Text);
                }
                else if (task.Text.Contains("Inspection"))
                {
                    if (NodeExists(task, "Polish (In-House)"))
                    {
                        list.Add("Polish (In-House)");
                    }
                    else if (NodeExists(task, "Polish (Outsource)"))
                    {
                        list.Add("Polish (Outsource)");
                    }
                }
            }
            else if (task.Text.Contains("Texturing"))
            {
                list.Add(task.PrevNode.Text);
            }
            else if (task.Text.Contains("Grind-Fitting"))
            {
                list.Add(task.PrevNode.Text);
            }

            return list;
        }

        private void RemoveSelectedNodeFromTree()
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;

            if (selectedNode == null || selectedNode.Level == 0)
            {
                return;
            }
            else if (selectedNode.Level == 1)
            {
                Project.RemoveComponent(selectedNode.Text);
            }
            else if (selectedNode.Level == 2)
            {
                var component = Project.Components.Find(x => x.Component == selectedNode.Parent.Text);
                component.RemoveTask(selectedNode.Index);
            }

            MoldBuildTreeView.SelectedNode.Remove();
            MoldBuildTreeView.Focus();
        }

        private void MoveSelectedNodeUp(TreeNode node)
        {
            TreeNode parent = node.Parent;

            if (node.Level > 0 && node.Level < 3)
            {
                int index = parent.Nodes.IndexOf(node);

                if (index > 0)
                {
                    if (node.Level == 1)
                    {
                        Project.MoveComponentUp(node.Index);
                    }
                    else if (node.Level == 2)
                    {
                        //var component = Project.ComponentList.Where(x => x.Name == node.Parent.Text).First();
                        SelectedComponent.MoveTaskUp(node.Index);
                    }

                    //MoldBuildTreeView.Nodes.Clear();

                    //MoldBuildTreeView.Nodes.Add("ToolNumber*");

                    //LoadProjectToForm(Project);

                    //MoldBuildTreeView.Nodes[0].Expand();

                    //MoldBuildTreeView.EndUpdate();

                    parent.Nodes.RemoveAt(index);
                    parent.Nodes.Insert(index - 1, node);

                    MoldBuildTreeView.SelectedNode = parent.Nodes[index - 1];
                    MoldBuildTreeView.Focus();
                }
            }
        }

        private void MoveSelectedNodeDown(TreeNode node)
        {
            TreeNode parent = node.Parent;

            if (node.Level > 0 && node.Level < 3)
            {
                int index = parent.Nodes.IndexOf(node);

                if (index < parent.Nodes.Count - 1)
                {
                    if (node.Level == 1)
                    {
                        Project.MoveComponentDown(node.Index);
                    }
                    else if (node.Level == 2)
                    {
                        //var component = Project.ComponentList.Where(x => x.Name == node.Parent.Text).First();
                        SelectedComponent.MoveTaskDown(node.Index);
                    }

                    parent.Nodes.RemoveAt(index);
                    parent.Nodes.Insert(index + 1, node);

                    MoldBuildTreeView.SelectedNode = parent.Nodes[index + 1];
                    MoldBuildTreeView.Focus();
                }
            }
        }

        private void SelectNextTask()
        {
            if (MoldBuildTreeView.SelectedNode.Level == 2)
            {
                TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
                // Tree needs to have focus in order to select the next node.
                MoldBuildTreeView.Focus();

                if (selectedNode.Index < selectedNode.Parent.Nodes.Count - 1)
                {
                    MoldBuildTreeView.SelectedNode = selectedNode.Parent.Nodes[selectedNode.Index + 1];
                }
            }
        }

        private void OpenWorkloadSheet()
        {
            FileInfo fi = new FileInfo(@"X:\TOOLROOM\FORMS\Work Load.xlsm");

            if (fi.Exists)
            {
                System.Diagnostics.Process.Start("EXCEL.EXE", "/r \"" + fi.FullName + "\"");
            }
            else
            {
                MessageBox.Show("Can't find Work Load Sheet.");
            }
        }

        private void OpenWorkloadSheetExcelCOM()
        {
            FileInfo fi = new FileInfo(@"X:\TOOLROOM\FORMS\Work Load.xlsm");

            if (fi.Exists)
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                excelApp.DisplayAlerts = false;
                excelApp.Workbooks.Open(@"X:\TOOLROOM\FORMS\Work Load.xlsm", ReadOnly: true);
                excelApp.DisplayAlerts = true;
            }
            else
            {
                MessageBox.Show("Can't find Work Load Sheet.");
            }
        }

        private void PrintObjectTree()
        {
            Console.WriteLine($"{Project.JobNumber} {Project.ProjectNumber} {Project.DueDate} {Project.ToolMaker} {Project.Designer} {Project.RoughProgrammer} {Project.FinishProgrammer} {Project.ElectrodeProgrammer}");

            foreach (ComponentModel component in Project.Components)
            {
                Console.WriteLine($"{component.Component}");

                //foreach(TaskInfo task in component.TaskList)
                //{
                //    Console.WriteLine($"    {task.TaskName}");
                //    Console.WriteLine($"        {task.Hours}");
                //    Console.WriteLine($"        {task.Duration}");
                //    Console.WriteLine($"        {task.Machine}");
                //    Console.WriteLine($"        {task.Personnel}");
                //    Console.WriteLine($"        {task.Predecessors}");
                //    Console.WriteLine($"        {task.Notes}");
                //}
            }
        }

        private void SetProjectInfo()
        {
            Project.SetProjectInfo
            (
                jobNumber: MoldBuildTreeView.Nodes[0].Text,
                projectNumber: ProjectNumberTextBox.Text,
                dueDate: dueDateTimePicker.Value,
                toolMaker: ToolMakerComboBox.Text,
                designer: DesignerComboBox.Text,
                roughProgrammer: RoughProgrammerComboBox.Text,
                electrodeProgrammer: ElectrodeProgrammerComboBox.Text,
                finishProgrammer: FinishProgrammerComboBox.Text,
                edmSinkerOperator: EDMSinkerOperatorComboBox.Text,
                roughCNCOperator: RoughCNCOperatorComboBox.Text,
                electrodeCNCOperator: ElectrodeCNCOperatorComboBox.Text,
                finishCNCOperator: FinishCNCOperatorComboBox.Text,
                edmWireOperator: EDMWireOperatorComboBox.Text
            );
        }

        private void SelectRelatedTasks(string taskName)
        {
            if (TaskListBox.GetSelected(TaskListBox.Items.IndexOf(taskName)))
            {
                if (taskName.Contains("Rough"))
                {
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("Program Rough"), true);
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("CNC Rough"), true);
                }
                else if (taskName.Contains("Electrodes"))
                {
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("Program Electrodes"), true);
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("CNC Electrodes"), true);
                    //TaskListBox.SetSelected(TaskListBox.Items.IndexOf("Inspection Post CNC Electrodes"), true);

                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("EDM Sinker"), true);
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("Inspection Post EDM Sinker"), true);
                }
                else if (taskName.Contains("Finish") && !taskName.Contains("Grind"))
                {
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("Program Finish"), true);
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("CNC Finish"), true);
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("Inspection Post CNC Finish"), true);
                }
                else if (taskName.Contains("Polish"))
                {
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("Inspection Post Polish"), true);
                }
            }
        }

        private void SelectPredecessors(TreeNode selectedNode)
        {
            List<string> predecessorList = new List<string>();

            predecessorsListBox.ClearSelected();

            if (selectedNode.Nodes.Count > 0) // && quoteLoaded == true
            {
                if (selectedNode.Nodes[4].Text == "")
                {
                    // Leave predecessor list blank.
                }
                else if (selectedNode.Nodes[4].Text.Contains(","))
                {
                    predecessorList = selectedNode.Nodes[4].Text.Split(',').ToList();
                }
                else
                {
                    predecessorList.Add(selectedNode.Nodes[4].Text);
                }

                int baseCount = 0;

                //int baseCount = countTasks(MoldBuildTreeView, selectedNode.Parent.Text); // Uncomment to reactivate to count tasks from components that are higher on the list to find task ID.

                foreach (string item in predecessorList)
                {
                    predecessorsListBox.SelectedIndex = Convert.ToInt32(item) - baseCount - 1;
                }
            }
            else
            {
                predecessorList = PredecessorsToPreselectList(selectedNode);

                foreach (string item in predecessorList)
                {
                    if (predecessorsListBox.Items.Contains(item))
                    {
                        predecessorsListBox.SelectedItem = item;
                    }
                }
            }
        }

        private void CheckForTasksWithNoSuccessors()
        {
            string[] preds = null;

            foreach (ComponentModel component in Project.Components)
            {
                List<int> predList = new List<int>();
                int n = 1;

                foreach (TaskModel task in component.Tasks)
                {
                    if (task.Predecessors == "")
                    {

                    }
                    else if (task.Predecessors.Contains(','))
                    {
                        preds = task.Predecessors.Split(',');

                        for (int i = 0; i < preds.Count(); i++)
                        {
                            //if (i < preds.Count() - 1)
                            //{
                            predList.Add(Convert.ToInt16(preds[i]));
                            //}
                            //else
                            //{
                            //newPreds.Append(Convert.ToInt32(preds[i]) + baseNumber);
                            //}
                        }
                    }
                    else
                    {
                        predList.Add(Convert.ToInt16(task.Predecessors));
                    }
                }

                var result = from predInts in predList
                             orderby predInts ascending
                             select predInts;

                foreach (int pred in result)
                {
                    //Console.WriteLine(n + " " + pred);

                    if (n != pred)
                    {
                        Console.WriteLine(n);
                        n = pred;
                        n++;
                    }
                    else
                    {
                        n++;
                    }
                }
            }
        }

        private void CheckComponentForTasksWithNoSuccessors(TreeNode selectedNode)
        {
            string[] preds = null;

            foreach (ComponentModel component in Project.Components)
            {
                List<int> predList = new List<int>();
                int n = 1;

                foreach (TaskModel task in component.Tasks)
                {
                    if (task.Predecessors == "")
                    {

                    }
                    else if (task.Predecessors.Contains(','))
                    {
                        preds = task.Predecessors.Split(',');

                        for (int i = 0; i < preds.Count(); i++)
                        {
                            //if (i < preds.Count() - 1)
                            //{
                            predList.Add(Convert.ToInt16(preds[i]));
                            //}
                            //else
                            //{
                            //newPreds.Append(Convert.ToInt32(preds[i]) + baseNumber);
                            //}
                        }
                    }
                    else
                    {
                        predList.Add(Convert.ToInt16(task.Predecessors));
                    }
                }

                var result = from predInts in predList
                             orderby predInts ascending
                             select predInts;

                foreach (int pred in result)
                {
                    //Console.WriteLine(n + " " + pred);

                    if (n != pred)
                    {
                        Console.WriteLine(n);
                        n = pred;
                        n++;
                    }
                    else
                    {
                        n++;
                    }
                }
            }
        }

        private void ActivateTaskHandlers()
        {
            // TaskInfo controls.
            
            hoursNumericUpDown.ValueChanged += new System.EventHandler(hoursNumericUpDown_ValueChanged);
            hoursNumericUpDown.ValueChanged += new System.EventHandler(TaskInfo_Changed);
            durationNumericUpDown.ValueChanged += new System.EventHandler(TaskInfo_Changed);
            durationUnitsComboBox.TextChanged += new System.EventHandler(TaskInfo_Changed);
            matchHoursCheckBox.CheckStateChanged += new System.EventHandler(TaskInfo_Changed);
            machineComboBox.TextChanged += new System.EventHandler(TaskInfo_Changed);
            personnelComboBox.TextChanged += new System.EventHandler(TaskInfo_Changed);
            predecessorsListBox.SelectedIndexChanged += new System.EventHandler(predecessorsListBox_SelectedIndexChanged);
            predecessorsListBox.SelectedIndexChanged += new System.EventHandler(TaskInfo_Changed);
            taskNotesTextBox.TextChanged += new System.EventHandler(TaskInfo_Changed);
        }

        private void DeactivateTaskHandlers()
        {
            // TaskInfo controls.

            hoursNumericUpDown.ValueChanged -= new System.EventHandler(hoursNumericUpDown_ValueChanged);
            hoursNumericUpDown.ValueChanged -= new System.EventHandler(TaskInfo_Changed);
            durationNumericUpDown.ValueChanged -= new System.EventHandler(TaskInfo_Changed);
            durationUnitsComboBox.TextChanged -= new System.EventHandler(TaskInfo_Changed);
            matchHoursCheckBox.CheckStateChanged -= new System.EventHandler(TaskInfo_Changed);
            machineComboBox.TextChanged -= new System.EventHandler(TaskInfo_Changed);
            personnelComboBox.TextChanged -= new System.EventHandler(TaskInfo_Changed);
            predecessorsListBox.SelectedIndexChanged -= new System.EventHandler(predecessorsListBox_SelectedIndexChanged);
            predecessorsListBox.SelectedIndexChanged -= new System.EventHandler(TaskInfo_Changed);
            taskNotesTextBox.TextChanged -= new System.EventHandler(TaskInfo_Changed);
        }

        private void ActivateComponentHandlers()
        {
            // ComponentInfo controls.
            quantityNumericUpDown.ValueChanged += new System.EventHandler(quantityNumericUpDown_ValueChanged);
            sparesNumericUpDown.ValueChanged += new System.EventHandler(sparesNumericUpDown_ValueChanged);
            materialComboBox.TextChanged += new System.EventHandler(materialComboBox_TextChanged);
            finishTextBox.TextChanged += new System.EventHandler(finishTextBox_TextChanged);
            ComponentPictureEdit.EditValueChanged += new System.EventHandler(ComponentPictureEdit_EditValueChanged);
            componentNotesTextBox.TextChanged += new System.EventHandler(componentNotesTextBox_TextChanged);
        }

        private void DeactivateComponentHandlers()
        {
            // ComponentInfo controls.
            quantityNumericUpDown.ValueChanged -= new System.EventHandler(quantityNumericUpDown_ValueChanged);
            sparesNumericUpDown.ValueChanged -= new System.EventHandler(sparesNumericUpDown_ValueChanged);
            materialComboBox.TextChanged -= new System.EventHandler(materialComboBox_TextChanged);
            finishTextBox.TextChanged -= new System.EventHandler(finishTextBox_TextChanged);
            ComponentPictureEdit.EditValueChanged -= new System.EventHandler(ComponentPictureEdit_EditValueChanged);
            componentNotesTextBox.TextChanged -= new System.EventHandler(componentNotesTextBox_TextChanged);
        }
        private void SaveTemplate(bool autosaved = false)
        {
            string fileName;
            SetProjectInfo();

            if (autosaved == true)
            {
                fileName = @"X:\TOOLROOM\Workload Tracking System\Templates\Created Projects\" + Project.JobNumber + " - #" + Project.ProjectNumber + ".xml";
            }
            else
            {
                //fileName = @"X:\TOOLROOM\Workload Tracking System\Templates\Created Projects\" + Project.JobNumber + " - #" + Project.ProjectNumber + ".xml";
                fileName = Template.SaveTemplateFile(Project.JobNumber + " - #" + Project.ProjectNumber, @"X:\TOOLROOM\Workload Tracking System\Templates");
            }

            if (fileName != "")
            {
                //tmpt.WriteProjectToTextFile(Project, fileName);
                Template.WriteToXmlFile(fileName, Project);
            }
        }
        private void ProjectCreationForm_Shown(object sender, EventArgs e)
        {
            //MessageBox.Show("Shown");
            formLoad = false;
        }

        private void ProjectCreationForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (excelApp != null)
            {
                excelApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(excelApp);
            }
        }

        private void RenameButton_Click(object sender, EventArgs e)
        {
            if (MoldBuildTreeView.SelectedNode == null)
            {
                MessageBox.Show("Please select a node to rename.");
                return;
            }

            RenameNode();
        }

        private void UpButton_Click(object sender, EventArgs e)
        {
            //MoldBuildTreeView.BeginUpdate();

            MoveSelectedNodeUp(MoldBuildTreeView.SelectedNode);


        }

        private void DownButton_Click(object sender, EventArgs e)
        {
            //MoldBuildTreeView.BeginUpdate();

            MoveSelectedNodeDown(MoldBuildTreeView.SelectedNode);

            //MoldBuildTreeView.Nodes.Clear();

            //MoldBuildTreeView.Nodes.Add("ToolNumber*");

            //LoadProjectToForm(Project);

            //MoldBuildTreeView.Nodes[0].Expand();

            

            //MoldBuildTreeView.SelectedNode = parent.Nodes[index + 1];
            //MoldBuildTreeView.Focus();

            //MoldBuildTreeView.EndUpdate();
        }

        private void ASideRadioButton_Click(object sender, EventArgs e)
        {
            prefix = "A-";
            BSideRadioButton.Checked = false;
        }

        private void BSideRadioButton_Click(object sender, EventArgs e)
        {
            prefix = "B-";
            ASideRadioButton.Checked = false;
        }

        private void ComponentListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComponentTextBox.Text = prefix + ComponentListBox.Text;
        }

        private void ResourceComboBox_DropDown(object sender, EventArgs e)
        {
            System.Windows.Forms.ComboBox combo = sender as System.Windows.Forms.ComboBox;
            combo.SelectedValueChanged -= Personnel_ValueChanged;
            PopulateComboBox(combo);
            combo.SelectedValueChanged += Personnel_ValueChanged;
        }

        private void TaskInfo_Changed(object sender, EventArgs e)
        {
            UpdateButton.Appearance.BackColor = Color.Orange;
            Project.IsChanged = true;
        }

        private void hoursNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            //AddTaskInfoToSelectedTask("Hours", hoursNumericUpDown.Value.ToString() + " Hour(s)");
            UpdateDuration();
        }

        private void matchHoursCheckBox_CheckStateChanged(object sender, EventArgs e)
        {
            UpdateDuration();
        }

        private void predecessorsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            foreach (int index in predecessorsListBox.SelectedIndices)
            {
                if (selectedNode.Index == index)
                {
                    MessageBox.Show("A task cannot be its own predecessor.");
                    //Console.WriteLine(Project.ComponentList[selectedNode.Parent.Index].TaskList[selectedNode.Index].Predecessors);

                    //SelectPredecessors(selectedNode);
                }
            }
        }

        private void MoldBuildTreeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                //if (MoldBuildTreeView.SelectedNode.Level == 1)
                //{
                //    var result = XtraInputBox.Show("Copied Component Name", "Copy Component", SelectedComponent.Component);

                //    if (result.Length > 0) // Editor returns empty string when cancel is clicked.
                //    {
                //        AddCopiedComponentToTree(new ComponentModel(SelectedComponent, result));
                //    }
                //}
            }
        }

        private void MoldBuildTreeView_BeforeSelect(object sender, TreeViewCancelEventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;

            if (selectedNode != null)
            {
                selectedNode.BackColor = Color.White;
                selectedNode.ForeColor = Color.Black;
            }
        }

        private void MoldBuildTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            string[] str1Arr, str2Arr;
            TaskModel taskInfo;

            if (selectedNode.Level == 0 && formLoad == false)
            {
                tabControl1.SelectedTab = tabPage1;
                tabControl2.Enabled = true;
                tabControl2.SelectedTab = tabPage6;
            }
            else if (selectedNode.Level == 1)
            {
                DeactivateComponentHandlers();

                SelectedComponent = Project.Components.ElementAt(selectedNode.Index);
                quantityNumericUpDown.Value = SelectedComponent.Quantity;
                sparesNumericUpDown.Value = SelectedComponent.Spares;
                materialComboBox.Text = SelectedComponent.Material;
                finishTextBox.Text = SelectedComponent.Finish;
                ComponentPictureEdit.Image = SelectedComponent.picture;

                if (selectedNode.ContextMenuStrip == null)
                {
                    selectedNode.ContextMenuStrip = ComponentMenu;
                }

                //if(ActiveComponent.Picture.Count > 0)
                //{
                //    componentPictureBox.Image = ActiveComponent.Picture[0];
                //}
                //else
                //{
                //    componentPictureBox.Image = null;
                //}

                componentNotesTextBox.Text = SelectedComponent.Notes;

                tabControl1.SelectedTab = tabPage5;
                tabControl2.Enabled = true;
                tabControl2.SelectedTab = tabPage7;

                ActivateComponentHandlers();
            }
            else if (selectedNode.Level == 2)
            {
                try
                {
                    DeactivateTaskHandlers();

                    tabControl1.SelectedTab = tabPage4;
                    tabControl2.Enabled = false;

                    SelectedComponent = Project.Components.ElementAt(selectedNode.Parent.Index);

                    SelectedTask = SelectedComponent.Tasks.ElementAt(selectedNode.Index);

                    machineComboBox.DataSource = GetMachineList(SelectedTask.TaskName);
                    personnelComboBox.DataSource = GetPersonnelList(SelectedTask.TaskName);
                    predecessorsListBox.DataSource = GetPredecessorList(SelectedComponent);

                    predecessorsListBox.ClearSelected();

                    taskInfo = GetDefaultTaskInfo(SelectedTask.TaskName);

                    if (selectedNode.ContextMenuStrip == null)
                    {
                        selectedNode.ContextMenuStrip = TaskMenu;
                    }

                    if (selectedNode.Nodes.Count > 0)
                    {
                        str1Arr = selectedNode.Nodes[0].Text.Split(' ');
                        str2Arr = selectedNode.Nodes[1].Text.Split(' ');

                        if (int.TryParse(str1Arr[0], out int hours))
                        {
                            hoursNumericUpDown.Value = hours;
                        }
                        else
                        {
                            hoursNumericUpDown.Value = 0;
                        }

                        if (int.TryParse(str2Arr[0], out int duration))
                        {
                            durationNumericUpDown.Value = duration;
                        }
                        else
                        {
                            durationNumericUpDown.Value = 0;
                        }

                        machineComboBox.SelectedText = selectedNode.Nodes[2].Text;
                        personnelComboBox.SelectedText = selectedNode.Nodes[3].Text;

                        //int baseCount = countTasks(MoldBuildTreeView, selectedNode.Parent.Text); // Uncomment to reactivate to count tasks from components that are higher on the list to find task ID.

                        //int baseCount = 0;

                        //foreach (string item in predecessorList)
                        //{
                        //    predecessorsListBox.SelectedIndex = Convert.ToInt32(item) - baseCount - 1;
                        //}

                        taskNotesTextBox.Text = selectedNode.Nodes[5].Text;
                    }
                    else
                    {
                        hoursNumericUpDown.Value = taskInfo.Hours;
                        durationNumericUpDown.Value = Convert.ToDecimal(taskInfo.Duration);
                        machineComboBox.SelectedText = taskInfo.Machine;
                        personnelComboBox.SelectedText = taskInfo.Personnel;

                        //predecessorList = predecessorsToPreselectList(selectedNode);

                        //foreach (string item in predecessorList)
                        //{
                        //    if (predecessorsListBox.Items.Contains(item))
                        //    {
                        //        predecessorsListBox.SelectedItem = item;
                        //    }
                        //}

                        taskNotesTextBox.Text = "";
                    }

                    SelectPredecessors(selectedNode);

                    ActivateTaskHandlers();
                }
                catch (Exception er)
                {
                    MessageBox.Show($"{er.Message}\n\n{er.StackTrace}");
                }
            }
        }

        private void MoldBuildTreeView_Leave(object sender, EventArgs e)
        {
            if (MoldBuildTreeView.SelectedNode != null)
            {
                TreeNode selectedNode = MoldBuildTreeView.SelectedNode;

                selectedNode.BackColor = SystemColors.Highlight;
                selectedNode.ForeColor = SystemColors.HighlightText;
            }
        }
        private void ProjectMenuStrip_Click(object sender, EventArgs e)
        {
            ContextMenuStrip contextMenuStrip = (ContextMenuStrip)sender;

            ToolStripMenuItem selectedToolStripMenuItem = null;

            ComponentModel componentToAdd = new ComponentModel();

            string fileName;

            foreach (ToolStripMenuItem item in contextMenuStrip.Items)
            {
                if (item.Pressed == true)
                {
                    selectedToolStripMenuItem = item;
                }
            }

            if (selectedToolStripMenuItem.Text == "Rename")
            {
                RenameNode();
            }
            else if (selectedToolStripMenuItem.Text == "Load Component")
            {
                contextMenuStrip.Close();
                fileName = Template.OpenTemplateFile(@"X:\TOOLROOM\Workload Tracking System\Templates\Components");
                if (fileName.Length > 0)
                {
                    componentToAdd = Template.ReadFromXmlFile<ComponentModel>(fileName);
                    foreach (TaskModel task in componentToAdd.Tasks)
                    {
                        task.Personnel = GetTaskPersonnel(task.TaskName);
                        task.ID = 0;
                    }
                    AddCopiedComponentToTree(componentToAdd); // Adds component to project object.
                }
            }
        }
        private void ComponentMenuStrip_Click(object sender, EventArgs e)
        {
            ContextMenuStrip contextMenuStrip = (ContextMenuStrip)sender;

            ToolStripMenuItem selectedToolStripMenuItem = null;

            List<TaskModel> tasksToAdd = new List<TaskModel>();

            string fileName;

            foreach (ToolStripMenuItem item in contextMenuStrip.Items)
            {
                if (item.Pressed == true)
                {
                    selectedToolStripMenuItem = item;
                }
            }

            try
            {
                if (selectedToolStripMenuItem.Text == "Rename")
                {
                    RenameNode();
                }
                else if (selectedToolStripMenuItem.Text == "Copy")
                {
                    var result = XtraInputBox.Show("Copied Component Name", "Copy Component", SelectedComponent.Component);

                    if (result.Length > 0) // Editor returns empty string when cancel is clicked.
                    {
                        AddCopiedComponentToTree(new ComponentModel(SelectedComponent, result));
                    }
                }
                else if (selectedToolStripMenuItem.Text == "Create Template")
                {
                    contextMenuStrip.Close();
                    fileName = Template.SaveTemplateFile(SelectedComponent.Component, @"X:\TOOLROOM\Workload Tracking System\Templates\Components");
                    if (fileName.Length > 0)
                    {
                        Template.WriteToXmlFile(fileName, SelectedComponent);
                    }
                }
                else if (selectedToolStripMenuItem.Text == "Load Template")
                {
                    contextMenuStrip.Close();
                    fileName = Template.OpenTemplateFile(@"X:\TOOLROOM\Workload Tracking System\Templates\Components");
                    if (fileName.Length > 0)
                    {
                        tasksToAdd = Template.ReadFromXmlFile<ComponentModel>(fileName).Tasks;
                        SelectedComponent.Tasks.AddRange(tasksToAdd);
                        foreach (TaskModel task in SelectedComponent.Tasks)
                        {
                            task.Personnel = GetTaskPersonnel(task.TaskName);
                            task.ID = 0;
                        }
                        AddTaskListToTreeView(MoldBuildTreeView.SelectedNode, tasksToAdd);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }
        private void TaskMenuStrip_Click(object sender, EventArgs e)
        {
            ContextMenuStrip contextMenuStrip = (ContextMenuStrip)sender;

            ToolStripMenuItem selectedToolStripMenuItem = null;

            foreach (ToolStripMenuItem item in contextMenuStrip.Items)
            {
                if (item.Pressed == true)
                {
                    selectedToolStripMenuItem = item;
                }
            }

            try
            {
                if (selectedToolStripMenuItem.Text == "Rename")
                {
                    RenameNode();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine();
            }
        }
        private void updateInfoButton_Click(object sender, EventArgs e)
        {
            SetTaskInfoForSelectedTask();
            UpdateButton.Appearance.BackColor = Color.Transparent;
        }

        private void ProjectNumberTextBox_TextChanged(object sender, EventArgs e)
        {
            ProjectNumberTextBox.BackColor = Color.White;

            if (Project.SetProjectNumber(ProjectNumberTextBox.Text))
            {
                Console.WriteLine($"Project set to {Project.ProjectNumber}.");
            }
            else
            {
                ProjectNumberTextBox.Text = Project.ProjectNumber.ToString();
            }
        }
        private void ProjectNumberTextBox_Leave(object sender, EventArgs e)
        {
            //MessageBox.Show("Left Project # Box");
        }
        private void ToolMakerComboBox_TextChanged(object sender, EventArgs e)
        {
            ToolMakerComboBox.BackColor = Color.White;
            Project.ToolMaker = ToolMakerComboBox.Text;
        }
        private void AddTasksButton_Click(object sender, EventArgs e)
        {
            AddSelectedTasksToSelectedComponent();
        }
        private void AddComponentButton_Click(object sender, EventArgs e)
        {
            AddComponentToTree(ComponentTextBox.Text);
        }
        private void TaskListBox_MouseClick(object sender, MouseEventArgs e)
        {
            string selectedItemName = TaskListBox.Items[TaskListBox.IndexFromPoint(e.Location)].ToString();
            SelectRelatedTasks(selectedItemName);
            //MessageBox.Show(selectedItemName);
        }

        private void GetQuoteButton_Click(object sender, EventArgs e)
        {
            //checkForTasksWithNoSuccessors();

            try
            {
                string filename;

                OpenFileDialog snapshotOpenFileDialog = new OpenFileDialog
                {
                    InitialDirectory = @"C:\Users\" + Environment.UserName + @"\Downloads",
                    Filter = "Excel Files (*.xlsm, *.xlsx)|*.xlsm;*.xlsx"
                };

                Nullable<bool> result = Convert.ToBoolean(snapshotOpenFileDialog.ShowDialog());

                if (result == true)
                {
                    filename = snapshotOpenFileDialog?.FileName;

                    if (filename == "")
                    {
                        return;
                    }

                    Project.SetQuoteInfo(ExcelInteractions.GetQuoteInfo(filename));

                    LoadProjectToForm(ConvertQuoteToProject(Project));
                    //quoteLoaded = true;
                }
                else
                {
                    return;
                }

                Console.WriteLine($"{Project.QuoteInfo.ProgramRoughHours} {Project.QuoteInfo.ProgramFinishHours} {Project.QuoteInfo.ProgramElectrodeHours} {Project.QuoteInfo.CNCRoughHours} {Project.QuoteInfo.CNCFinishHours} {Project.QuoteInfo.CNCElectrodeHours} {Project.QuoteInfo.EDMSinkerHours}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }
        private void Personnel_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                //SetProjectInfo();

                SetPersonnnel(sender);

                MoldBuildTreeView.SelectedNode = MoldBuildTreeView.Nodes[0];

                MoldBuildTreeView.Nodes[0].Nodes.Clear();

                LoadProjectToForm(Project);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #region Component Controls

        private void quantityNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            if (MoldBuildTreeView.SelectedNode.Level == 1)
            {
                SelectedComponent.SetQuantity(Convert.ToInt16(quantityNumericUpDown.Value));
            }
        }

        private void sparesNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            if (MoldBuildTreeView.SelectedNode.Level == 1)
            {
                SelectedComponent.SetSpares(Convert.ToInt16(sparesNumericUpDown.Value));
            }
        }

        private void materialComboBox_TextChanged(object sender, EventArgs e)
        {
            if (MoldBuildTreeView.SelectedNode.Level == 1)
            {
                SelectedComponent.SetMaterial(materialComboBox.Text);
            }
        }

        private void componentNotesTextBox_TextChanged(object sender, EventArgs e)
        {
            if (MoldBuildTreeView.SelectedNode.Level == 1)
            {
                SelectedComponent.SetNote(componentNotesTextBox.Text);
            }
        }

        private void finishTextBox_TextChanged(object sender, EventArgs e)
        {
            if (MoldBuildTreeView.SelectedNode.Level == 1)
            {
                SelectedComponent.SetFinish(finishTextBox.Text);
            }
        } 

        private void ComponentPictureEdit_EditValueChanged(object sender, EventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;

            if (selectedNode != null && selectedNode.Level == 1)
            {
                if (ComponentModel.IsGoodComponentPicture(ComponentPictureEdit.Image) == false)
                {
                    ComponentPictureEdit.EditValueChanged -= new System.EventHandler(this.ComponentPictureEdit_EditValueChanged);
                    ComponentPictureEdit.Image = null;
                    ComponentPictureEdit.EditValueChanged += new System.EventHandler(this.ComponentPictureEdit_EditValueChanged);

                    return;
                }

                SelectedComponent.picture = ComponentPictureEdit.Image;
            }
            else
            {
                MessageBox.Show("Please select a Component to add a picture to.");
                ComponentPictureEdit.EditValueChanged -= new System.EventHandler(this.ComponentPictureEdit_EditValueChanged);
                ComponentPictureEdit.Image = null;
                ComponentPictureEdit.EditValueChanged += new System.EventHandler(this.ComponentPictureEdit_EditValueChanged);
            }
        }
        #endregion

        private void overlapAllowedCheckEdit_CheckedChanged(object sender, EventArgs e)
        {
            if (overLapAllowedCheckEdit.Checked == true)
            {
                Project.OverlapAllowed = true;
            }
            else
            {
                Project.OverlapAllowed = false;
            }
        }

        private void SelectProjectButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (ProjectSelectionForm form = new ProjectSelectionForm())
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        if (form.Project.MWONumber != 0)
                        {
                            this.ProjectNumberTextBox.Text = form.Project.MWONumber.ToString();
                        }
                        else if (form.Project.ProjectNumber != 0)
                        {
                            this.ProjectNumberTextBox.Text = form.Project.ProjectNumber.ToString();
                        }
                        else
                        {
                            this.ProjectNumberTextBox.Text = "";
                        }

                        this.MoldBuildTreeView.Nodes[0].Text = form.Project.JobNumber;
                        this.dueDateTimePicker.Value = form.Project.DueDate;
                        this.ToolMakerComboBox.Text = form.Project.ToolMaker;
                        this.DesignerComboBox.Text = form.Project.Designer;
                        this.RoughProgrammerComboBox.Text = form.Project.RoughProgrammer;
                        this.ElectrodeProgrammerComboBox.Text = form.Project.ElectrodeProgrammer;
                        this.FinishProgrammerComboBox.Text = form.Project.FinishProgrammer;
                        this.Project.Apprentice = form.Project.Apprentice;
                        this.Project.Customer = form.Project.Customer;
                        this.Project.Project = form.Project.Name;
                    }
                    else
                    {

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }
        private void saveTemplateButton_Click(object sender, EventArgs e)
        {
            int projectNumberResult;

            try
            {
                if (MoldBuildTreeView.Nodes[0].Text == "Tool Number*")
                {
                    MessageBox.Show("Please enter a tool number.");
                    MoldBuildTreeView.Nodes[0].BackColor = Color.Red;
                    MoldBuildTreeView.SelectedNode = MoldBuildTreeView.Nodes[0];
                    MoldBuildTreeView.Focus();
                    return;
                }

                if (ProjectNumberTextBox.Text == "")
                {
                    MessageBox.Show("Please enter a project number.");
                    ProjectNumberTextBox.BackColor = Color.Red;
                    tabControl1.SelectedTab = tabPage1;
                    return;
                }

                if (!int.TryParse(ProjectNumberTextBox.Text, out projectNumberResult))
                {
                    MessageBox.Show("Please enter a number for project number.");
                    ProjectNumberTextBox.BackColor = Color.Red;
                    tabControl1.SelectedTab = tabPage1;
                    return;
                }

                SaveTemplate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void loadTemplateButton_Click(object sender, EventArgs e)
        {
            string fileName = Template.OpenTemplateFile(@"X:\TOOLROOM\Workload Tracking System\Templates");
            ProjectModel tempProject = new ProjectModel();
            ComponentModel tempComponent = new ComponentModel();
            List<ComponentModel> componentsToRemove = new List<ComponentModel>();
            Console.WriteLine("Load Template Button Click.");

            try
            {
                if (fileName != "")
                {
                    // Project Info insertion removed per Mark's request 9/1/2020.  Reactivated upon consideration of request from Barry Black. 9/4/2020

                    DialogResult dialogResult = MessageBox.Show("Do you want to load project info from this template in addition to components? \n\n" +
                                                                "Existing project info will be overwritten.", "Load Project Info?", MessageBoxButtons.YesNo);

                    if (fileName.Split('.')[fileName.Count(x => x == '.')] == "txt")
                    {
                        tempProject = Template.ReadProjectFromTextFile(fileName, SchedulerStorageProp);
                    }
                    else if (fileName.Split('.')[fileName.Count(x => x == '.')] == "xml")
                    {
                        tempProject = Template.ReadFromXmlFile<ProjectModel>(fileName); 
                    }

                    if (dialogResult == DialogResult.Yes)
                    {
                        LoadProjectInfoToForm(tempProject);
                    }

                    foreach (ComponentModel component in tempProject.Components)
                    {
                        if (Project.Components.Exists(x => x.Component == component.Component))
                        {
                            DialogResult dialogResult2 = MessageBox.Show($"There's already a component named {component.Component}.\n\n" +
                                                                         $"Do you wish to overwrite it?", "Overwrite?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                            if (dialogResult2 == DialogResult.Yes)
                            {
                                tempComponent = Project.Components.Find(x => x.Component == component.Component);

                                tempComponent.UpdateComponent(component);
                            }
                        }
                        else
                        {
                            Project.AddComponent(component);
                        }
                    }

                    MoldBuildTreeView.SelectedNode = MoldBuildTreeView.Nodes[0];

                    MoldBuildTreeView.Nodes[0].Nodes.Clear();

                    LoadProjectToForm(Project);

                    //printObjectTree();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            RemoveSelectedNodeFromTree();
        }

        private void CreateProjectButton_Click(object sender, EventArgs e)
        {
            DataValidated = true;

            if (CreateProjectButton.Text == "Create")
            {
                if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanCreateProjects))
                {
                    MessageBox.Show("This login is not authorized to create projects.");
                    return;
                } 
            }

            if (CreateProjectButton.Text == "Change")
            {
                if (!UserList.Exists(x => x.LoginName == Environment.UserName.ToString().ToLower() && x.CanChangeProjectData))
                {
                    MessageBox.Show("This login is not authorized to change data.");
                    return;
                }
            }

            if (MoldBuildTreeView.Nodes[0].Text == "Tool Number*")
            {
                MessageBox.Show("Please enter a Tool Number or Project Name.");
                MoldBuildTreeView.Nodes[0].BackColor = Color.Red;
                MoldBuildTreeView.SelectedNode = MoldBuildTreeView.Nodes[0];
                MoldBuildTreeView.Focus();
                return;
            }
            else if (MoldBuildTreeView.Nodes[0].Text.Contains('#'))
            {
                MessageBox.Show("Tool Number cannot contain a '#' symbol.");
                MoldBuildTreeView.Nodes[0].BackColor = Color.Red;
                MoldBuildTreeView.SelectedNode = MoldBuildTreeView.Nodes[0];
                MoldBuildTreeView.Focus();
                return;
            }
            else if (MoldBuildTreeView.Nodes[0].Text.Contains(' '))
            {
                MessageBox.Show("A Tool Number cannot contain spaces.  Use underscore instead.");
                MoldBuildTreeView.Nodes[0].BackColor = Color.Red;
                MoldBuildTreeView.SelectedNode = MoldBuildTreeView.Nodes[0];
                MoldBuildTreeView.Focus();
                return;
            }
            else if (MoldBuildTreeView.Nodes[0].Text.Length > 20)
            {
                MessageBox.Show("A Tool Number can't have more than 20 characters.");
                MoldBuildTreeView.Nodes[0].BackColor = Color.Red;
                MoldBuildTreeView.SelectedNode = MoldBuildTreeView.Nodes[0];
                MoldBuildTreeView.Focus();
                return;
            }

            if (ProjectNumberTextBox.Text == "")
            {
                MessageBox.Show("Project must have a Project # or a Work Order #.");
                ProjectNumberTextBox.BackColor = Color.Red;
                tabControl1.SelectedTab = tabPage1;
                return;
            }
            else if (ToolMakerComboBox.Text == "")
            {
                MessageBox.Show("Project must have a lead or tool maker.");
                ToolMakerComboBox.BackColor = Color.Red;
                tabControl1.SelectedTab = tabPage1;
                return;
            }

            if (Project.Components.Count == 0)
            {
                MessageBox.Show("No components entered.");
                return;
            }

            foreach (var item in Project.Components)
            {
                item.AllTasksDated = item.CheckIfAllTasksDated();

                if (item.Component.Length > ComponentModel.ComponentCharacterLimit)
                {
                    MessageBox.Show($"Component: '{item.Component}' is greater than {ComponentModel.ComponentCharacterLimit} characters. \n\nPlease shorten name.");
                    return;
                }
            }

            if (missingTaskInfo == true)
            {
                return;
            }

            if (Project.HasSelfReferencingPredecessors() || Project.HasIsolatedTasks()) // || Project.HasIsolatedTasks()
            {
                // MessageBox is in the HasSelfReferencingPredecessors method.
                return;
            }

            try
            {
                SetProjectInfo();
                Project.SetActiveCounts();

                //Stopwatch sw = new Stopwatch();

                //sw.Start();

                if (CreateProjectButton.Text == "Create")
                {
                    if (Database.CreateProject(Project))
                    {
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else if (CreateProjectButton.Text == "Change")
                {
                    if (Database.UpdateWholeProject(Project))
                    {
                        this.DialogResult = DialogResult.OK;

                    }
                }

                //sw.Stop();

                //MessageBox.Show(sw.ElapsedMilliseconds.ToString() + " ms");

                // Automatically save project when created or changed.
                // See what happens when a matching template is overwritten.

                SaveTemplate(true);

                // Save this in case I get a request to go back to asking to save a template.

                //DialogResult result = MessageBox.Show("Would you like to save a template?", "Template", MessageBoxButtons.YesNo);

                //if (result == DialogResult.Yes)
                //{
                //    SaveTemplate(true);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }

            //printObjectTree();
        }
    }
}
