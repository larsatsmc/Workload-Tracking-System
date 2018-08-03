using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using ClassLibrary;

namespace Toolroom_Scheduler
{
    /// <summary>
    /// Class for the Project Creation Form.
    /// </summary> 
    public partial class Project_Creation_Form : Form
    {
		Excel.Application excelApp;
		string prefix, component, taskName;
        int ID, taskCount;
        bool formLoad = false;
        bool missingTaskInfo = false;
        bool quoteLoaded = false;

        public ProjectInfo Project { get; private set; }
        public bool DataValidated { get; private set; }

        /// <summary>
        /// Initializes a new Project Form.
        /// </summary> 
        public Project_Creation_Form()
        {
			Console.WriteLine("Project Creation Form Load");

            formLoad = true;
            Project = new ProjectInfo();

            InitializeComponent();
        }
        /// <summary>
        /// Initializes a new Project Form and sets the project property of the form to an instance of a property.
        /// </summary> 
        public Project_Creation_Form(ProjectInfo project)
        {
            Console.WriteLine("Project Creation Form Load");

            formLoad = true;
            this.Project = project;

            InitializeComponent();
        }

		private void PopulateComboBox(ComboBox cb)
		{
            Database db = new Database();

            if (cb.Name == "ToolMakerComboBox")
            {
                cb.DataSource = db.GetResourceList("Tool Maker");
            }
            else if (cb.Name == "DesignerComboBox")
            {
                cb.DataSource = db.GetResourceList("Designer");
            }
            else if (cb.Name == "RoughProgrammerComboBox")
            {
                cb.DataSource = db.GetResourceList("Rough Programmer");
            }
            else if (cb.Name == "FinishProgrammerComboBox")
            {
                cb.DataSource = db.GetResourceList("Finish Programmer");
            }
            else if (cb.Name == "ElectrodeProgrammerComboBox")
            {
                cb.DataSource = db.GetResourceList("Electrode Programmer");
            }
        }

		private void RenameNode(string newName)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            if (selectedNode == null || newName == "") return;
            MoldBuildTreeView.SelectedNode.Text = newName;

            if(selectedNode.Level == 0 && selectedNode.BackColor == Color.Red)
            {
                selectedNode.BackColor = Color.White;
                selectedNode.ForeColor = Color.Black;
            }
            else if(selectedNode.Level == 1)
            {
                if (!Project.ComponentNameExists(newName))
                {
                    ClassLibrary.Component component = Project.ComponentList.ElementAt(selectedNode.Index);
                    component.SetName(newName);
                }
            }
            else if(selectedNode.Level == 2)
            {
                ClassLibrary.Component component = Project.ComponentList.ElementAt(selectedNode.Parent.Index);
                TaskInfo task = component.TaskList.ElementAt(selectedNode.Index);
                task.SetName(newName);
            }
        }

        private int CountTasks(TreeView treeView, string component)
        {
            taskCount = 0;
            TreeNodeCollection nodes = treeView.Nodes[0].Nodes;

            try
            {

                //if (nodes.Count == 1)
                //{
                //    return taskCount + 1;
                //}

                foreach (TreeNode n1 in nodes)
                {

                    if (n1.Text == component)
                    {
                        return taskCount;
                    }

                    foreach (TreeNode n2 in n1.Nodes)
                    {
                        taskCount++;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return 0;
        }

        private List<TaskInfo> ReadTree(TreeView treeView, bool isTemplate)
        {
            // Print each node recursively.
            TreeNodeCollection nodes = treeView.Nodes;
            //TaskInfo[] tiArr = new TaskInfo[treeView.GetNodeCount(true) - 1];
            List<TaskInfo> tiList = new List<TaskInfo>();
            ID = 0;

            try
            {
                foreach (TreeNode n in nodes)
                {
                    ReadTreeRecursive(tiList, n, isTemplate);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return tiList;
        }

        private List<TaskInfo> ReadTree(TreeView treeView, bool isTemplate, ProjectInfo pi)
        {
            // Print each node recursively.
            TreeNodeCollection nodes = treeView.Nodes;
            //TaskInfo[] tiArr = new TaskInfo[treeView.GetNodeCount(true) - 1];
            List<TaskInfo> tiList = new List<TaskInfo>();
            ID = 0;

            try
            {
                foreach (TreeNode n in nodes)
                {
                    ReadTreeRecursive(tiList, n, isTemplate, pi);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return tiList;
        }

        private void ReadTreeRecursive(TreeNode treeNode)
        {
            // Print the node.
            TreeNode parent = treeNode.Parent;

            if (treeNode.Level == 0)
            {
                Console.WriteLine("MOLD: " + treeNode.Text);
            }
            else if (treeNode.Level == 1)
            {
                //Console.WriteLine("  COMPONENT: " + treeNode.Text);
                //component = treeNode.Text;
                //ID++;
                //taskInfoList.Add(new TaskInfo(ID, "", component, true));
            }
            else if (treeNode.Level == 2)
            {
                Console.WriteLine("    " + treeNode.Text);
                taskName = treeNode.Text;
                taskCount++;

            }
            // MessageBox.Show(treeNode.Text);
            // Print each node recursively.
            foreach (TreeNode tn in treeNode.Nodes)
            {
                ReadTreeRecursive(tn);
            }
        }

        private void ReadTreeRecursive(List<TaskInfo> taskInfoList, TreeNode treeNode, bool isTemplate)
        {
            // Print the node.
            TreeNode parent = treeNode.Parent;

            if (treeNode.Level == 0)
            {
                Console.WriteLine("MOLD: " + treeNode.Text);
                ID++;
                
            }
            else if (treeNode.Level == 1)
            {
                Console.WriteLine("  COMPONENT: " + treeNode.Text);
                component = treeNode.Text;
                ID++;
                taskInfoList.Add(new TaskInfo(ID, "", component, true));
            }
            else if(treeNode.Level == 2)
            {
                Console.WriteLine("    " + treeNode.Text);
                taskName = treeNode.Text;
                ID++;
                if(treeNode.Nodes.Count > 0)
                {
                    taskInfoList.Add(new TaskInfo(ID, taskName, component, false, treeNode.Nodes[0].Text, treeNode.Nodes[1].Text, treeNode.Nodes[2].Text, treeNode.Nodes[3].Text, treeNode.Nodes[4].Text, treeNode.Nodes[5].Text));
                }
                else if(isTemplate)
                {
                    taskInfoList.Add(new TaskInfo(ID, taskName, "", false));
                }
                else if(!isTemplate)
                {
                    MessageBox.Show("No task info entered for " + taskName + " of the " + component);
                    missingTaskInfo = true;
                }
            }
            // MessageBox.Show(treeNode.Text);
            // Print each node recursively.
            foreach (TreeNode tn in treeNode.Nodes)
            {
                ReadTreeRecursive(taskInfoList, tn, isTemplate);
            }
        }

        private void ReadTreeRecursive(List<TaskInfo> taskInfoList, TreeNode treeNode, bool isTemplate, ProjectInfo pi)
        {
            // Print the node.
            TreeNode parent = treeNode.Parent;

            if (treeNode.Level == 0)
            {
                Console.WriteLine("MOLD: " + treeNode.Text);
                //ID++;
            }
            else if (treeNode.Level == 1)
            {
                Console.WriteLine("  COMPONENT: " + treeNode.Text);
                component = treeNode.Text;
                //ID++;
                //taskInfoList.Add(new TaskInfo(ID, "", component, true));
            }
            else if (treeNode.Level == 2)
            {
                Console.WriteLine("    " + treeNode.Text);
                taskName = treeNode.Text;
                ID++;
                if (treeNode.Nodes.Count > 0)
                {
                    taskInfoList.Add(new TaskInfo(ID, pi.JobNumber, pi.ProjectNumber, taskName, component, false, treeNode.Nodes[0].Text, treeNode.Nodes[1].Text, treeNode.Nodes[2].Text, treeNode.Nodes[3].Text, treeNode.Nodes[4].Text, treeNode.Nodes[5].Text));
                }
                else if (!isTemplate)
                {
                    MessageBox.Show("No task info entered for " + taskName + " of the " + component);
                    missingTaskInfo = true;
                }
            }
            // MessageBox.Show(treeNode.Text);
            // Print each node recursively.
            foreach (TreeNode tn in treeNode.Nodes)
            {
                ReadTreeRecursive(taskInfoList, tn, isTemplate, pi);
            }
        }

        private void AddComponentToTree(string newNodeName)
        {
            if (newNodeName == "" || !Project.AddComponent(newNodeName)) return;

			TreeNode newNode = new TreeNode(newNodeName);
			MoldBuildTreeView.Nodes[0].Nodes.Add(newNode);

            if(MoldBuildTreeView.Nodes[0].Nodes.Count == 1)
            {
                MoldBuildTreeView.Nodes[0].Expand();
            }

            MoldBuildTreeView.SelectedNode = MoldBuildTreeView.Nodes[0].LastNode;
        }

        private void SelectInterdependentTasks(string selectedTask)
        {
            if(selectedTask == "Program Rough" || selectedTask == "CNC Rough" || selectedTask == "Inspection Post CNC Rough")
            {

            }
        }

        private void AddSelectedTasksToSelectedComponent()
        {
            string processName;
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            var item = TaskListBox.SelectedItem;

            if (selectedNode == null || item == null || selectedNode.Level != 1)
            {
                MessageBox.Show("Please select a component to add tasks to or select tasks to add to a component.");
                return;
            }

            foreach (int i in TaskListBox.SelectedIndices)
            {
                var component = Project.ComponentList.Where(x => x.Name == selectedNode.Text).First();
                processName = TaskListBox.Items[i].ToString();
                MoldBuildTreeView.SelectedNode.Nodes.Add(processName);
                component.AddTask(processName, component.Name);
            }
            
        }

        private void AddTaskInfoToSelectedTask(string taskInfoField, string taskInfoEntry)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            List<string> fieldList = new List<string> {"No Hours", "No Duration", "No Machine", "No Personnel", "No Predecessors", "No Note" };

            if (selectedNode.Level != 2)
            {
                MessageBox.Show("Please select a task to add info to.");
                return;
            }

            if (selectedNode.Nodes.Count == 0 && !TaskInfoIsEmpty())
            {
                foreach (string text in fieldList)
                {
                    MoldBuildTreeView.SelectedNode.Nodes.Add(text);
                }
                
            }
            else if (selectedNode.Nodes.Count != 0 && TaskInfoIsEmpty())
            {
                for(int i = selectedNode.Nodes.Count - 1; i >= 0; i--)
                {
                    selectedNode.Nodes[i].Remove();
                }
            }

            if (selectedNode.Nodes.Count != 0)
            {
                if (taskInfoField == "Hours")
                {
                    selectedNode.Nodes[0].Text = taskInfoEntry;
                }
                else if (taskInfoField == "Duration")
                {
                    selectedNode.Nodes[1].Text = taskInfoEntry;
                }
                else if (taskInfoField == "Machine")
                {
                    selectedNode.Nodes[2].Text = taskInfoEntry;
                }
                else if (taskInfoField == "Personnel")
                {
                    selectedNode.Nodes[3].Text = taskInfoEntry;
                }
                else if (taskInfoField == "Predecessor")
                {
                    selectedNode.Nodes[4].Text = taskInfoEntry;
                }
                else if (taskInfoField == "Note")
                {
                    selectedNode.Nodes[5].Text = taskInfoEntry;
                }
            } 

        }

        private void SetTaskInfoForSelectedTask()
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            ClassLibrary.Component component;
            TaskInfo task;
            //string predecessorString = getSelectedPredecessorText(predecessorsListBox); // Uncomment to use project.

            // Check if task is selected.
            if (selectedNode.Level != 2)
            {
                MessageBox.Show("Please select a task to add info to.");
                return;
            }

            string predecessorString = GetSelectedPredecessorIndices(predecessorsListBox, 0); // countTasks(MoldBuildTreeView, selectedNode.Parent.Text)

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

            component = Project.ComponentList.Find(x => x.Name == selectedNode.Parent.Text);
            task = component.TaskList.ElementAt(selectedNode.Index);

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
                machineComboBox.SelectedItem,
                personnelComboBox.SelectedItem,
                predecessorString,
                taskNotesTextBox.Text
            );

            SelectNextTask();
        }

        private void SetObjectTaskInfo(string component, string taskName)
        {



        }

        private void LoadTaskListToTree(List<TaskInfo> tiList)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            int componentCount = 0, taskCount = 0;

            foreach (TreeNode node in MoldBuildTreeView.Nodes[0].Nodes)
            {
                componentCount++;
            }

            foreach (TaskInfo task in tiList)
            {
                if(task.Level == 1) // TaskInfo item is a component;
                {
                    MoldBuildTreeView.Nodes[0].Nodes.Add(task.Text);
                    componentCount++;
                    taskCount = 0;
                }
                else if(task.Level == 2) // TaskInfo item is a task.
                {
                    MoldBuildTreeView.Nodes[0].Nodes[componentCount - 1].Nodes.Add(task.Text);
                    taskCount++;
                }
                else if(task.Level == 3) // TaskInfo item is task info.
                {
                    MoldBuildTreeView.Nodes[0].Nodes[componentCount - 1].Nodes[taskCount - 1].Nodes.Add(task.Text);
                }
            }
        }
        
        private void LoadProjectToForm(ProjectInfo project)
        {
            TreeNode currentComponentNode, currentTaskNode;

            PrintObjectTree();

            if(project.HasProjectInfo)
            {
                MoldBuildTreeView.Nodes[0].Text = project.JobNumber;
                ProjectNumberTextBox.Text = project.ProjectNumber.ToString();
                dueDateTimePicker.Value = project.DueDate;
                ToolMakerComboBox.SelectedText = project.ToolMaker;
                DesignerComboBox.Text = project.Designer;
                RoughProgrammerComboBox.Text = project.RoughProgrammer;
                ElectrodeProgrammerComboBox.Text = project.ElectrodeProgrammer;
                FinishProgrammerComboBox.Text = project.FinishProgrammer;
            }

            foreach (ClassLibrary.Component component in project.ComponentList)
            {
                currentComponentNode = MoldBuildTreeView.Nodes[0].Nodes.Add(component.Name);

                foreach (TaskInfo task in component.TaskList)
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

        private ProjectInfo ConvertQuoteToProject(ProjectInfo project)
        {
            // Need to check if form already contains project data.
            if(project.ComponentList.Count > 0)
            {
                MessageBox.Show("Can't add a quote to a work project tree with data in it.");
                return project;
            }

            project.JobNumber = project.QuoteInfo.Customer + "_" + project.QuoteInfo.PartName; // What to do when these two pieces of information are missing?
            project.SetProjectDueDate(DateTime.Today);
            project.HasProjectInfo = true;
            project.AddComponent("Quote");

            // Task list is automatically generated inside the QuoteInfo class when quote is read.
            project.ComponentList.First().AddTaskList(project.QuoteInfo.TaskList);

            return project;
        }

        private void LoadQuotedProjectToForm(ProjectInfo project)
        {
            TreeNode quoteNode, currentTaskNode;
            MoldBuildTreeView.Nodes[0].Text = project.QuoteInfo.Customer + "_" + project.QuoteInfo.PartName; // What to do when these two pieces of information are missing?
            quoteNode = MoldBuildTreeView.Nodes[0].Nodes.Add("Quote");

            foreach (TaskInfo task in project.QuoteInfo.TaskList)
            {
                currentTaskNode = quoteNode.Nodes.Add(task.TaskName);
                currentTaskNode.Nodes.Add(task.Hours + " Hour(s)");
                currentTaskNode.Nodes.Add(task.Duration);
                currentTaskNode.Nodes.Add("");
                currentTaskNode.Nodes.Add("");
                currentTaskNode.Nodes.Add("");
                currentTaskNode.Nodes.Add("");
            }

            MoldBuildTreeView.Nodes[0].Expand();
            quoteNode.Expand();
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
            if(value == "")
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
                machineList = new List<string> { "", "Mazak", "Haas", "Old 640 Yasda"};
            }
            else if(taskName == "CNC Finish")
            {
                machineList = new List<string> { "", "430 Yasda", "New 640 Yasda", "950 Yasda" };
            }
            else if(taskName == "CNC Electrodes")
            {
                machineList = new List<string> { "", "Sodick", "Makino" };
            }
            else
            {
                machineList = new List<string> { "" };
            }

            return machineList;
        }

        private List<string> GetPersonnelList (string taskName)
        {
            List<string> personnelList;

            if (taskName == "Program Rough")
            {
                personnelList = new List<string> { "", "Alex Anderson", "Shawn Swiggum", "Derek Timm", "Micah Bruns" };
            }
            else if (taskName == "Program Finish")
            {
                personnelList = new List<string> { "", "Alex Anderson", "Shawn Swiggum", "Rod Schilts" };
            }
            else if (taskName == "Program Electrodes")
            {
                personnelList = new List<string> { "", "Alex Anderson", "Shawn Swiggum", "Rod Schilts", "Josh Meservey" };
            }
            else if (taskName == "CNC Rough")
            {
                personnelList = new List<string> { "", "Derek Timm", "Micah Bruns", "Ed Mendez", "Jon Gruntner" };
            }
            else if (taskName == "CNC Finish")
            {
                personnelList = new List<string> { "", "Derek Timm", "Micah Bruns", "Ed Mendez", "Jon Gruntner" };
            }
            else if (taskName == "CNC Electrodes")
            {
                personnelList = new List<string> { "", "Mark Rasmussen", "Rod Shilts" };
            }
            else
            {
                personnelList = new List<string> { "" };
            }

            return personnelList;
        }

        private string GetPresetPersonnel(string taskName)
        {
            if(taskName == "Program Rough")
            {
                return RoughProgrammerComboBox.Text;
            }
            else if(taskName == "Program Finish")
            {
                return FinishProgrammerComboBox.Text;
            }
            else if(taskName == "Program Electrodes")
            {
                return ElectrodeProgrammerComboBox.Text;
            }
            else
            {
                return "";
            }
        }

        private string GetPresetHours(string taskName)
        {
            return "";
        }

        private TaskInfo GetPresets(string taskName)
        {
            TaskInfo ti;

            if(taskName == "Program Rough")
            {
                ti = new TaskInfo(RoughProgrammerComboBox.Text, "0", "1");
            }
            else if(taskName == "Program Finish")
            {
                ti = new TaskInfo(FinishProgrammerComboBox.Text, "0", "1");
            }
            else if (taskName == "Program Electrodes")
            {
                ti = new TaskInfo(ElectrodeProgrammerComboBox.Text, "0", "1");
            }
            else if (taskName == "CNC Rough")
            {
                ti = new TaskInfo("", "0", "1");
            }
            else if (taskName == "CNC Finish")
            {
                ti = new TaskInfo("", "0", "1");
            }
            else if (taskName == "CNC Electrodes")
            {
                ti = new TaskInfo("", "0", "1");
            }
            else if (taskName == "Heat Treat")
            {
                ti = new TaskInfo("", "0", "5");
            }
            else if (taskName.Contains("Inspection"))
            {
                ti = new TaskInfo("", "1", "0");
            }
            else if (taskName.Contains("Grind"))
            {
                ti = new TaskInfo("", "1", "0");
            }
            else if (taskName.Contains("EDM Sinker"))
            {
                ti = new TaskInfo("", "0", "1");
            }
            else if (taskName.Contains("EDM Wire"))
            {
                ti = new TaskInfo("", "0", "1");
            }
            else if (taskName == "Polish")
            {
                ti = new TaskInfo("", "0", "8");
            }
            else if (taskName == "Texturing")
            {
                ti = new TaskInfo("", "0", "5");
            }
            else
            {
                ti = new TaskInfo("", "0", "0"); ;
            }

            return ti;
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

        private void SetSelectedPredecessors(ListBox listBox, string taskName)
        {
            if(taskName == "CNC Rough")
            {

            }
        }

        private string GetSelectedPredecessorText(ListBox listBox)
        {
            StringBuilder predecessorString = new StringBuilder();

            foreach(string item in listBox.SelectedItems)
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

        private string GetSelectedPredecessorIndices(ListBox listBox, int baseCount)
        {
            StringBuilder predecessorString = new StringBuilder();
            int index;

            foreach (int n in listBox.SelectedIndices)
            {
                if (predecessorString.Length == 0)
                {
                    index = n + baseCount + 1;
                    predecessorString.Append(index);
                }
                else
                {
                    index = n + baseCount + 1;
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
            foreach(TreeNode item in task.Parent.Nodes)
            {
                if(item.Text == node)
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
                    //list.Add("Design / Make Drawings");
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
                    if(NodeExists(task, "Inspection Post CNC Electrodes"))
                    {
                        list.Add("Inspection Post CNC Electrodes");
                    }
                    else if(NodeExists(task, "CNC Electrodes"))
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
            else if(selectedNode.Level == 1)
            {
                Project.RemoveComponent(selectedNode.Text);
            }
            else if(selectedNode.Level == 2)
            {
                var component = Project.ComponentList.Find(x => x.Name == selectedNode.Parent.Text);
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
                        var component = Project.ComponentList.Where(x => x.Name == node.Parent.Text).First();
                        component.MoveTaskUp(node.Index);
                    }

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
                        var component = Project.ComponentList.Where(x => x.Name == node.Parent.Text).First();
                        component.MoveTaskDown(node.Index);
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
            if(MoldBuildTreeView.SelectedNode.Level == 2)
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
				excelApp.Workbooks.Open(@"X:\TOOLROOM\FORMS\Work Load.xlsm", ReadOnly:true);
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

            foreach(ClassLibrary.Component component in Project.ComponentList)
            {
                Console.WriteLine($"{component.Name}");

                foreach(TaskInfo task in component.TaskList)
                {
                    Console.WriteLine($"    {task.TaskName}");
                    Console.WriteLine($"        {task.Hours}");
                    Console.WriteLine($"        {task.Duration}");
                    Console.WriteLine($"        {task.Machine}");
                    Console.WriteLine($"        {task.Personnel}");
                    Console.WriteLine($"        {task.Predecessors}");
                    Console.WriteLine($"        {task.Notes}");
                }
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
                finishProgrammer: FinishProgrammerComboBox.Text
            );
        }

        private void SelectRelatedTasks(string taskName)
        {
            if(TaskListBox.GetSelected(TaskListBox.Items.IndexOf(taskName)))
            {
                if(taskName.Contains("Rough"))
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
                else if(taskName.Contains("Finish") && !taskName.Contains("Grind"))
                {
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("Program Finish"), true);
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("CNC Finish"), true);
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("Inspection Post CNC Finish"), true);
                }
                else if(taskName.Contains("Polish"))
                {
                    TaskListBox.SetSelected(TaskListBox.Items.IndexOf("Inspection Post Polish"), true);
                }
            }
        }

        private void SelectPredecessors(TreeNode selectedNode)
        {
            List<string> predecessorList = new List<string>();

            predecessorsListBox.ClearSelected();

            if (selectedNode.Nodes.Count > 0  && quoteLoaded == false)
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

            foreach (ClassLibrary.Component component in Project.ComponentList)
            {
                List<int> predList = new List<int>();
                int n = 1;

                foreach (TaskInfo task in component.TaskList)
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

            foreach (ClassLibrary.Component component in Project.ComponentList)
            {
                List<int> predList = new List<int>();
                int n = 1;

                foreach (TaskInfo task in component.TaskList)
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
            durationNumericUpDown.ValueChanged += new System.EventHandler(durationNumericUpDown_ValueChanged);
            durationUnitsComboBox.TextChanged += new System.EventHandler(durationUnitsComboBox_TextChanged);
            matchHoursCheckBox.CheckStateChanged += new System.EventHandler(matchHoursCheckBox_CheckStateChanged);
            machineComboBox.TextChanged += new System.EventHandler(machineComboBox_TextChanged);
            personnelComboBox.TextChanged += new System.EventHandler(personnelComboBox_TextChanged);
            predecessorsListBox.SelectedIndexChanged += new System.EventHandler(predecessorsListBox_SelectedIndexChanged);
            taskNotesTextBox.TextChanged += new System.EventHandler(taskNotesTextBox_TextChanged);
        }

        private void DeactivateTaskHandlers()
        {
            // TaskInfo controls.
            hoursNumericUpDown.ValueChanged -= new System.EventHandler(hoursNumericUpDown_ValueChanged);
            durationNumericUpDown.ValueChanged -= new System.EventHandler(durationNumericUpDown_ValueChanged);
            durationUnitsComboBox.TextChanged -= new System.EventHandler(durationUnitsComboBox_TextChanged);
            matchHoursCheckBox.CheckStateChanged -= new System.EventHandler(matchHoursCheckBox_CheckStateChanged);
            machineComboBox.TextChanged -= new System.EventHandler(machineComboBox_TextChanged);
            personnelComboBox.TextChanged -= new System.EventHandler(personnelComboBox_TextChanged);
            predecessorsListBox.SelectedIndexChanged -= new System.EventHandler(predecessorsListBox_SelectedIndexChanged);
            taskNotesTextBox.TextChanged -= new System.EventHandler(taskNotesTextBox_TextChanged);
        }

        private void ActivateComponentHandlers()
        {
            // ComponentInfo controls.
            quantityNumericUpDown.ValueChanged += new System.EventHandler(quantityNumericUpDown_ValueChanged);
            sparesNumericUpDown.ValueChanged += new System.EventHandler(sparesNumericUpDown_ValueChanged);
            materialComboBox.SelectedIndexChanged += new System.EventHandler(materialComboBox_SelectedIndexChanged);
            componentNotesTextBox.TextChanged += new System.EventHandler(componentNotesTextBox_TextChanged);
        }

        private void DeactivateComponentHandlers()
        {
            // ComponentInfo controls.
            quantityNumericUpDown.ValueChanged -= new System.EventHandler(quantityNumericUpDown_ValueChanged);
            sparesNumericUpDown.ValueChanged -= new System.EventHandler(sparesNumericUpDown_ValueChanged);
            materialComboBox.SelectedIndexChanged -= new System.EventHandler(materialComboBox_SelectedIndexChanged);
            componentNotesTextBox.TextChanged -= new System.EventHandler(componentNotesTextBox_TextChanged);
        }

        private void RenameButton_Click(object sender, EventArgs e)
        {
            string input = Interaction.InputBox("Enter a new name:", "Change Name", MoldBuildTreeView.SelectedNode.Text, -1, -1);

            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;

            if (selectedNode.Level >= 0 && selectedNode.Level <= 2)
            {
                RenameNode(input);
            }

            if(selectedNode.Level == 2)
            {
                predecessorsListBox.SelectedIndexChanged -= new System.EventHandler(predecessorsListBox_SelectedIndexChanged);

                predecessorsListBox.DataSource = GetPredecessorList(selectedNode.Parent);

                predecessorsListBox.ClearSelected();

                SelectPredecessors(selectedNode);

                predecessorsListBox.SelectedIndexChanged += new System.EventHandler(predecessorsListBox_SelectedIndexChanged);
            }
        }

        private void UpButton_Click(object sender, EventArgs e)
        {
            MoveSelectedNodeUp(MoldBuildTreeView.SelectedNode);
        }

        private void DownButton_Click(object sender, EventArgs e)
        {
            MoveSelectedNodeDown(MoldBuildTreeView.SelectedNode);
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

        private void Project_Creation_Form_FormClosed(object sender, FormClosedEventArgs e)
		{
			if(excelApp != null)
			{
				excelApp.Quit();
				GC.Collect();
				GC.WaitForPendingFinalizers();
				Marshal.ReleaseComObject(excelApp);
			}
		}

        private void ComponentListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComponentTextBox.Text = prefix + ComponentListBox.Text;
        }

        private void TaskListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(TaskListBox.SelectedItem.ToString());
        }

        private void LookupDataButton_Click(object sender, EventArgs e)
        {
            //openWorkloadSheetExcelCOM();
            OpenWorkloadSheet();
        }

        private void ToolMakerComboBox_DropDown(object sender, EventArgs e)
        {
            PopulateComboBox((ComboBox)sender);
        }

        private void DesignerComboBox_DropDown(object sender, EventArgs e)
        {
            PopulateComboBox((ComboBox)sender);
        }

        private void RoughProgrammerComboBox_DropDown(object sender, EventArgs e)
        {
            PopulateComboBox((ComboBox)sender);
        }

        private void FinishProgrammerComboBox_DropDown(object sender, EventArgs e)
        {
            PopulateComboBox((ComboBox)sender);
        }

        private void ElectrodeProgrammerComboBox_DropDown(object sender, EventArgs e)
        {
            PopulateComboBox((ComboBox)sender);
        }

        private void Project_Creation_Form_Load(object sender, EventArgs e)
        {
            if(Project.HasProjectInfo)
            {
                this.Text = "Edit Project";
                LoadProjectToForm(Project);
                this.CreateProjectButton.Text = "Change";
            }

            prefix = "A-";
            MoldBuildTreeView.Nodes[0].Expand();
        }

        private void hoursNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            //AddTaskInfoToSelectedTask("Hours", hoursNumericUpDown.Value.ToString() + " Hour(s)");
            UpdateDuration();
            updateInfoButton.BackColor = Color.Orange;
        }

        private void durationNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            //AddTaskInfoToSelectedTask("Duration", durationNumericUpDown.Value.ToString() + " " + durationUnitsComboBox.Text + " Duration");
            updateInfoButton.BackColor = Color.Orange;
        }

        private void machineComboBox_TextChanged(object sender, EventArgs e)
        {
            //AddTaskInfoToSelectedTask("Machine", machineComboBox.Text);
            updateInfoButton.BackColor = Color.Orange;
        }

        private void personnelComboBox_TextChanged(object sender, EventArgs e)
        {
            //AddTaskInfoToSelectedTask("Personnel", personnelComboBox.Text);
            updateInfoButton.BackColor = Color.Orange;
        }

        private void predecessorsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            foreach (int index in predecessorsListBox.SelectedIndices)
            {
                if(selectedNode.Index == index)
                {
                    MessageBox.Show("A task cannot be its own predecessor.");
                    //Console.WriteLine(Project.ComponentList[selectedNode.Parent.Index].TaskList[selectedNode.Index].Predecessors);

                    SelectPredecessors(selectedNode);
                }
            }
            
            updateInfoButton.BackColor = Color.Orange;
        }

        private void taskNotesTextBox_TextChanged(object sender, EventArgs e)
        {
            updateInfoButton.BackColor = Color.Orange;
        }

        private void durationUnitsComboBox_TextChanged(object sender, EventArgs e)
        {
            //AddTaskInfoToSelectedTask("Duration", durationNumericUpDown.Value.ToString() + " " + durationUnitsComboBox.Text + " Duration");
            updateInfoButton.BackColor = Color.Orange;
        }

        private void MoldBuildTreeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            // Maybe use to look up or open up something.
        }

        private void MoldBuildTreeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {

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
            ClassLibrary.Component component;
            string[] str1Arr, str2Arr;
            List<string> predecessorList = new List<string>();
            TaskInfo ti;

            if (selectedNode.Level == 0 && formLoad == false)
            {
                tabControl1.SelectedTab = tabPage2;
            }
            else if (selectedNode.Level == 1)
            {
                DeactivateComponentHandlers();

                component = Project.ComponentList.Find(x => x.Name == selectedNode.Text);
                quantityNumericUpDown.Value = component.Quantity;
                sparesNumericUpDown.Value = component.Spares;
                materialComboBox.Text = component.Material;

                componentPictureBox.Image = component.Picture;

                //if(component.Picture.Count > 0)
                //{
                //    componentPictureBox.Image = component.Picture[0];
                //}
                //else
                //{
                //    componentPictureBox.Image = null;
                //}
                
                componentNotesTextBox.Text = component.Notes;

                tabControl1.SelectedTab = tabPage3;

                ActivateComponentHandlers();
            }
            else if (selectedNode.Level == 2)
            {
                try
                {
                    DeactivateTaskHandlers();

                    tabControl1.SelectedTab = tabPage4;

                    machineComboBox.DataSource = GetMachineList(selectedNode.Text);
                    personnelComboBox.DataSource = GetPersonnelList(selectedNode.Text);
                    predecessorsListBox.DataSource = GetPredecessorList(selectedNode.Parent);

                    predecessorsListBox.ClearSelected();

                    ti = GetPresets(selectedNode.Text);

                    if (selectedNode.Nodes.Count > 0)
                    {
                        str1Arr = selectedNode.Nodes[0].Text.Split(' ');
                        str2Arr = selectedNode.Nodes[1].Text.Split(' ');

                        if(int.TryParse(str1Arr[0], out int hours))
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
                        hoursNumericUpDown.Value = ti.Hours;
                        durationNumericUpDown.Value = Convert.ToDecimal(ti.Duration);
                        machineComboBox.SelectedText = ti.Machine;
                        personnelComboBox.SelectedText = ti.Resource;

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
                    MessageBox.Show(er.Message);
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

        private void updateInfoButton_Click(object sender, EventArgs e)
        {
            SetTaskInfoForSelectedTask();
            updateInfoButton.BackColor = Color.Transparent;
        }

        private void matchHoursCheckBox_CheckStateChanged(object sender, EventArgs e)
        {
            UpdateDuration();
        }

        private void ProjectNumberTextBox_TextChanged(object sender, EventArgs e)
        {
            ProjectNumberTextBox.BackColor = Color.White;

            if(CreateProjectButton.Text == "Change")
            {
                if(ProjectNumberTextBox.Text == "0")
                {
                    MessageBox.Show("Project number cannot be 0.");
                    ProjectNumberTextBox.Text = Project.ProjectNumber.ToString();
                }
                else
                {
                    Project.SetProjectNumber(ProjectNumberTextBox.Text);
                }
            }
        }

        private void ToolMakerComboBox_TextChanged(object sender, EventArgs e)
        {
            ToolMakerComboBox.BackColor = Color.White;
        }

        private void AddTasksButton_Click(object sender, EventArgs e)
        {
            AddSelectedTasksToSelectedComponent();
        }

        private void AddComponentButton_Click(object sender, EventArgs e)
        {
            AddComponentToTree(ComponentTextBox.Text);
        }

        private void Project_Creation_Form_Shown(object sender, EventArgs e)
        {
            //MessageBox.Show("Shown");
            formLoad = false;
        }

        private void saveTemplateButton_Click(object sender, EventArgs e)
        {
            int projectNumberResult;

            if(MoldBuildTreeView.Nodes[0].Text == "Tool Number*")
            {
                MessageBox.Show("Please enter a tool number.");
                MoldBuildTreeView.Nodes[0].BackColor = Color.Red;
                MoldBuildTreeView.SelectedNode = MoldBuildTreeView.Nodes[0];
                MoldBuildTreeView.Focus();
                return;
            }

            if(ProjectNumberTextBox.Text == "")
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

            string fileName;
            Templates tmpt = new Templates();
            SetProjectInfo();
            fileName = tmpt.SaveTemplateFile(MoldBuildTreeView.Nodes[0].Text + " - #" + projectNumberResult);

            if(fileName != "")
            {
                tmpt.WriteProjectToTextFile(Project, fileName);
            }
        }

        private void TaskListBox_MouseClick(object sender, MouseEventArgs e)
        {
            string selectedItemName = TaskListBox.Items[TaskListBox.IndexFromPoint(e.Location)].ToString();
            SelectRelatedTasks(selectedItemName);
            //MessageBox.Show(selectedItemName);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //checkForTasksWithNoSuccessors();

            ExcelInteractions ei = new ExcelInteractions();
            string filename;

            OpenFileDialog snapshotOpenFileDialog = new OpenFileDialog
            {
                InitialDirectory = @"C:\Users\" + Environment.UserName + @"\Downloads",
                Filter = "Excel Files (*.xlsm, *.xlsx)|*.xlsm;*.xlsx"
            };

            Nullable<bool> result = Convert.ToBoolean(snapshotOpenFileDialog.ShowDialog());

            if(result == true)
            {
                filename = snapshotOpenFileDialog?.FileName;

                if (filename == "")
                {
                    return;
                }
                
                Project.SetQuoteInfo(ei.GetQuoteInfo(filename));

                LoadProjectToForm(ConvertQuoteToProject(Project));
                quoteLoaded = true;
            }
            else
            {
                return;
            }

            Console.WriteLine($"{Project.QuoteInfo.ProgramRoughHours} {Project.QuoteInfo.ProgramFinishHours} {Project.QuoteInfo.ProgramElectrodeHours} {Project.QuoteInfo.CNCRoughHours} {Project.QuoteInfo.CNCFinishHours} {Project.QuoteInfo.CNCElectrodeHours} {Project.QuoteInfo.EDMSinkerHours}");
        }

        private void quantityNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            ClassLibrary.Component selectedComponent;

            if (selectedNode.Level == 1)
            {
                selectedComponent = Project.ComponentList.Find(x => x.Name == selectedNode.Text);
                selectedComponent.SetQuantity(Convert.ToInt16(quantityNumericUpDown.Value));
            }
        }

        private void sparesNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            ClassLibrary.Component selectedComponent;

            if (selectedNode.Level == 1)
            {
                selectedComponent = Project.ComponentList.Find(x => x.Name == selectedNode.Text);
                selectedComponent.SetSpares(Convert.ToInt16(sparesNumericUpDown.Value));
            }
        }

        private void materialComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            ClassLibrary.Component selectedComponent;

            if (selectedNode.Level == 1)
            {
                selectedComponent = Project.ComponentList.Find(x => x.Name == selectedNode.Text);
                selectedComponent.SetMaterial(materialComboBox.Text);
            }
        }

        private void componentNotesTextBox_TextChanged(object sender, EventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            ClassLibrary.Component selectedComponent;

            if (selectedNode.Level == 1)
            {
                selectedComponent = Project.ComponentList.Find(x => x.Name == selectedNode.Text);
                selectedComponent.SetNote(componentNotesTextBox.Text);
            }
        }

        private void clipboardButton_Click(object sender, EventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            ClassLibrary.Component selectedComponent;

            if (selectedNode.Level == 1)
            {
                selectedComponent = Project.ComponentList.Find(x => x.Name == selectedNode.Text);
                selectedComponent.SetPicture();
                componentPictureBox.Image = selectedComponent.Picture;
            }
            else
            {
                MessageBox.Show("Please select a ClassLibrary.Component to add a picture to.");
            }
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            TreeNode selectedNode = MoldBuildTreeView.SelectedNode;
            ClassLibrary.Component selectedComponent;

            if (selectedNode.Level == 1)
            {
                if (Clipboard.ContainsImage())
                {
                    selectedComponent = Project.ComponentList.Find(x => x.Name == selectedNode.Text);
                    Clipboard.SetImage(selectedComponent.Picture);
                }
            }
        }

        private void loadTemplateButton_Click(object sender, EventArgs e)
        {
            Templates tmpt = new Templates();
            string fileName = tmpt.OpenTemplateFile();
            Console.WriteLine("Load Template Button Click.");

            if(fileName != "")
            {
                Project = tmpt.ReadProjectFromTextFile(fileName);
                LoadProjectToForm(Project);
                MoldBuildTreeView.Nodes[0].Expand();
                //printObjectTree();
            }
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            RemoveSelectedNodeFromTree();
        }

        private void CreateProjectButton_Click(object sender, EventArgs e)
        {
            DataValidated = true;

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

            if(Project.ComponentList.Count == 0)
            {
               MessageBox.Show("No components entered.");
               return;
            }

            Database db = new Database();

            if(missingTaskInfo == true)
            {
                return;
            }

            SetProjectInfo();

            if(CreateProjectButton.Text == "Create")
            {
                if(db.LoadProjectToDB(Project))
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
            else if(CreateProjectButton.Text == "Change")
            {
                if(db.EditProjectInDB(Project))
                {
                    this.DialogResult = DialogResult.OK;
                }
            }

            //printObjectTree();
        }
    }
}
