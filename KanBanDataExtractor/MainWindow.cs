using ClassLibrary;
using ClosedXML.Excel;
using DevExpress.XtraEditors;
using DevExpress.XtraLayout;
using System;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace KanBanDataExtractor
{
    public partial class MainWindow : Form
    {
        BindingList<TaskRecord> TaskList = new BindingList<TaskRecord>();
        public MainWindow()
        {
            InitializeComponent();
            gridControl1.DataSource = TaskList;
        }
        private DateTime? GetDateValue(string date)
        {
            if (DateTime.TryParse(date, out DateTime result))
            {
                return result;
            }
            else
            {
                return null;
            }
        }
        public class TaskRecord
        {
            public int ProjectNumber { get; set; }
            public string JobNumber { get; set; }
            public string Component { get; set; }
            public string Material { get; set; }
            public int Quantity { get; set; }
            public int TaskID { get; set; }
            public string TaskName { get; set; }
            public string Duration { get; set; }
            public DateTime? StartDate { get; set; }
            public DateTime? FinishDate { get; set; }
            public int Hours { get; set; }
            public string Notes { get; set; }
            public string Initials { get; set; }
            public string DateCompleted { get; set; }
        }
        public class ProjectDataControl : XtraUserControl
        {
            TextEdit jobNumberTextEdit;
            TextEdit projectNumberTextEdit;
            public string JobNumber { get { return jobNumberTextEdit.Text; } }
            public string ProjectNumber { get { return projectNumberTextEdit.Text; } }

            public ProjectDataControl()
            {
                LayoutControl layoutControl = new LayoutControl();
                layoutControl.Dock = DockStyle.Fill;
                this.jobNumberTextEdit = new TextEdit();
                this.projectNumberTextEdit = new TextEdit();
                layoutControl.AddItem("Job #", jobNumberTextEdit);
                layoutControl.AddItem("Project #", projectNumberTextEdit);
                this.Controls.Add(layoutControl);
                this.Height = 100;
                this.Width = 200;
                this.Dock = DockStyle.Top;
            }
        }
        private void loadButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            ProjectModel project = new ProjectModel();
            ComponentModel component;
            TaskModel task;
            TaskRecord taskRecord;
            int row;
            Regex regex = new Regex(@"^(\d{6})- Proj #(\d{5,6}).*$");
            Regex regex2 = new Regex(@"^Sheet\d{1,2}");

            try
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(openFileDialog.FileName))
                    {

                        string fileName = openFileDialog.FileName;

                        Console.WriteLine(fileName);
                        Console.WriteLine(Path.GetFileName(fileName));

                        if (regex.IsMatch(Path.GetFileName(fileName)))
                        {
                            var match = regex.Match(Path.GetFileName(fileName));
                            project.JobNumber = match.Groups[1].Value;
                            project.ProjectNumber = int.Parse(match.Groups[2].Value);
                        }
                        else
                        {
                            MessageBox.Show("Could not find Project data in file name.\n\r\n\rPlease enter a Job Number and Project / MWO Number.");
                            ProjectDataControl projectDataControl = new ProjectDataControl();

                            if (XtraDialog.Show(projectDataControl, "Set Project Data", MessageBoxButtons.OKCancel) == DialogResult.OK)
                            {
                                project.JobNumber = projectDataControl.JobNumber;
                                if (int.TryParse(projectDataControl.ProjectNumber, out int result))
                                {
                                    project.ProjectNumber = result; 
                                }
                                else
                                {
                                    MessageBox.Show("Please enter a valid MWO or Project Number.");
                                    return;
                                }
                            }
                            else
                            {
                                MessageBox.Show("Kan Ban Load Cancelled.");
                                return;
                            }
                        }


                        using (IXLWorkbook workbook = new XLWorkbook(fileName))
                        {
                            foreach (IXLWorksheet worksheet in workbook.Worksheets)
                            {
                                if (worksheet.Name != "Summary" && worksheet.Name != "Notes" && regex2.IsMatch(worksheet.Name) == false)
                                {
                                    component = new ComponentModel();

                                    component.Component = worksheet.Cell(2, 1).Value.ToString().Split(':')[1].Trim();
                                    component.Material = worksheet.Cell(3, 1).Value.ToString().Split(':')[1].Trim();
                                    component.Quantity = int.Parse(worksheet.Cell(1, 8).Value.ToString().Split(':')[1].Trim());

                                    project.Components.Add(component);

                                    row = 6;

                                    while (int.TryParse(worksheet.Cell(row, 1).Value.ToString(), out int result))
                                    {
                                        task = new TaskModel();
                                        task.JobNumber = component.JobNumber;
                                        task.Component = component.Component;
                                        task.TaskID = int.Parse(worksheet.Cell(row, 1).Value.ToString());
                                        task.TaskName = worksheet.Cell(row, 2).Value.ToString();
                                        task.Duration = worksheet.Cell(row, 3).Value.ToString();
                                        task.StartDate = GetDateValue(worksheet.Cell(row, 4).Value.ToString());
                                        task.FinishDate = GetDateValue(worksheet.Cell(row, 5).Value.ToString());
                                        task.Hours = int.Parse(worksheet.Cell(row, 6).Value.ToString());
                                        task.Notes = worksheet.Cell(row, 7).Value.ToString();
                                        task.Initials = worksheet.Cell(row, 9).Value.ToString();
                                        task.DateCompleted = worksheet.Cell(row, 10).Value.ToString();
                                        component.Tasks.Add(task);

                                        row++;
                                    }
                                }
                            }
                        }

                        foreach (var comp in project.Components)
                        {
                            Console.WriteLine(comp.Component);

                            foreach (TaskModel tsk in comp.Tasks)
                            {
                                taskRecord = new TaskRecord();

                                taskRecord.ProjectNumber = project.ProjectNumber;
                                taskRecord.JobNumber = project.JobNumber;
                                taskRecord.Component = comp.Component;
                                taskRecord.Material = comp.Material;
                                taskRecord.Quantity = comp.Quantity;
                                taskRecord.TaskID = tsk.TaskID;
                                taskRecord.TaskName = tsk.TaskName;
                                taskRecord.Duration = tsk.Duration;
                                taskRecord.StartDate = tsk.StartDate;
                                taskRecord.FinishDate = tsk.FinishDate;
                                taskRecord.Hours = tsk.Hours;
                                taskRecord.Notes = tsk.Notes;
                                taskRecord.Initials = tsk.Initials;
                                taskRecord.DateCompleted = tsk.DateCompleted;

                                TaskList.Add(taskRecord);
                                Console.WriteLine($"{tsk.TaskID} {tsk.TaskName}");
                            }
                        }

                        MessageBox.Show("Kan Ban Loaded!");
                    }
                    else
                    {
                        MessageBox.Show("That file does not exist.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }

        private void exportButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel File (*.xlsx) |*.xlsx";
            saveFileDialog.InitialDirectory = @"C:\Users\" + Environment.UserName + @"\Desktop";
            saveFileDialog.FileName = "Tool Room Tasks " + DateTime.Today.Month + "-" + DateTime.Today.Day + "-" + DateTime.Today.Year;
            saveFileDialog.DefaultExt = "xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                taskRecordView.ExportToXlsx(saveFileDialog.FileName);
            }
        }
    }
}
