using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using Task = Microsoft.Office.Interop.MSProject.Task;
using Project = Microsoft.Office.Interop.MSProject;
using System.Windows.Forms;

namespace Toolroom_Scheduler
{
    class MSProject
    {
		OpenFileDialog projectOpenDialog;
		Project.Application projApp;
        Project.Project proj;
        Task task;

		public string OpenProjectFile()
		{
			string filename;
			projApp = new Project.Application();
			projectOpenDialog = new OpenFileDialog();
			//MessageBox.Show("Make Project file is saved if it is currently open! Then click OK.");
			projectOpenDialog.InitialDirectory = @"C:\Users\joshua.meservey\Desktop\Toolroom Schedules";
			projectOpenDialog.Filter = "MPP Files (*.mpp)|*.mpp";
			Nullable<bool> result = Convert.ToBoolean(projectOpenDialog.ShowDialog());

			filename = projectOpenDialog?.FileName;

			if (projectOpenDialog.FileName != "")
			{
				this.projApp.FileOpenEx(filename);
				this.proj = projApp.ActiveProject;
				//ProjApp.Visible = true;

			}

			return filename;
		}

		public List<TaskInfo> ReadProjectFile()
		{
			int r, count, highestTaskID;
			string component;
			string taskname;
			string[] nameArr;
			string[] nameArr2;
			string duration;
			string summary;
			string predecessor;
			List<TaskInfo> taskInfoList = new List<TaskInfo>();
            Database db = new Database();

            r = 0;

            component = "";
			count = proj.Tasks.Count;

			nameArr = proj.Name.Split('.');
			nameArr2 = nameArr[0].Split('-');

            if(db.ProjectTasksExist(nameArr2[0], nameArr2[1]))
            {
                // Get highest task id for project in database.
                highestTaskID = db.GetHighestProjectTaskID(nameArr2[0], Convert.ToInt32(nameArr2[1]));

                if(highestTaskID != count)
                {
                    r = highestTaskID;
                }
                
                // Iterate to first component after highest task id.
                // Add New Components to taskInfoList and then to datatable then to database.
                // Add new components and tasks to existing Kan Ban Workbook for project.
            }
            else
            {
                //r = 1;
            }

            Console.WriteLine("ReadProjectFile");

			do
			{
				if (projApp.GetCellInfo(2, r).Text != "" && projApp.GetCellInfo(10, r).Text == "No")
				{
					taskname = projApp.GetCellInfo(2, r).Text;
					duration = projApp.GetCellInfo(3, r).Text;
                    predecessor = projApp.GetCellInfo(6, r).Text;
                    summary = projApp.GetCellInfo(10, r).Text;

					//                                 Project Number           Tool Number             Component                      Task ID                                      Task Name                            Duration                       Predecessors                 Resource Name                        Machine                      Hours                                 Tool Maker                Priority                               Date Added             Notes
					taskInfoList.Add(new TaskInfo(Convert.ToInt32(nameArr2[1]), nameArr2[0].ToString(), component, Convert.ToInt16(projApp.GetCellInfo(1, r).Task.ID), projApp.GetCellInfo(2, r).Text.Trim(), projApp.GetCellInfo(3, r).Text, projApp.GetCellInfo(6, r).Text, projApp.GetCellInfo(7, r).Text, projApp.GetCellInfo(8,r).Text, Convert.ToInt16(projApp.GetCellInfo(9, r).Text), nameArr2[2], Convert.ToInt16(projApp.GetCellInfo(11, r).Text),  DateTime.Now, projApp.GetCellInfo(12, r).Text));
					//                  Project Number Tool Number   Component                     Task ID                                 Task Name                               Duration                             Predecessors                Resource Name                        Machine                          Hours                  Tool Maker            Priority                   Date Added             Notes
					Console.WriteLine($"{nameArr2[1]} {nameArr2[0]} {component} {projApp.GetCellInfo(1, r).Task.ID.ToString()} {projApp.GetCellInfo(2, r).Text.Trim()} {projApp.GetCellInfo(3, r).Text} {projApp.GetCellInfo(6, r).Text} {projApp.GetCellInfo(7, r).Text} { projApp.GetCellInfo(8, r).Text } {projApp.GetCellInfo(9, r).Text}  {nameArr2[2]} {projApp.GetCellInfo(11, r).Text} {DateTime.Now} {projApp.GetCellInfo(12, r).Text}");

				}
				else if (projApp.GetCellInfo(10, r).Text == "Yes")
				{
					component = projApp.GetCellInfo(2, r).Text;
				}

				//Console.WriteLine(taskname + " " + duration + " " + summary + " " + predecessor);
				r++;
			} while (r < count);
            //Console.WriteLine(i);

            //CloseProject();

			return taskInfoList;
		}

		public void CloseProject()
		{
            projApp?.Quit();

            if (proj != null)
			{
                
                Marshal.ReleaseComObject(proj); // Elvis sign does not work here for some reason.
			}

			if (projApp != null)
			{
				Marshal.ReleaseComObject(projApp);
			}
		}

		public void CreateMSProjectFile(ProjectInfo pi, List<TaskInfo> tiList)
        {
            try
            {
                projApp = new Project.Application();
                proj = this.projApp.Projects.Add();

                proj.Application.TableEditEx(Name: proj.CurrentTable, TaskTable: true, FieldName: "Name", Width: 25, WrapText: false);
                proj.Application.TableEditEx(Name: proj.CurrentTable, TaskTable: true, FieldName: "Duration", Width: 10, WrapText: false);
                proj.Application.TableEditEx(Name: proj.CurrentTable, TaskTable: true, FieldName: "Start", Width: 15, WrapText: false);
                proj.Application.TableEditEx(Name: proj.CurrentTable, TaskTable: true, FieldName: "Finish", Width: 15, WrapText: false);
                proj.Application.TableEditEx(Name: proj.CurrentTable, TaskTable: true, NewFieldName: "Text1", ColumnPosition: 8, Width: 20, WrapText: false);
                proj.Application.CustomFieldRename(Project.PjCustomField.pjCustomTaskText1, "Machine");
                proj.Application.TableEditEx(Name: proj.CurrentTable, TaskTable: true, NewFieldName: "Text2", ColumnPosition: 9, Width: 10, WrapText: false);
                proj.Application.CustomFieldRename(Project.PjCustomField.pjCustomTaskText2, "Hours");
                proj.Application.TableEditEx(Name: proj.CurrentTable, TaskTable: true, NewFieldName: "Summary", ColumnPosition: 10, Width: 10, WrapText: false);
                proj.Application.TableEditEx(Name: proj.CurrentTable, TaskTable: true, NewFieldName: "Priority", ColumnPosition: 11, Width: 10, WrapText: false);
                proj.Application.TableEditEx(Name: proj.CurrentTable, TaskTable: true, NewFieldName: "Notes", ColumnPosition: 12, Width: 40, WrapText: false);

                proj.Application.TableApply(proj.CurrentTable);

                proj.SaveAs(@"C:\Users\joshua.meservey\Desktop\Toolroom Schedules\" + pi.JobNumber + "-" + pi.ProjectNumber + "-" + pi.ToolMaker + ".mpp");

                AddMoldDesignTask(pi, 1);

                int i = 2;

                tiList = convertPredecessorTextToNumbers(tiList);

                foreach (TaskInfo ti in tiList)
                {
                    AddTask(pi, ti, i);
                    i++;
                }

                projApp.Visible = true;
            }
            catch(Exception e)
            {
                MessageBox.Show("Create Project File:" + e.Message);
            }

        }

        public void AddMoldDesignTask(ProjectInfo pi, int index)
        {
            task = proj.Tasks.Add("Design / Make Drawings", index);
            task.OutlineLevel = 1;
            task.Manual = false;
            task.Duration = "8 days";
			task.ResourceNames = pi.Designer;
        }

        public void AddTask(ProjectInfo pi, TaskInfo ti, int index)
        {
            Console.WriteLine(ti.Component);

            if(ti.IsSummary == true)
            {
                //this.ProjApp
                
                task = proj.Tasks.Add(ti.Component,index);
                task.OutlineLevel = 1;
                task.Manual = false;
                //this.proj.Task(ti.TaskName)
            }
            else if(ti.IsSummary == false)
            {
                task = proj.Tasks.Add(ti.TaskName, index);
                task.OutlineLevel = 2;
                task.Manual = false;
                setTaskInfo(pi, ti, task);
            }
            
        }

        public void FormatProject()
        {

        }

        public void ShowProject()
        {
            this.projApp.Visible = true;
        }

        public void SaveAndCloseProject()
        {
            //this.ProjApp.ScreenUpdating = false;
            this.projApp.FileSave();
            //this.ProjApp.FileExit();
            this.projApp.Quit();
            Marshal.ReleaseComObject(proj);
            Marshal.ReleaseComObject(projApp);
        }

        private void setTaskInfo(ProjectInfo pi, TaskInfo ti, Task t)
        {
            if (t.Name == "Design / Make Drawings")
            {
                t.Duration = "1 day";
				t.ResourceNames = pi.Designer;
			}
            else
            {
                t.Duration = ti.Duration;
                t.Predecessors = ti.Predecessors;
                t.ResourceNames = ti.Resource;
                t.Text1 = ti.Machine;
                t.Text2 = ti.Hours.ToString();
                t.Notes = ti.Notes;
            }
        }

        public List<TaskInfo> convertPredecessorTextToNumbers(List<TaskInfo> list)
        {
            StringBuilder newPredecessorString = new StringBuilder();
            string[] predecessorArr;

            foreach (TaskInfo task in list)
            {
                if(task.IsSummary == false)
                {
                    if(task.Predecessors.Contains(','))
                    {
                        predecessorArr = task.Predecessors.Split(',');
                    }
                    else
                    {
                        predecessorArr = new string[] { task.Predecessors };
                    }
                    
                    var tempComponentTaskList = from preds in list
                                                where preds.Component == task.Component
                                                select preds;

                    foreach(string pred in predecessorArr)
                    {
                        foreach (TaskInfo task2 in tempComponentTaskList)
                        {
                            if(pred == "Design / Make Drawings")
                            {
                                newPredecessorString.Append("1");
                                break;
                            }
                            else if(pred == task2.TaskName)
                            {
                                if(newPredecessorString.Length == 0)
                                {
                                    newPredecessorString.Append(task2.ID);
                                }
                                else
                                {
                                    newPredecessorString.Append("," + task2.ID);
                                }

                                break;
                            }
                        }
                    }

                    task.Predecessors = newPredecessorString.ToString();
                    //Console.WriteLine($" {task.ID} {task.TaskName} {task.Duration} {task.Predecessors} ");
                    newPredecessorString.Clear();
                }
            }

            return list;
        }

        private void iterateThroughProject()
        {
            int r, rTemp, i, index, count, componentStartRow, currentCNCTaskID;
            string predecessorList = "";
            count = proj.Tasks.Count;
            r = 0;
            i = 0;
            componentStartRow = 0;

            do
            {
                if (projApp.GetCellInfo(2, r).Text != "" && projApp.GetCellInfo(10, r).Text == "No")
                {
                    if (projApp.GetCellInfo(2, r).Text == "CNC Rough")
                    {
                        currentCNCTaskID = projApp.GetCellInfo(2, r).Task.ID;
                        do
                        {
                            rTemp = componentStartRow;

                            if(projApp.GetCellInfo(2, rTemp).Text == "Order Steel / Steel Arrival")
                            {
                                predecessorList = projApp.GetCellInfo(2, rTemp).Task.ID.ToString();
                            }
                            else if (projApp.GetCellInfo(2, rTemp).Text == "Program Rough")
                            {
                                predecessorList = predecessorList + projApp.GetCellInfo(2, rTemp).Task.ID.ToString();
                            }
                            rTemp++;

                        } while (rTemp < currentCNCTaskID);
                    }
                    else if (projApp.GetCellInfo(2,r).Text == "CNC Finish" )
                    {

                    }
                    else if (projApp.GetCellInfo(2, r).Text == "CNC Electrodes")
                    {

                    }
                    //        taskname = ProjApp.GetCellInfo(2, r).Text;
                    //duration = ProjApp.GetCellInfo(3, r).Text;
                    //summary = ProjApp.GetCellInfo(6, r).Text;
                    //predecessor = ProjApp.GetCellInfo(9, r).Text;
                    ////                                 Project Number           Tool Number             Component                      Task ID                                      Task Name                            Duration                       Predecessors                 Resource Name                        Hours                                 Tool Maker                Priority                               Date Added             Notes
                    //taskInfoArr[i] = new TaskInfo(Convert.ToInt16(nameArr2[1]), nameArr2[0].ToString(), component, Convert.ToInt16(ProjApp.GetCellInfo(1, r).Task.ID), ProjApp.GetCellInfo(2, r).Text.Trim(), ProjApp.GetCellInfo(3, r).Text, ProjApp.GetCellInfo(9, r).Text, ProjApp.GetCellInfo(10, r).Text, Convert.ToInt16(ProjApp.GetCellInfo(7, r).Text), nameArr2[2], Convert.ToInt16(ProjApp.GetCellInfo(8, r).Text), DateTime.Now, ProjApp.GetCellInfo(11, r).Text);
                    ////                  Project Number Tool Number   Component                     Task ID                                 Task Name                               Duration                             Predecessors                Resource Name                        Hours                  Tool Maker            Prority                   Date Added             Notes
                    //Console.WriteLine($"{nameArr2[1]} {nameArr2[0]} {component} {ProjApp.GetCellInfo(1, r).Task.ID.ToString()} {ProjApp.GetCellInfo(2, r).Text.Trim()} {ProjApp.GetCellInfo(3, r).Text} {ProjApp.GetCellInfo(9, r).Text} {ProjApp.GetCellInfo(10, r).Text}  {ProjApp.GetCellInfo(7, r).Text}  {nameArr2[2]} {ProjApp.GetCellInfo(8, r).Text} {DateTime.Now} {ProjApp.GetCellInfo(11, r).Text}");

                    i++;
                }
                else if (projApp.GetCellInfo(10, r).Text == "Yes")
                {
                    componentStartRow = r + 1;
                }

                //Console.WriteLine(taskname + " " + duration + " " + summary + " " + predecessor);
                r++;
            } while (r < count);
        }
    }
}
