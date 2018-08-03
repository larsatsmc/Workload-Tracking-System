using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using ClassLibrary;

namespace Toolroom_Scheduler
{
    class Templates
    {
        public void WriteExample()
        {
            // These examples assume a "C:\Users\Public\TestFolder" folder on your machine.
            // You can modify the path if necessary.


            // Example #1: Write an array of strings to a file.
            // Create a string array that consists of three lines.
            string[] lines = { "First line", "Second line", "Third line" };
            // WriteAllLines creates a file, writes a collection of strings to the file,
            // and then closes the file.  You do NOT need to call Flush() or Close().
            System.IO.File.WriteAllLines(@"C:\Users\Public\TestFolder\WriteLines.txt", lines);


            // Example #2: Write one string to a text file.
            string text = "A class is the most powerful data type in C#. Like a structure, " +
                           "a class defines the data and behavior of the data type. ";
            // WriteAllText creates a file, writes the specified string to the file,
            // and then closes the file.    You do NOT need to call Flush() or Close().
            System.IO.File.WriteAllText(@"C:\Users\Public\TestFolder\WriteText.txt", text);

            // Example #3: Write only some strings in an array to a file.
            // The using statement automatically flushes AND CLOSES the stream and calls 
            // IDisposable.Dispose on the stream object.
            // NOTE: do not use FileStream for text files because it writes bytes, but StreamWriter
            // encodes the output as text.
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\Users\Public\TestFolder\WriteLines2.txt"))
            {
                foreach (string line in lines)
                {
                    // If the line doesn't contain the word 'Second', write the line to the file.
                    if (!line.Contains("Second"))
                    {
                        file.WriteLine(line);
                    }
                }
            }

            // Example #4: Append new text to an existing file.
            // The using statement automatically flushes AND CLOSES the stream and calls 
            // IDisposable.Dispose on the stream object.
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\Users\Public\TestFolder\WriteLines2.txt", true))
            {
                file.WriteLine("Fourth line");
            }
        }

        public void ReadExample()
        {
            int counter = 0;
            string line;

            // Read the file and display it line by line.  
            System.IO.StreamReader file =
                new System.IO.StreamReader(@"c:\test.txt");
            while ((line = file.ReadLine()) != null)
            {
                System.Console.WriteLine(line);
                counter++;
            }

            file.Close();
            System.Console.WriteLine("There were {0} lines.", counter);
            // Suspend the screen.  
            System.Console.ReadLine();
        }

        public void WriteProjectToTextFile(List<TaskInfo> list)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Users\Public\TestFolder\170255.txt"))
            {
                foreach (TaskInfo task in list)
                {
                    if(task.IsSummary)
                    {
                        file.WriteLine(task.Component);
                    }
                    else if(!task.IsSummary)
                    {
                        file.WriteLine($"        {task.TaskName}");
                        file.WriteLine($"                {task.Hours} Hour(s)");
                        file.WriteLine($"                {task.Duration} Day(s)");
                        file.WriteLine($"                {task.Machine}");
                        file.WriteLine($"                {task.Resource}");
                        file.WriteLine($"                {task.Predecessors}");
                        file.WriteLine($"                {task.Notes}");
                    }
                    else
                    {

                    }
                }
            }
        }

        public void WriteProjectToTextFile(List<TaskInfo> list, string fileName)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(fileName))
            {
                foreach (TaskInfo task in list)
                {
                    if (task.IsSummary)
                    {
                        file.WriteLine(task.Component);
                    }
                    else if (!task.IsSummary && task.Duration != null)
                    {
                        file.WriteLine($"        {task.TaskName}");
                        file.WriteLine($"                {task.Hours} Hour(s)");
                        file.WriteLine($"                {task.Duration}");
                        file.WriteLine($"                {task.Machine}");
                        file.WriteLine($"                {task.Resource}");
                        file.WriteLine($"                {task.Predecessors}");
                        file.WriteLine($"                {task.Notes}");
                    }
                    else if (!task.IsSummary && task.Duration == null)
                    {
                        file.WriteLine($"        {task.TaskName}");
                    }
                }
            }
        }

        public void WriteProjectToTextFile(ProjectInfo project, string fileName)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(fileName))
            {
                file.WriteLine($"{project.JobNumber},{project.ProjectNumber},{project.DueDate.ToShortDateString()},{project.ToolMaker},{project.Designer},{project.RoughProgrammer},{project.ElectrodeProgrammer},{project.FinishProgrammer}");

                foreach (Component component in project.ComponentList)
                {
                    file.WriteLine(component.Name);

                    foreach (TaskInfo task in component.TaskList)
                    {
                        file.WriteLine($"        {task.TaskName}");

                        if(task.HasInfo)
                        {
                            file.WriteLine($"                {task.Hours} Hour(s)");
                            file.WriteLine($"                {task.Duration}");
                            file.WriteLine($"                {task.Machine}");
                            file.WriteLine($"                {task.Resource}");
                            file.WriteLine($"                {task.Predecessors}");
                            file.WriteLine($"                {task.Notes}");
                        }
                    }
                }
            }
        }

        public void ReadProjectFromTextFile()
        {
            string line;

            System.IO.StreamReader file = new System.IO.StreamReader(@"C:\Users\Public\TestFolder\170255.txt");
            while ((line = file.ReadLine()) != null)
            {
                int count = line.TakeWhile(Char.IsWhiteSpace).Count();
                System.Console.WriteLine($"{count} {line}");
            }

            file.Close();
        }

        public List<TaskInfo> ReadTasksFromTextFile(string filePath)
        {
            List<TaskInfo> list = new List<TaskInfo>();
            string line;

            System.IO.StreamReader file = new System.IO.StreamReader(filePath);
            while ((line = file.ReadLine()) != null)
            {
                int count = line.TakeWhile(Char.IsWhiteSpace).Count();

                if(count == 0)
                {
                    list.Add(new TaskInfo(line, 1));
                }
                else if(count == 8)
                {
                    list.Add(new TaskInfo(line.Trim(), 2));
                }
                else if(count == 16)
                {
                    list.Add(new TaskInfo(line.Trim(), 3));
                }

                System.Console.WriteLine($"{count} {line}");
            }

            file.Close();

            return list;
        }

        public ProjectInfo ReadProjectFromTextFile(string filePath)
        {
            ProjectInfo project = new ProjectInfo();
            Component component = new Component();
            TaskInfo task = new TaskInfo();
            List<TaskInfo> list = new List<TaskInfo>();
            string[] projectInfoArr;
            string line;

            project.HasProjectInfo = false;

            System.IO.StreamReader file = new System.IO.StreamReader(filePath);

            while ((line = file.ReadLine()) != null)
            {
                int count = line.TakeWhile(Char.IsWhiteSpace).Count();

                if (count == 0 && line.Contains(','))
                {
                    project.HasProjectInfo = true;
                    projectInfoArr = line.Split(',');

                    project.SetProjectInfo
                    (
                        projectInfoArr[0],
                        projectInfoArr[1],
                        projectInfoArr[2],
                        projectInfoArr[3],
                        projectInfoArr[4],
                        projectInfoArr[5],
                        projectInfoArr[6],
                        projectInfoArr[7]
                    );
                }
                else if (count == 0)
                {
                    component = new Component(line.Trim());
                    project.AddComponent(component);
                }
                else if (count == 8)
                {
                    task = new TaskInfo(line.Trim(), component.Name);
                    component.AddTask(task);
                }
                else if (count == 16)
                {
                    task.HasInfo = true;

                    task.SetHours(line);

                    line = file.ReadLine();

                    task.SetDuration(line);

                    line = file.ReadLine();

                    task.SetMachine(line);

                    line = file.ReadLine();

                    task.SetPersonnel(line);

                    line = file.ReadLine();

                    task.SetPredecessors(line);

                    line = file.ReadLine();

                    task.SetNotes(line);
                }

                //System.Console.WriteLine($"{count} {line}");
            }

            file.Close();

            return project;
        }

        public string OpenTemplateFile()
        {
            string filename = "";
            OpenFileDialog openTemplateDialog = new OpenFileDialog();
            //MessageBox.Show("Make Project file is saved if it is currently open! Then click OK.");
            openTemplateDialog.InitialDirectory = @"X:\TOOLROOM\Workload Tracking System\Templates";
            openTemplateDialog.Filter = "Text Files (*.txt)|*.txt";
            openTemplateDialog.Title = "Load Template";

            if(openTemplateDialog.ShowDialog() == DialogResult.OK)
            {
                //MessageBox.Show("Open");
                filename = openTemplateDialog.FileName.ToString();
            }

            return filename;
        }
 
        public string SaveTemplateFile(string fileName)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = @"X:\TOOLROOM\Workload Tracking System\Templates";
            saveFileDialog.Filter = "Text files (*.txt)|*.txt";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = false;
            saveFileDialog.FileName = fileName;
            saveFileDialog.Title = "Save Template";

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

        public string SaveProjectTemplateFile(string fileName)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = @"X:\TOOLROOM\Workload Tracking System\Templates\Created Projects";
            saveFileDialog.Filter = "Text files (*.txt)|*.txt";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = false;
            saveFileDialog.FileName = fileName;
            saveFileDialog.Title = "Save Template";

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
    }
}
