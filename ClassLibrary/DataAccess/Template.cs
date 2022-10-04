using DevExpress.XtraScheduler;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace ClassLibrary
{
    public class Template
    {
        public Template()
        {

        }
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

        public void WriteProjectToTextFile(ProjectModel project, string fileName)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(fileName))
            {
                file.WriteLine($"{project.JobNumber},{project.ProjectNumber},{project.DueDate.ToShortDateString()},{project.ToolMaker},{project.Designer},{project.RoughProgrammer},{project.ElectrodeProgrammer},{project.FinishProgrammer}");

                foreach (ComponentModel component in project.Components)
                {
                    file.WriteLine($"{string.Empty.PadLeft(8)}{component.Component},{component.Quantity},{component.Spares},{component.Material},{component.Finish},{component.Notes},{component.GetPictureString()}");

                    foreach (TaskModel task in component.Tasks)
                    {
                        file.WriteLine($"{string.Empty.PadLeft(16)}{task.TaskName}");

                        if (task.HasInfo)
                        {
                            file.WriteLine($"{string.Empty.PadLeft(24)}{task.Hours} Hour(s)");
                            file.WriteLine($"{string.Empty.PadLeft(24)}{task.Duration}");
                            file.WriteLine($"{string.Empty.PadLeft(24)}{task.Machine}");
                            file.WriteLine($"{string.Empty.PadLeft(24)}{task.Personnel}");
                            file.WriteLine($"{string.Empty.PadLeft(24)}{task.Predecessors}");
                            file.WriteLine($"{string.Empty.PadLeft(24)}{task.Notes}");
                        }
                    }
                }
            }
        }

        public static ProjectModel ReadProjectFromTextFile(string filePath, SchedulerStorage schedulerStorage)
        {
            ProjectModel project = new ProjectModel();
            ComponentModel component = new ComponentModel();
            TaskModel task = new TaskModel();
            List<TaskModel> list = new List<TaskModel>();
            string[] projectInfoArr, componentInfoArr;
            string line, hours, duration, machine, personnel, predecessors, notes;

            project.HasProjectInfo = false;

            if (filePath == "")
            {
                throw new ArgumentException("No file path entered.");
            }
            else if (!System.IO.File.Exists(filePath))
            {
                throw new System.IO.FileNotFoundException("The file name that was entered does not exist.");
            }


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
                else if (count == 8)
                {
                    if (line.Contains(','))
                    {
                        componentInfoArr = line.Split(',');

                        if (componentInfoArr.Count() == 6)
                        {
                            component = new ComponentModel
                            (
                                componentInfoArr[0].Trim(),
                                componentInfoArr[1],
                                componentInfoArr[2],
                                componentInfoArr[3],
                                componentInfoArr[4],
                                componentInfoArr[5]
                            ); 
                        }
                        else if (componentInfoArr.Count() == 7)
                        {
                            component = new ComponentModel
                            (
                                componentInfoArr[0].Trim(),
                                componentInfoArr[1],
                                componentInfoArr[2],
                                componentInfoArr[3],
                                componentInfoArr[4],
                                componentInfoArr[5],
                                componentInfoArr[6]
                            );
                        }

                    }
                    else
                    {
                        component = new ComponentModel(line.Trim());
                    }

                    project.AddComponent(component);
                }
                else if (count == 16)
                {
                    task = new TaskModel(line.Trim(), component.Component);
                    component.AddTask(task);
                }
                else if (count == 24)
                {
                    task.HasInfo = true;

                    hours = line;

                    task.SetHours(hours);

                    duration = file.ReadLine();

                    task.SetDuration(duration);

                    machine = file.ReadLine();

                    task.SetMachine(machine);

                    personnel = file.ReadLine();

                    task.SetPersonnel(personnel);

                    predecessors = file.ReadLine();

                    task.SetPredecessors(predecessors);

                    notes = file.ReadLine();

                    task.SetNotes(notes);

                    task.SetResources(schedulerStorage);
                }

                //System.Console.WriteLine($"{count} {line}");
            }

            file.Close();

            return project;
        }
        /// <summary>
        /// Writes the given object instance to an XML file.
        /// <para>Only Public properties and variables will be written to the file. These can be any type though, even other classes.</para>
        /// <para>If there are public properties/variables that you do not want written to the file, decorate them with the [XmlIgnore] attribute.</para>
        /// <para>Object type must have a parameterless constructor.</para>
        /// </summary>
        /// <typeparam name="T">The type of object being written to the file.</typeparam>
        /// <param name="filePath">The file path to write the object instance to.</param>
        /// <param name="objectToWrite">The object instance to write to the file.</param>
        /// <param name="append">If false the file will be overwritten if it already exists. If true the contents will be appended to the file.</param>
        public static void WriteToXmlFile<T>(string filePath, T objectToWrite, bool append = false) where T : new()
        {
            TextWriter writer = null;
            try
            {
                var serializer = new XmlSerializer(typeof(T));
                writer = new StreamWriter(filePath, append);

                serializer.Serialize(writer, objectToWrite);
            }
            finally
            {
                if (writer != null)
                    writer.Close();
            }
        }

        /// <summary>
        /// Reads an object instance from an XML file.
        /// <para>Object type must have a parameterless constructor.</para>
        /// </summary>
        /// <typeparam name="T">The type of object to read from the file.</typeparam>
        /// <param name="filePath">The file path to read the object instance from.</param>
        /// <returns>Returns a new instance of the object read from the XML file.</returns>
        public static T ReadFromXmlFile<T>(string filePath) where T : new()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(T));

                using (XmlTextReader reader = new XmlTextReader(filePath))
                {
                    return (T)serializer.Deserialize(reader); 
                }
            }
            finally
            {
                //if (reader != null)
                //    reader.Close();
            }
        }
        public static string OpenTemplateFile(string initialDirectory)
        {
            string filename = "";
            OpenFileDialog openTemplateDialog = new OpenFileDialog();
            //MessageBox.Show("Make Project file is saved if it is currently open! Then click OK.");
            openTemplateDialog.InitialDirectory = initialDirectory;
            openTemplateDialog.Filter = "Template files (*.txt, *.xml)|*.txt;*.xml";
            openTemplateDialog.Title = "Load Template";

            if (openTemplateDialog.ShowDialog() == DialogResult.OK)
            {
                //MessageBox.Show("Open");
                filename = openTemplateDialog.FileName.ToString();
            }

            return filename;
        }

        public static string SaveTemplateFile(string fileName, string initialDirectory)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = initialDirectory;
            saveFileDialog.Filter = "XML Files (*.xml)|*.xml"; // Text files (*.txt)|*.txt|
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
