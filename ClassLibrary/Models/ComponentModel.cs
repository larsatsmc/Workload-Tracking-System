﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.XtraScheduler;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Xml.Serialization;

namespace ClassLibrary
{
    public class ComponentModel : INotifyPropertyChanged
    {
        public int ID { get; set; } = 0;
        public string JobNumber { get; set; }
        public int ProjectNumber { get; set; }
        public string Component { get; set; }
        [XmlIgnore]
        public string OldName { get; private set; }
        //public Image Picture { get; set; }
        [XmlIgnore]
        public Image picture;
        public byte[] Picture { get { return ImageToByteArray(picture); } set { picture = NullByteArrayCheck(value); } }
        public string Material { get; set; }
        public string Finish { get; set; }
        public string Notes { get; set; }
        public int Priority { get; set; }
        public int Quantity { get; set; }
        public int Spares { get; set; }
        public string Status { get; set; }
        [XmlIgnore]
        public double PercentComplete { get; set; }
        public List<TaskModel> Tasks { get; set; } = new List<TaskModel>();
        //public System.ComponentModel.BindingList<TaskModel> Tasks { get; set; }
        public bool ReloadTaskList { get; set; }
        public int Position { get; set; }
        public int TaskIDCount { get; set; }

        public DateTime? LatestFinishDate
        {
            get { return GetLatestFinishDate(); }
            //set { latestFinishDate = value; }
        }
        public bool AllTasksDated { get; set; }
        //public bool AllTasksDated
        //{
        //    get
        //    {
        //        if (Tasks.Count(x => x.StartDate == null || x.FinishDate == null) > 0)
        //        {
        //            return false;
        //        }
        //        else
        //        {
        //            return true;
        //        }
        //    }
        //}

        public static int ComponentCharacterLimit = 31;

        /// <summary>
        /// Creates instance of a component and sets TaskIDCount property to 0.
        /// </summary> 
        public ComponentModel() // Default contructors do not execute unless they are called.
        {
            TaskIDCount = 0;
        }
        public ComponentModel(ComponentModel component)
        {
            this.ID = component.ID;
            this.JobNumber = component.JobNumber;
            this.ProjectNumber = component.ProjectNumber;
            this.Component = component.Component;
            this.OldName = component.OldName;
            this.Notes = component.Notes;
            this.Priority = component.Priority;
            this.Position = component.Position;
            this.Material = component.Material;
            this.TaskIDCount = component.TaskIDCount;
            this.Quantity = component.Quantity;
            this.Spares = component.Spares;
            this.Picture = component.Picture;
            this.Finish = component.Finish;
            this.Tasks = component.Tasks;
            this.Status = component.Status;
            this.PercentComplete = component.PercentComplete;
        }
        public ComponentModel(ComponentModel component, string name)
        {
            this.JobNumber = component.JobNumber;
            this.ProjectNumber = component.ProjectNumber;
            this.Component = name;
            this.OldName = name;
            this.Notes = component.Notes;
            this.TaskIDCount = component.TaskIDCount;
            this.Quantity = component.Quantity;
            this.Spares = component.Spares;
            this.Material = component.Material;
            this.Finish = component.Finish;
            this.Tasks = AddCopiedTaskList(component.Tasks, name);
            this.Picture = component.Picture;
        }
        /// <summary>
        /// Creates instance of a component and sets properties for template.
        /// </summary> 
        public ComponentModel(string name, string quantity, string spares, string material, string finish, string notes)
        {
            this.Component = name;
            this.Quantity = Convert.ToInt16(quantity);
            this.Spares = Convert.ToInt16(spares);
            this.Material = material;
            this.Finish = finish;
            this.Notes = notes;
            this.Tasks = new List<TaskModel>();
        }
        public ComponentModel(string name, string quantity, string spares, string material, string finish, string notes, string picture)
        {
            this.Component = name;
            this.Quantity = Convert.ToInt16(quantity);
            this.Spares = Convert.ToInt16(spares);
            this.Material = material;
            this.Finish = finish;
            this.Notes = notes;
            if (picture.Trim().Length > 0)
            {
                //this.Picture = ByteArrayToImage(Convert.FromBase64String(picture));
                this.Picture = Convert.FromBase64String(picture);
            }
            this.Tasks = new List<TaskModel>();
        }
        /// <summary>
        /// Creates instance of a component with given name, sets TaskIDCount property to 0, and initializes a list of type TaskInfo.
        /// </summary> 
        public ComponentModel(string name)
        {
            this.TaskIDCount = 0;
            this.Quantity = 1;
            this.Spares = 0;
            this.Notes = "";
            this.Material = "";
            this.Finish = "";
            this.Tasks = new List<TaskModel>();
            this.Component = name;
        }
        /// <summary>
        /// Creates instance of a component with given name, sets TaskIDCount property to 0, and initializes a list of type TaskInfo.
        /// </summary> 
        public ComponentModel(object name)
        {
            this.TaskIDCount = 0;
            this.Quantity = 1;
            this.Spares = 0;
            this.Notes = "";
            this.Material = "";
            this.Finish = "";
            this.Status = "";
            this.Tasks = new List<TaskModel>();
            this.Component = ConvertObjectToString(name);
        }
        /// <summary>
        /// Creates instance of a component with given name, sets TaskIDCount property to 0, and initializes a list of type TaskInfo.
        /// </summary> 
        public ComponentModel(object name, object notes, object priority, object position, object material, object finish, object taskIDCount)
        {
            this.Tasks = new List<TaskModel>();
            this.Component = ConvertObjectToString(name);
            this.OldName = this.Component;
            this.Notes = ConvertObjectToString(notes);
            this.Priority = NullIntegerCheck(priority);
            this.Position = NullIntegerCheck(position);
            this.Material = ConvertObjectToString(material);
            this.Finish = ConvertObjectToString(finish);
            this.TaskIDCount = NullIntegerCheck(taskIDCount);
        }
        /// <summary>
        /// Creates instance of a component with given name, sets TaskIDCount property to 0, and initializes a list of type TaskInfo.
        /// </summary> 
        public ComponentModel(object component, object notes, object priority, object position, object quantity, object spares, object picture, object material, object finish, object taskIDCount)
        {
            this.Tasks = new List<TaskModel>();
            this.Component = ConvertObjectToString(component);
            this.OldName = this.Component;
            this.Notes = ConvertObjectToString(notes);
            this.Priority = NullIntegerCheck(priority);
            this.Position = NullIntegerCheck(position);
            this.Quantity = NullIntegerCheck(quantity);
            this.Spares = NullIntegerCheck(spares);
            this.picture = NullByteArrayCheck(picture);
            this.Material = ConvertObjectToString(material);
            this.Finish = ConvertObjectToString(finish);
            this.TaskIDCount = NullIntegerCheck(taskIDCount);
        }
        /// <summary>
        /// Sets the name of a component.
        /// </summary>   
        public bool SetName(string name)
        {
            if (name.Length > ComponentCharacterLimit)
            {
                MessageBox.Show($"Component: '{name}' is greater than {ComponentCharacterLimit} characters. \n\nPlease shorten name.");
                return false;
            }

            this.Component = name;

            if(Tasks.Any())
            {
                foreach(TaskModel task in Tasks)
                {
                    task.SetComponent(name);
                }
            }

            return true;
        }
        /// <summary>
        /// Adds a task to a component.
        /// </summary>
        public void AddTask(string name, string component, SchedulerDataStorage schedulerStorage)
        {
            this.ReloadTaskList = true;
            this.Tasks.Add(new TaskModel(++TaskIDCount, name, component, schedulerStorage));
        }
        /// <summary>
        /// Adds a task to a component.
        /// </summary>
        public void AddTask(TaskModel task)
        {
            task.SetTaskID(++TaskIDCount);
            this.Tasks.Add(task);
        }
        /// <summary>
        /// Adds a tasklist to a component.
        /// </summary>
        public void AddTaskList(List<TaskModel> taskList)
        {
            this.Tasks = new List<TaskModel>();
            this.Tasks = taskList;

            this.Tasks.ForEach(x => x.TaskID = ++TaskIDCount);
        }
        /// <summary>
        /// Adds a copied tasklist to a component. (Sets IDs to 0)
        /// </summary>
        public List<TaskModel> AddCopiedTaskList(List<TaskModel> taskList, string newComponentName)
        {
            List<TaskModel> tasks = new List<TaskModel>();

            taskList.ForEach(task => tasks.Add(new TaskModel(task, newComponentName)));

            return tasks;
        }        
        public bool CheckIfAllTasksDated()
        {
            if (Tasks.Count(x => x.StartDate == null || x.FinishDate == null) > 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        /// <summary>
        /// Adds a picture to a component's picture list from a filepath.
        /// </summary>        
        //public void AddPicture(string filePath)
        //{
        //    try
        //    {
        //        Image image = Image.FromFile(filePath);
        //        PictureList.Add(image);
        //    }
        //    catch (OutOfMemoryException)
        //    {
        //        MessageBox.Show("The chosen file is not an image.");
        //    }

        //}
        /// <summary>
        /// Adds a picture to a component's picture list from the clipboard.
        /// </summary>
        public void SetPicture()
        {
            try
            {
                if(Clipboard.ContainsImage())
                {
                    this.picture = Clipboard.GetImage();
                }
                else
                {
                    MessageBox.Show("The clipboard contains no image.");
                }
            }
            catch (OutOfMemoryException)
            {
                MessageBox.Show("The item in the clipboard is not an image.");
            }
        }
        /// <summary>
        /// Adds a picture to a component's picture list from the pictureedit box.
        /// </summary>
        public void SetPicture(Image image)
        {
            this.picture = image;
        }
        /// <summary>
        /// Gets a picture from component class in the form of a byte array.
        /// </summary> 
        public byte[] GetPictureByteArray()
        {
            if(this.Picture != null)
            {
                //return ImageToByteArray(this.Picture);
                return this.Picture;
            }
            else
            {
                return null;
            }
            
        }
        /// <summary>
        /// Sets a picture in component class from a byte array.
        /// </summary> 
        public void SetPictureFromByteArray(byte[] pictureByteArr)
        {
            if (pictureByteArr != null)
            {
                //this.Picture = ByteArrayToImage(pictureByteArr);
            }
            else
            {
                MessageBox.Show("Byte array was null.");
            }

        }
        /// <summary>
        /// Removes a task from a component.
        /// </summary> 
        public void RemoveTask(int deletedTaskIndex)
        {
            this.ReloadTaskList = true;
            this.Tasks.Remove(Tasks.ElementAt(deletedTaskIndex));
            this.TaskIDCount = --TaskIDCount;
            //this.RemovedMatchingPredecessors(deletedTaskIndex + 1);
        }
        /// <summary>
        /// Moves a task up the task list.
        /// </summary> 
        public void MoveTaskUp(int promotedTaskIndex)
        {
            TaskModel promotedTask;

            if(promotedTaskIndex > 0)
            {
                this.ReloadTaskList = true;
                promotedTask = Tasks.ElementAt(promotedTaskIndex);

                this.Tasks.RemoveAt(promotedTaskIndex);
                this.Tasks.Insert(promotedTaskIndex - 1, promotedTask);
                //this.ModifyMatchingPredecessors(promotedTaskIndex, promotedTaskIndex - 1);
            }
            else
            {
                MessageBox.Show("Cannot move task any higher.");
            }
        }
        /// <summary>
        /// Moves a task down the task list.
        /// </summary>
        public void MoveTaskDown(int demotedTaskIndex)
        {
            TaskModel demotedTask;

            if (demotedTaskIndex < Tasks.Count - 1)
            {
                this.ReloadTaskList = true;
                demotedTask = Tasks.ElementAt(demotedTaskIndex);

                this.Tasks.RemoveAt(demotedTaskIndex);
                this.Tasks.Insert(demotedTaskIndex + 1, demotedTask);
                //this.ModifyMatchingPredecessors(demotedTaskIndex, demotedTaskIndex + 1);
            }
            else
            {
                MessageBox.Show("Cannot move task any lower.");
            }
        }
        /// <summary>
        /// Sets the note property of a component.
        /// </summary> 
        public void SetNote(string noteText)
        {
            this.Notes = noteText;
        }
        /// <summary>
        /// Sets the position property of a component.
        /// </summary> 
        public void SetPosition(int position)
        {
            this.Position = position;
        }
        /// <summary>
        /// Sets the quantity property of a component.
        /// </summary> 
        public void SetQuantity(int quantity)
        {
            this.Quantity = quantity;
        }
        /// <summary>
        /// Sets the spares property of a component.
        /// </summary> 
        public void SetSpares(int spares)
        {
            this.Spares = spares;
        }
        /// <summary>
        /// Sets the material property of a component.
        /// </summary> 
        public void SetMaterial(string material)
        {
            this.Material = material;
        }
        /// <summary>
        /// Sets the finish property of a component.
        /// </summary> 
        public void SetFinish(string finish)
        {
            this.Finish = finish;
        }
        /// <summary>
        /// Gets a task with matching ID from list of tasks for component.
        /// </summary> 
        public TaskModel GetTask(int taskID)
        {
            TaskModel task = (TaskModel)Tasks.Where(t => t.TaskID == taskID);

            return task;
        }
        // Used by Department Task View and Tasks Gridview in Project View tab.  Can't push dates backward.
        public void ChangeTaskDate(string fieldName, TaskModel task)
        {
            // Checks if the start date changed.
            if (fieldName == "StartDate")
            {
                if (task.StartDate == null)
                {
                    task.FinishDate = null;
                }
                else
                {
                    task.FinishDate = GeneralOperations.AddBusinessDays((DateTime)task.StartDate, task.Duration);
                }
            }

            TaskModel temp = Tasks.Find(x => x.ID == task.ID);

            temp.StartDate = task.StartDate;
            temp.FinishDate = task.FinishDate;

            if (task.FinishDate != null)
            {
                UpdateSuccessorTaskDates(task.TaskID, (DateTime)task.FinishDate);
            }

            this.AllTasksDated = CheckIfAllTasksDated();

            Database.UpdateComponent(this, "AllTasksDated");

            Database.UpdateTaskDates(Tasks);
        }
        /// <summary>
        /// Updates a task and handles moving predecessors or successors that overlap.
        /// </summary> 
        public bool UpdateTaskDates(TaskModel task)
        {
            bool batchUpdateTasks = false;

            if (task.StartDate == new DateTime(0001, 1, 1))
            {
                task.StartDate = null;
            }

            if (task.FinishDate == new DateTime(0001, 1, 1))
            {
                task.FinishDate = null;
            }

            if (task.StartDate == null)
            {
                goto SkipBackDating;
            }

            //DateTime? latestPredecessorFinishDate = this.GetLatestPredecessorFinishDate(task.Predecessors);

            TaskModel latestPredecessor = this.GetLatestPredecessor(task.Predecessors);

            if (latestPredecessor == null)
            {
                goto SkipBackDating;
            }

            if (task.Predecessors != "" && latestPredecessor.FinishDate > task.StartDate)
            {
                DialogResult dialogResult = MessageBox.Show("There is overlap between this task and one or more predecessors.  \n" +
                                                            "Do you wish to push the overlapping predecessors back?", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                if (dialogResult == DialogResult.Yes)
                {
                    // TODO: Validate this process.
                    this.BackDateTask(latestPredecessor.TaskID, (DateTime)task.StartDate);
                    batchUpdateTasks = true;
                }
                else if (dialogResult == DialogResult.No)
                {

                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    return false;
                }
            }

        SkipBackDating:

            if (this.SuccessorOverlap(task))
            {
                DialogResult dialogResult = MessageBox.Show("There is overlap between this task and one or more successors. \n" +
                                                            "Do you wish to push these tasks forward?", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                if (dialogResult == DialogResult.Yes)
                {
                    if (task.FinishDate != null)
                    {
                        UpdateSuccessorTaskDates(task.TaskID, (DateTime)task.FinishDate);
                        batchUpdateTasks = true;
                    }
                }
                else if (dialogResult == DialogResult.No)
                {

                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    return false;
                }
            }

            Database.UpdateTask(task);

            if (batchUpdateTasks == true)
            {
                Database.UpdateTaskDates(this.Tasks);
            }

            this.AllTasksDated = CheckIfAllTasksDated();

            return true;
        }
        /// <summary>
        /// Updates a task and moves all other tasks within component by the same number of days.
        /// </summary> 
        public bool UpdateTaskDates(TaskModel movedTask, DateTime oldTaskStartDate)
        {
            List<TaskModel> tempTasks = Tasks.FindAll(x => x.ProjectNumber == movedTask.ProjectNumber && x.Component == movedTask.Component && x.TaskID != movedTask.TaskID && x.Status != "Completed");

            TimeSpan dateDifference = ((DateTime)movedTask.StartDate - oldTaskStartDate);

            foreach (TaskModel task in tempTasks)
            {
                if (dateDifference.Days > 0)
                {
                    task.StartDate = GeneralOperations.AddBusinessDays((DateTime)task.StartDate, dateDifference.Days);
                    task.FinishDate = GeneralOperations.AddBusinessDays((DateTime)task.FinishDate, dateDifference.Days);  
                }  // GeneralOperations.AddBusinessDays(Convert.ToDateTime(task.FinishDate), dateDifference.Days);
                else
                {
                    task.StartDate = GeneralOperations.SubtractBusinessDays((DateTime)task.StartDate, dateDifference.Days * -1);
                    task.FinishDate = GeneralOperations.SubtractBusinessDays((DateTime)task.FinishDate, dateDifference.Days * -1);
                }
            }

            foreach (TaskModel task in Tasks)
            {
                Console.WriteLine($"Task: {task.TaskName,-13} Start Date: {((DateTime)task.StartDate).ToShortDateString(),-10} Finish Date: {GeneralOperations.AddBusinessDays((DateTime)task.StartDate, task.Duration).ToShortDateString()}");
            }

            Database.UpdateTaskDates(Tasks);

            return true;
        }
        /// <summary>
        /// Forward dates all tasks in given component.
        /// </summary> 
        public void ForwardDate(DateTime forwardDate)
        {
            List<TaskModel> leadingTasks = this.Tasks.FindAll(x => x.Predecessors == "" || x.Predecessors == null);

            foreach (var task in leadingTasks)
            {
                this.ForwardDateTask(task.TaskID, forwardDate);
            }

            this.AllTasksDated = CheckIfAllTasksDated();
        }
        public void ForwardDateTask(int successorID, DateTime predecessorFinishDate)
        {
            TaskModel successorTask = this.Tasks.Find(x => x.TaskID == successorID);

            if (successorTask.StartDate == null || predecessorFinishDate > successorTask.StartDate)
            {
                successorTask.StartDate = predecessorFinishDate;
                successorTask.FinishDate = GeneralOperations.AddBusinessDays((DateTime)successorTask.StartDate, successorTask.Duration);
            }

            var result = from tasks in Tasks
                         where tasks.HasMatchingPredecessor(successorID)
                         select tasks;

            foreach (TaskModel newSuccessorTask in result)
            {
                ForwardDateTask(newSuccessorTask.TaskID, (DateTime)successorTask.FinishDate);
            }
        }
        public void BackDate(DateTime backDate)
        {
            this.ClearTaskDates();

            TaskModel finalTask = this.Tasks.Find(x => x.TaskID == this.Tasks.Max(x2 => x2.TaskID));

            BackDateTask(finalTask.TaskID, backDate);

            this.AllTasksDated = CheckIfAllTasksDated();
        }
        private void BackDateTask(int predecessorID, DateTime successorStartDate)
        {
            TaskModel predecessorTask = this.Tasks.Find(x => x.TaskID == predecessorID);

            if (predecessorTask.FinishDate == null || predecessorTask.FinishDate > successorStartDate)
            {
                predecessorTask.FinishDate = successorStartDate;
                predecessorTask.StartDate = GeneralOperations.SubtractBusinessDays((DateTime)predecessorTask.FinishDate, predecessorTask.Duration); 
            }

            foreach (int predecessor in predecessorTask.GetPredecessorList())
            {
                BackDateTask(predecessor,(DateTime)predecessorTask.StartDate);
            }
        }
        public void UpdatePredecessorTaskDates(int predecessorID, int daysToMove)
        {
            TaskModel predecessorTask = this.Tasks.FirstOrDefault(x => x.TaskID == predecessorID && x.Status != "Completed");

            if (predecessorTask != null)
            {
                if (predecessorTask.FinishDate != null)
                {
                    predecessorTask.StartDate = GeneralOperations.SubtractBusinessDays((DateTime)predecessorTask.StartDate, daysToMove);
                    predecessorTask.FinishDate = GeneralOperations.SubtractBusinessDays((DateTime)predecessorTask.FinishDate, daysToMove); 
                }

                foreach (int predecessor in predecessorTask.GetPredecessorList())
                {
                    UpdatePredecessorTaskDates(predecessor, daysToMove);
                }
            }
        }
        public List<TaskModel> UpdateSuccessorTaskDates(int taskID, DateTime? finishDate, bool fillBlanks = false, bool pullBackStartDates = false)
        {
            var result = from task in Tasks
                         where task.HasMatchingPredecessor(taskID)
                         select task;

            Console.WriteLine("Update Start and Finish Dates");

            foreach (TaskModel task in result)
            {
                if (task.StartDate == null)
                {
                    if (fillBlanks == true)
                    {
                        task.StartDate = finishDate;
                        task.FinishDate = GeneralOperations.AddBusinessDays((DateTime)task.StartDate, task.Duration.ToString());
                    }
                }
                else if (finishDate > (DateTime)task.StartDate) // If start date of current task comes before finish date of predecessor.
                {
                    task.StartDate = finishDate;
                    task.FinishDate = GeneralOperations.AddBusinessDays((DateTime)task.StartDate, task.Duration.ToString());
                }
                else if (finishDate < (DateTime)task.StartDate && pullBackStartDates == true) // If start date of current task comes after the finish date of predecessor.
                {
                    task.StartDate = finishDate;
                    task.FinishDate = GeneralOperations.AddBusinessDays((DateTime)task.StartDate, task.Duration.ToString());
                }

                if (task.FinishDate != null)
                    UpdateSuccessorTaskDates(task.TaskID, Convert.ToDateTime(task.FinishDate), fillBlanks, pullBackStartDates);
            }

            return Tasks;
        }
        public List<TaskModel> UpdateSuccessorTaskDates(TaskModel movedTask, int daysToMove, bool fillBlanks = false)
        {
            var result = from tasks in Tasks
                         where tasks.HasMatchingPredecessor(movedTask.TaskID)
                         select tasks;

            Console.WriteLine("Update Start and Finish Dates");

            foreach (TaskModel task in result)
            {
                if (task.StartDate == null)
                {
                    if (fillBlanks == true)
                    {
                        task.StartDate = GeneralOperations.AddBusinessDays((DateTime)task.StartDate, daysToMove);
                        task.FinishDate = GeneralOperations.AddBusinessDays((DateTime)task.FinishDate, daysToMove);
                    }
                }
                else
                {
                    task.StartDate = GeneralOperations.AddBusinessDays((DateTime)task.StartDate, daysToMove);
                    task.FinishDate = GeneralOperations.AddBusinessDays((DateTime)task.FinishDate, daysToMove);
                }

                if (task.FinishDate != null)
                    UpdateSuccessorTaskDates(task, daysToMove);
            }

            return Tasks;
        }
        public void ClearTaskDates()
        {
            foreach (TaskModel task in Tasks)
            {
                task.StartDate = null;
                task.FinishDate = null;
            }
        }
        public void UpdateComponent(ComponentModel component) // Used when existing components are overwritten by components from template.
        {
            //this.JobNumber = component.JobNumber;
            //this.ProjectNumber = component.ProjectNumber;
            this.OldName = component.Component;
            this.Notes = component.Notes;
            this.TaskIDCount = component.TaskIDCount;
            this.Quantity = component.Quantity;
            this.Spares = component.Spares;
            this.Material = component.Material;
            this.Finish = component.Finish;
            this.Tasks = component.Tasks;
            this.Picture = component.Picture;
        }
        private string ConvertObjectToString(object obj)
        {
            if (obj != null)
            {
                return obj.ToString();
            }
            else
            {
                return "";
            }
        }

        private int NullIntegerCheck(object checkValue)
        {
            if (!DBNull.Value.Equals(checkValue))
            {
                return Convert.ToInt16(checkValue);
            }
            else
            {
                return 0;
            }
        }

        public static Image NullByteArrayCheck(object obj)
        {
            if (!DBNull.Value.Equals(obj) && obj != null)
            {
                return ByteArrayToImage((byte[]) obj);
            }
            else
            {
                return null;
            }
        }

        public static byte[] ImageToByteArray(Image imageIn)
        {
            //MemoryStream ms = new MemoryStream();
            //imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Gif);
            //return ms.ToArray();

            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(imageIn, typeof(byte[]));
        }
        public static Image ByteArrayToImage(byte[] byteArrayIn)
        {
            //MemoryStream ms = new MemoryStream(byteArrayIn);
            //Image returnImage = Image.FromStream(ms);
            //return returnImage;

            if (byteArrayIn.Length == 0)
            {
                return null;
            }

            ImageConverter ic = new ImageConverter();
            Image img = (Image)ic.ConvertFrom(byteArrayIn);
            return img;
        }
        public string GetPictureString()
        {
            //return Convert.ToBase64String(ImageToByteArray(this.Picture));
            return Convert.ToBase64String(this.Picture);
        }
        public static bool IsGoodComponentPicture(Image image)
        {
            int maxWidth = 865, maxHeight = 795;

            MessageBox.Show($"Width: {image.Width} Height: {image.Height}");

            if (image != null && (image.Width > maxWidth || image.Height > maxHeight))
            {
                MessageBox.Show($"The picture you inserted is too wide or too tall to put in a Kan Ban and print correctly. " +
                                $"({image.Width} pixels wide x {image.Height} pixels tall)\n\n" +
                                $"Max width is {maxWidth} pixels and max height is {maxHeight} pixels.\n\n" +
                                $"Take a smaller picture using Snipping Tool, Snip and Sketch, or Snagit and try again.");

                return false;
            }

            return true;
        }
        public List<TaskModel> FindSelfReferencingTasks()
        {
            var result = from task in this.Tasks
                         where task.HasMatchingPredecessor(this.Tasks.IndexOf(task) + 1)
                         select task;

            return result.ToList();
        }
        public List<TaskModel> FindTasksWithNullDates()
        {
            var result = from task in this.Tasks
                         where task.HasNullDates()
                         select task;

            return result.ToList();
        }
        public List<TaskModel> FindIsolatedTasks()
        {
            var result = from task in this.Tasks
                         where task.Predecessors.Length == 0 && 
                               // !task.TaskName.Contains("Program") && // Why was this commented out?
                               this.Tasks.Count(x => x.HasMatchingPredecessor(this.Tasks.IndexOf(task) + 1)) == 0 &&
                               this.Tasks.Count > 1
                         select task;

            //foreach (var item in this.Tasks)
            //{
            //    if (item.Predecessors.Length == 0 && Tasks.Count(x => x.HasMatchingPredecessor(Tasks.IndexOf(item) + 1)) == 0 && Tasks.Count > 1)
            //    {
            //        Console.WriteLine($"{item.TaskName} {item.TaskID} {item.Predecessors}");
            //    }
            //}

            return result.ToList();
        }
        public void ModifyMatchingPredecessors(int id, int newID)
        {
            var result = from task in this.Tasks
                         where task.HasMatchingPredecessor(id)
                         select task;

            foreach (var task in result.ToList())
            {
                task.ChangeMatchingPredecessor(id, newID);
            }
        }
        public void RemovedMatchingPredecessors(int id)
        {
            var result = from task in this.Tasks
                         where task.HasMatchingPredecessor(id)
                         select task;

            foreach (var task in result.ToList())
            {
                task.RemoveMatchingPredecessor(id);
            }
        }
        public DateTime? GetLatestFinishDate()
        {
            return this.Tasks.Max(x => x.FinishDate);
        }
        public TaskModel GetLatestPredecessor(string predecessors)
        {
            DateTime? latestFinishDate = null;
            DateTime? currentDate = null;
            string[] predecessorArr;
            string predecessor, latestPredecessor = "";

            if (predecessors != "")
            {
                if (predecessors.Contains(","))
                {
                    predecessorArr = predecessors.Split(',');

                    foreach (string currPredecessor in predecessorArr)
                    {
                        predecessor = currPredecessor.Trim();

                        if (predecessor == "")
                        {
                            break;
                        }

                        currentDate = this.Tasks.Find(x => x.TaskID.ToString() == predecessor).FinishDate;

                        if (latestFinishDate == null || latestFinishDate < currentDate)
                        {
                            latestPredecessor = predecessor;
                            latestFinishDate = currentDate;
                        }
                    }

                    return this.Tasks.Find(x => x.TaskID.ToString() == latestPredecessor);
                }
                else
                {
                    return this.Tasks.Find(x => x.TaskID.ToString() == predecessors);
                } 
            }
            else
            {
                return null;
            }
        }
        public DateTime? GetLatestPredecessorFinishDate(string predecessors)
        {
            DateTime? latestFinishDate = null;
            DateTime? currentDate = null;
            string[] predecessorArr;
            string predecessor;

            predecessorArr = predecessors.Split(',');

            foreach (string currPredecessor in predecessorArr)
            {
                predecessor = currPredecessor.Trim();

                if (predecessor == "")
                {
                    break;
                }

                currentDate = this.Tasks.Find(x => x.TaskID.ToString() == predecessor).FinishDate;

                if (latestFinishDate == null || latestFinishDate < currentDate)
                {
                    latestFinishDate = currentDate;
                }
            }

            return latestFinishDate;
        }
        public bool SuccessorOverlap(TaskModel task)
        {
            List<TaskModel> successors = this.Tasks.FindAll(x => x.HasMatchingPredecessor(task.TaskID));

            foreach (TaskModel currentTask in successors)
            {
                if (task.FinishDate > currentTask.StartDate)
                {
                    return true;
                } 
            }

            return false;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            if (propertyName == "ProjectNumber")
            {
                Tasks.ForEach(x => x.ProjectNumber = this.ProjectNumber);
            }
            else if (propertyName == "JobNumber")
            {
                Tasks.ForEach(x => x.JobNumber = this.JobNumber);
            }
            else if (propertyName == "Component")
            {
                Tasks.ForEach(x => x.Component = this.Component);
            }
        }
    }
}
