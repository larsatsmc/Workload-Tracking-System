using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;

namespace Toolroom_Scheduler
{
    public class Component
    {
        /// <summary>
        /// Gets the name property of a Component object.
        /// </summary>
        public string Name { get; private set; }
        public string OldName { get; private set; }
        public Image Picture { get; private set; }
        public string Material { get; private set; }
        public string Notes { get; private set; }
        public List<TaskInfo> TaskList { get; private set; }
        public bool ReloadTaskList { get; set; }
        public int Priority { get; private set; }
        public int Position { get; private set; }
        public int TaskIDCount { get; private set; }
        public int Quantity { get; private set; }
        public int Spares { get; private set; }
        public string Initials { get; private set; }

        /// <summary>
        /// Creates instance of a component and sets TaskIDCount property to 0.
        /// </summary> 
        public Component() // Default contructors do not execute unless they are called.
        {
            TaskIDCount = 0;
        }
        /// <summary>
        /// Creates instance of a component with given name, sets TaskIDCount property to 0, and initializes a list of type TaskInfo.
        /// </summary> 
        public Component(string name)
        {
            TaskIDCount = 0;
            this.Quantity = 1;
            this.Spares = 0;
            this.Notes = "";
            this.Material = "";
            TaskList = new List<TaskInfo>();
            this.Name = name;
        }
        /// <summary>
        /// Creates instance of a component with given name, sets TaskIDCount property to 0, and initializes a list of type TaskInfo.
        /// </summary> 
        public Component(object name)
        {
            TaskIDCount = 0;
            this.Quantity = 1;
            this.Spares = 0;
            this.Notes = "";
            this.Material = "";
            TaskList = new List<TaskInfo>();
            this.Name = ConvertObjectToString(name);
        }
        /// <summary>
        /// Creates instance of a component with given name, sets TaskIDCount property to 0, and initializes a list of type TaskInfo.
        /// </summary> 
        public Component(object name, object notes, object priority, object position, object quantity, object spares, object picture, object material, object taskIDCount)
        {
            TaskList = new List<TaskInfo>();
            this.Name = ConvertObjectToString(name);
            this.OldName = this.Name;
            this.Notes = ConvertObjectToString(notes);
            this.Priority = NullIntegerCheck(priority);
            this.Position = NullIntegerCheck(position);
            this.Quantity = NullIntegerCheck(quantity);
            this.Spares = NullIntegerCheck(spares);
            this.Picture = NullByteArrayCheck(picture);
            this.Material = ConvertObjectToString(material);
            this.TaskIDCount = NullIntegerCheck(taskIDCount);
        }
        /// <summary>
        /// Sets the name of a component.
        /// </summary>   
        public void SetName(string name)
        {
            this.Name = name;

            if(TaskList.Any())
            {
                foreach(TaskInfo task in TaskList)
                {
                    task.SetComponent(name);
                }
            }
        }
        /// <summary>
        /// Adds a task to a component.
        /// </summary>
        public void AddTask(string name, string component)
        {
            this.ReloadTaskList = true;
            this.TaskList.Add(new TaskInfo(++TaskIDCount, name, component));
        }
        /// <summary>
        /// Adds a task to a component.
        /// </summary>
        public void AddTask(TaskInfo task)
        {
            task.SetTaskID(++TaskIDCount);
            this.TaskList.Add(task);
        }
        /// <summary>
        /// Adds a task to a component.
        /// </summary>
        public void AddTaskList(List<TaskInfo> taskList)
        {
            TaskList = new List<TaskInfo>();
            this.TaskList = taskList;
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
                    this.Picture = Clipboard.GetImage();
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
        /// Gets a picture from component class in the form of a byte array.
        /// </summary> 
        public byte[] GetPictureByteArray()
        {
            if(this.Picture != null)
            {
                return ImageToByteArray(this.Picture);
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
                this.Picture = ByteArrayToImage(pictureByteArr);
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
            TaskList.Remove(TaskList.ElementAt(deletedTaskIndex));
            this.TaskIDCount = --TaskIDCount;
        }
        /// <summary>
        /// Moves a task up the task list.
        /// </summary> 
        public void MoveTaskUp(int promotedTaskIndex)
        {
            TaskInfo promotedTask;

            if(promotedTaskIndex > 0)
            {
                this.ReloadTaskList = true;
                promotedTask = TaskList.ElementAt(promotedTaskIndex);

                TaskList.RemoveAt(promotedTaskIndex);
                TaskList.Insert(promotedTaskIndex - 1, promotedTask);
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
            TaskInfo demotedTask;

            if (demotedTaskIndex < TaskList.Count - 1)
            {
                this.ReloadTaskList = true;
                demotedTask = TaskList.ElementAt(demotedTaskIndex);

                TaskList.RemoveAt(demotedTaskIndex);
                TaskList.Insert(demotedTaskIndex + 1, demotedTask);
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
        /// Gets a task with matching ID from list of tasks for component.
        /// </summary> 
        public TaskInfo GetTask(int taskID)
        {
            TaskInfo task = (TaskInfo)TaskList.Where(t => t.ID == taskID);

            return task;
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

        private Image NullByteArrayCheck(object obj)
        {
            if (!DBNull.Value.Equals(obj))
            {
                return ByteArrayToImage((byte[]) obj);
            }
            else
            {
                return null;
            }
        }

        private byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            //MemoryStream ms = new MemoryStream();
            //imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Gif);
            //return ms.ToArray();

            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(imageIn, typeof(byte[]));
        }

        private Image ByteArrayToImage(byte[] byteArrayIn)
        {
            //MemoryStream ms = new MemoryStream(byteArrayIn);
            //Image returnImage = Image.FromStream(ms);
            //return returnImage;

            ImageConverter ic = new ImageConverter();
            Image img = (Image)ic.ConvertFrom(byteArrayIn);
            return img;
        }
    }
}
