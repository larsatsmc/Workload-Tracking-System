using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace ClassLibrary
{
    public class ComponentModel
    {
        /// <summary>
        /// Gets the name property of a Component object.
        /// </summary>
        public string Name { get; private set; }
        public string OldName { get; private set; }
        public Image Picture { get; private set; }
        public string Material { get; private set; }
        public string Notes { get; private set; }
        public List<TaskModel> TaskList { get; private set; }
        public bool ReloadTaskList { get; set; }
        public int Hours { get; private set; }
        public int Priority { get; private set; }
        public int Position { get; private set; }
        public int TaskIDCount { get; private set; }
        public int Quantity { get; private set; }
        public int Spares { get; private set; }
        public string Initials { get; private set; }
        public string Finish { get; private set; }
        public string Status { get; private set; }
        public int PercentComplete { get; private set; }

        public static int ComponentCharacterLimit = 31;
        /// <summary>
        /// Creates instance of a component and sets TaskIDCount property to 0.
        /// </summary> 
        public ComponentModel() // Default contructors do not execute unless they are called.
        {
            TaskIDCount = 0;
        }
        /// <summary>
        /// Creates instance of a component and sets properties for template.
        /// </summary> 
        public ComponentModel(string name, string quantity, string spares, string material, string finish, string notes)
        {
            this.Name = name;
            this.Quantity = Convert.ToInt16(quantity);
            this.Spares = Convert.ToInt16(spares);
            this.Material = material;
            this.Finish = finish;
            this.Notes = notes;
            this.TaskList = new List<TaskModel>();
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
            this.TaskList = new List<TaskModel>();
            this.Name = name;
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
            this.TaskList = new List<TaskModel>();
            this.Name = ConvertObjectToString(name);
        }
        /// <summary>
        /// Creates instance of a component with given name, sets TaskIDCount property to 0, and initializes a list of type TaskInfo.
        /// </summary> 
        public ComponentModel(object name, object notes, object priority, object position, object material, object finish, object taskIDCount)
        {
            this.TaskList = new List<TaskModel>();
            this.Name = ConvertObjectToString(name);
            this.OldName = this.Name;
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
        public ComponentModel(object name, object notes, object priority, object position, object quantity, object spares, object picture, object material, object finish, object taskIDCount)
        {
            this.TaskList = new List<TaskModel>();
            this.Name = ConvertObjectToString(name);
            this.OldName = this.Name;
            this.Notes = ConvertObjectToString(notes);
            this.Priority = NullIntegerCheck(priority);
            this.Position = NullIntegerCheck(position);
            this.Quantity = NullIntegerCheck(quantity);
            this.Spares = NullIntegerCheck(spares);
            this.Picture = NullByteArrayCheck(picture);
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

            this.Name = name;

            if(TaskList.Any())
            {
                foreach(TaskModel task in TaskList)
                {
                    task.SetComponent(name);
                }
            }

            return true;
        }
        /// <summary>
        /// Adds a task to a component.
        /// </summary>
        public void AddTask(string name, string component)
        {
            this.ReloadTaskList = true;
            this.TaskList.Add(new TaskModel(++TaskIDCount, name, component));
        }
        /// <summary>
        /// Adds a task to a component.
        /// </summary>
        public void AddTask(TaskModel task)
        {
            task.SetTaskID(++TaskIDCount);
            this.TaskList.Add(task);
        }
        /// <summary>
        /// Adds a tasklist to a component.
        /// </summary>
        public void AddTaskList(List<TaskModel> taskList)
        {
            this.TaskList = new List<TaskModel>();
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
        /// Adds a picture to a component's picture list from the pictureedit box.
        /// </summary>
        public void SetPicture(Image image)
        {
            this.Picture = image;
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
            this.TaskList.Remove(TaskList.ElementAt(deletedTaskIndex));
            this.TaskIDCount = --TaskIDCount;
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
                promotedTask = TaskList.ElementAt(promotedTaskIndex);

                this.TaskList.RemoveAt(promotedTaskIndex);
                this.TaskList.Insert(promotedTaskIndex - 1, promotedTask);
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

            if (demotedTaskIndex < TaskList.Count - 1)
            {
                this.ReloadTaskList = true;
                demotedTask = TaskList.ElementAt(demotedTaskIndex);

                this.TaskList.RemoveAt(demotedTaskIndex);
                this.TaskList.Insert(demotedTaskIndex + 1, demotedTask);
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
            TaskModel task = (TaskModel)TaskList.Where(t => t.ID == taskID);

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

            ImageConverter ic = new ImageConverter();
            Image img = (Image)ic.ConvertFrom(byteArrayIn);
            return img;
        }
        public static bool IsGoodComponentPicture(Image image)
        {
            int maxWidth = 865, maxHeight = 795;

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
    }
}
