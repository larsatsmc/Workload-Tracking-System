using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ClassLibrary
{
    public class AptResourceModel
    {
        public int AptID { get; set; }
        public string TaskName { get; set; }
        public int ParentID { get; set; }
        public static List<AptResourceModel> GetProjectResourceData(ProjectModel project)
        {
            List<AptResourceModel> aptResourceList = new List<AptResourceModel>();
            AptResourceModel componentResource;
            AptResourceModel taskResource;
            int count = 1;
            int parentID;
            int baseCount;

            foreach (ComponentModel component in project.Components)
            {
                componentResource = new AptResourceModel();

                baseCount = count;
                componentResource.AptID = count;                
                componentResource.TaskName = component.Component;
                parentID = count++;

                aptResourceList.Add(componentResource);

                foreach (TaskModel task in component.Tasks)
                {
                    taskResource = new AptResourceModel();

                    task.AptID = count;
                    task.SetNewPredecessors(baseCount);
                    taskResource.AptID = count++;
                    taskResource.TaskName = task.TaskName;
                    taskResource.ParentID = parentID;

                    aptResourceList.Add(taskResource);
                }
            }

            return aptResourceList;
        }
    }
}