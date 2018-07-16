using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Toolroom_Scheduler
{
    public class QuoteInfo
    {
        public string Customer { get; private set; }
        public string PartName { get; private set; }
        public int ProgramRoughHours { get; private set; }
        public int ProgramFinishHours { get; private set; }
        public int ProgramElectrodeHours { get; private set; }
        public int CNCRoughHours { get; private set; }
        public int CNCFinishHours { get; private set; }
        public int GrindFittingHours { get; private set; }
        public int CNCElectrodeHours { get; private set; }
        public int EDMSinkerHours { get; private set; }
        public int EDMWireHours { get; private set; }
        public List<TaskInfo> TaskList { get; private set; }

        public QuoteInfo(string customer, string partName, int programRoughHours, int programFinishHours, int programElectrodeHours, int cncRoughHours, int cncFinishHours, int grindFittingHours, int cncElectrodeHours, int edmSinkerHours)
        {
            this.Customer = customer;
            this.PartName = partName;
            this.ProgramRoughHours = programRoughHours;
            this.ProgramFinishHours = programFinishHours;
            this.ProgramElectrodeHours = programElectrodeHours;
            this.CNCRoughHours = cncRoughHours;
            this.CNCFinishHours = cncFinishHours;
            this.GrindFittingHours = grindFittingHours;
            this.CNCElectrodeHours = cncElectrodeHours;
            this.EDMSinkerHours = edmSinkerHours;

            CreateTaskList();
        }

        private int GetTaskHours(string taskName)
        {
            if(taskName == "Program Rough")
            {
                return this.ProgramRoughHours;
            }
            else if(taskName == "Program Finish")
            {
                return this.ProgramFinishHours;
            }
            else if (taskName == "Program Electrodes")
            {
                return this.ProgramElectrodeHours;
            }
            else if (taskName == "CNC Rough")
            {
                return this.CNCRoughHours;
            }
            else if (taskName == "CNC Finish")
            {
                return this.CNCFinishHours;
            }
            else if (taskName == "CNC Electrodes")
            {
                return this.CNCElectrodeHours;
            }
            else if (taskName == "EDM Sinker")
            {
                return this.EDMSinkerHours;
            }
            else if (taskName == "EDM Wire")
            {
                return this.EDMSinkerHours;
            }
            else if (taskName == "Grind-Fitting")
            {
                return this.GrindFittingHours;
            }

            return 0;
        }

        private void CreateTaskList()
        {
            List<string> taskNameList = new List<string> { "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "Grind-Fitting", "CNC Electrodes", "EDM Sinker"};
            TaskList = new List<TaskInfo>();

            foreach (string taskName in taskNameList)
            {
                TaskInfo task = new TaskInfo();
                int hours = GetTaskHours(taskName);

                task.SetName(taskName);
                task.SetComponent("Quote");
                task.SetHours(hours);
                task.SetDuration((int)(hours * 1.4 / 8));
                task.HasInfo = true;


                TaskList.Add(task);
            }
        }
    }
}
