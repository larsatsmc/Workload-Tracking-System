using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
{
    public class QuoteModel
    {
        public string Customer { get; private set; }
        public string PartName { get; private set; }
        public int DesignHours { get; private set; }
        public int ProgramRoughHours { get; private set; }
        public int ProgramFinishHours { get; private set; }
        public int ProgramElectrodeHours { get; private set; }
        public int CNCRoughHours { get; private set; }
        public int CNCFinishHours { get; private set; }
        public int GrindFittingHours { get; private set; }
        public int CNCElectrodeHours { get; private set; }
        public int EDMSinkerHours { get; private set; }
        public int EDMWireHours { get; private set; }
        public List<TaskModel> TaskList { get; private set; }

        public QuoteModel()
        {

        }

        public QuoteModel(object customer, object partName, object designHours, object designElectrodeHours, object programRoughHours, object programFinishHours, object programElectrodeHours, object cncRoughHours, object cncFinishHours, object grindFittingHours, object cncElectrodeHours, object edmSinkerHours, object edmWireHours)
        {
            this.Customer = customer.ToString();
            this.PartName = partName.ToString();
            this.DesignHours = Convert.ToInt16(designHours);
            this.ProgramRoughHours = Convert.ToInt16(programRoughHours);
            this.ProgramFinishHours = Convert.ToInt16(programFinishHours);
            this.ProgramElectrodeHours = Convert.ToInt16(programElectrodeHours) + Convert.ToInt16(designElectrodeHours);
            this.CNCRoughHours = Convert.ToInt16(cncRoughHours);
            this.CNCFinishHours = Convert.ToInt16(cncFinishHours);
            this.GrindFittingHours = Convert.ToInt16(grindFittingHours);
            this.CNCElectrodeHours = Convert.ToInt16(cncElectrodeHours);
            this.EDMSinkerHours = Convert.ToInt16(edmSinkerHours);
            this.EDMWireHours = Convert.ToInt16(edmWireHours);

            CreateTaskListForQuote();
        }

        private int GetTaskHours(string taskName)
        {
            if (taskName == "Design")
            {
                return this.DesignHours;
            }
            else if(taskName == "Program Rough")
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
            else if (taskName == "EDM Wire (In-House)")
            {
                return this.EDMWireHours;
            }
            else if (taskName == "Grind-Fitting")
            {
                return this.GrindFittingHours;
            }

            return 0;
        }

        private void CreateTaskListForQuote()
        {
            List<string> taskNameList = new List<string> {"Design", "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "Grind-Fitting", "CNC Electrodes", "EDM Sinker", "EDM Wire (In-House)"};
            List<string> taskPredecessorList = new List<string>() { "", "1", "1", "2", "2", "3,5", "5", "4", "6", "9"};
            int index = 0;
            TaskList = new List<TaskModel>();

            foreach (string taskName in taskNameList)
            {
                TaskModel task = new TaskModel();
                int hours = GetTaskHours(taskName);

                task.SetName(taskName);
                task.SetComponent("Quote");
                task.SetHours(hours);
                task.SetDuration((int)(hours * 1.4 / 8));
                task.Predecessors = taskPredecessorList.ElementAt(index++);
                task.HasInfo = true;


                TaskList.Add(task);
            }
        }
    }
}
