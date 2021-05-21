using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
{
    public class AptDependencyModel
    {
        public int ParentID { get; set; }
        public int DependentID { get; set; }
        public static List<AptDependencyModel> GetDependencyData(ProjectModel project)
        {
            List<AptDependencyModel> aptDependencies = new List<AptDependencyModel>();

            AptDependencyModel aptDependency;

            foreach (var component in project.Components)
            {
                foreach (var task in component.Tasks)
                {
                    if (task.NewPredecessors.Contains(","))
                    {
                        foreach (string predecessor in task.NewPredecessors.ToString().Split(','))
                        {
                            aptDependency = new AptDependencyModel();

                            aptDependency.DependentID = task.AptID;
                            aptDependency.ParentID = Convert.ToInt32(predecessor);

                            aptDependencies.Add(aptDependency);

                            //Console.WriteLine($"{nrow["TaskID"]} {predecessor}");
                        }
                    }
                    else if (task.NewPredecessors.ToString() != "")
                    {
                        aptDependency = new AptDependencyModel();

                        aptDependency.DependentID = task.AptID;
                        aptDependency.ParentID = Convert.ToInt32(task.NewPredecessors);

                        aptDependencies.Add(aptDependency);

                        //Console.WriteLine($"{nrow["TaskID"]} {nrow["Predecessors"]}");
                    }
                } 
            }

            return aptDependencies;
        }
    }
}