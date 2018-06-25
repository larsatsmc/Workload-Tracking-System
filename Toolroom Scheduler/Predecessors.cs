using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Toolroom_Scheduler
{
    class Predecessors
    {
        public Predecessors(int taskTableID, int predecessor)
        {
            this.TaskTableID = taskTableID;
            this.Predecessor = predecessor;
        }

        public int TaskTableID { get; private set; }
        public int Predecessor { get; private set; }
    }
}
