using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Toolroom_Scheduler
{
    class ResourceInfo
    {
        public ResourceInfo(string fn, string ln, string role)
        {
            this.FirstName = fn;
            this.LastName = ln;
            this.Role = role;
        }

        public string FirstName { get; private set; }
        public string LastName { get; private set; }
        public string Role { get; private set; }
    }
}
