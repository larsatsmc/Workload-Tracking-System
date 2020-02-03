using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
{
    public class RoleModel
    {
        public int ID { get; set; }
        public string Role { get; set; }
        public int ResourceID { get; set; }
        public int DepartmentID { get; set; }
    }
}
