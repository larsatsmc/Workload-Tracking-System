using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
{
    public class ResourceModel
    {
        public int ID { get; set; }
        public string ResourceName { get; set; }
        public string ResourceType { get; set; }
        public List<RoleModel> MyProperty { get; set; }
    }
}
