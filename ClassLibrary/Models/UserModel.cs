using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary.Models
{
    public class UserModel
    {
        public int ID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string LoginName { get; set; }
        public bool IsAdmin { get; set; }  // Can add or remove users.  Provide or revoke previliges.
        public bool CanChangeDates { get; set; }
        public bool EngineeringNumberVisible { get; set; }
        public bool CanReadOnly { get; set; }
        public bool CanOnlyChangeDesignWork { get; set; }
        public bool CanChangeProjectData { get; set; }
    }
}
