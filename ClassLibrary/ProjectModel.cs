using DevExpress.XtraScheduler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClassLibrary
{
    public class ProjectModel
    {
        public int ProjectNumber { get; private set; }
        public bool ProjectNumberChanged { get; set; }
        public int OldProjectNumber { get; private set; }
        public string Project { get; set; }
        public string JobNumber { get; set; }
        public int MWONumber { get; set; }
        public string Customer { get; set; } = "";
        public string Name { get; set; } = "";
        public DateTime DueDate { get; private set; }
        public int Priority { get; private set; }
        public string Status { get; private set; }
        public string ToolMaker { get; private set; }
        public string Designer { get; set; }
        public string RoughProgrammer { get; set; }
        public string ElectrodeProgrammer { get; set; }
        public string FinishProgrammer { get; set; }
        public string Apprentice { get; set; } = "";
        public string Engineer { get; private set; }
        public string KanBanWorkbookPath { get; private set; }
        public List<ComponentModel> ComponentList { get; private set; }
        public QuoteModel QuoteInfo { get; private set; }
        public bool HasProjectInfo { get; set; }
        public SchedulerStorage AvailableResources { get; set; }
        public bool OverlapAllowed { get; set; }

        public ProjectModel()
        {
            ComponentList = new List<ComponentModel>();
            OverlapAllowed = true;
        }

        public ProjectModel(string jn, int pn, DateTime dd, int p, string s, string tm, string d, string rp, string fp, string ep, string e)
        {
            this.HasProjectInfo = true;
            this.JobNumber = jn;
            this.ProjectNumber = pn;
            this.OldProjectNumber = pn;
            this.DueDate = new DateTime(dd.Year,dd.Month, dd.Day);
            this.Priority = p;
            this.Status = s;
            this.Designer = d;
            this.ToolMaker = tm;
            this.RoughProgrammer = rp;
            this.ElectrodeProgrammer = ep;
            this.FinishProgrammer = fp;
            this.Engineer = e;
        }

        public ProjectModel(string jobNumber, int projectNumber, DateTime dueDate, string status, string toolMaker, string designer, string roughProgrammer, string finishProgrammer, string electrodProgrammer, string apprentice, string kanBanWorkbookPath) // Project Creation Constructor. Leaving out status for now.  May add later.
        {
            this.HasProjectInfo = true;
            this.JobNumber = jobNumber;
            this.ProjectNumber = projectNumber;
            this.OldProjectNumber = projectNumber;
            this.DueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
            this.Status = status;
            this.Designer = designer;
            this.ToolMaker = toolMaker;
            this.RoughProgrammer = roughProgrammer;
            this.ElectrodeProgrammer = electrodProgrammer;
            this.FinishProgrammer = finishProgrammer;
            this.Apprentice = apprentice;
            this.ComponentList = new List<ComponentModel>();
            this.KanBanWorkbookPath = kanBanWorkbookPath;
        }

        public ProjectModel(object jobNumber, object projectNumber, object mwoNumber, object customer, object project, object dueDate, object toolMaker, object designer, object roughProgrammer, object finishProgrammer, object electrodProgrammer, object apprentice)
        {
            this.JobNumber = ConvertObjectToString(jobNumber);
            this.ProjectNumber = ConvertObjectToInt(projectNumber);
            this.Customer = ConvertObjectToString(customer);
            this.Name = ConvertObjectToString(project);
            this.MWONumber = ConvertObjectToInt(mwoNumber);
            this.DueDate = ConvertObjectToDateTime(dueDate);
            this.Designer = ConvertObjectToString(designer);
            this.ToolMaker = ConvertObjectToString(toolMaker);
            this.RoughProgrammer = ConvertObjectToString(roughProgrammer);
            this.ElectrodeProgrammer = ConvertObjectToString(electrodProgrammer);
            this.FinishProgrammer = ConvertObjectToString(finishProgrammer);
            this.Apprentice = ConvertObjectToString(apprentice);
        }

        public ProjectModel(string jobNumber, int projectNumber, string name, string customer, DateTime dueDate, string status, string toolMaker, string designer, string roughProgrammer, string finishProgrammer, string electrodProgrammer, string apprentice, string kanBanWorkbookPath, bool overlapAllowed) // Project Creation Constructor. Leaving out status for now.  May add later.
        {
            this.HasProjectInfo = true;
            this.JobNumber = jobNumber;
            this.ProjectNumber = projectNumber;
            this.OldProjectNumber = projectNumber;
            this.Name = name;
            this.Customer = customer;
            this.DueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
            this.Status = status;
            this.Designer = designer;
            this.ToolMaker = toolMaker;
            this.RoughProgrammer = roughProgrammer;
            this.ElectrodeProgrammer = electrodProgrammer;
            this.FinishProgrammer = finishProgrammer;
            this.Apprentice = apprentice;
            this.ComponentList = new List<ComponentModel>();
            this.KanBanWorkbookPath = kanBanWorkbookPath;
            this.OverlapAllowed = overlapAllowed;
        }

        public ProjectModel(DateTime dueDate, string toolMaker, string designer, string roughProgrammer, string finishProgrammer, string electrodeProgrammer, string kanBanWorkbookPath) // Project Data Retrieval Constructor.
        {
            this.HasProjectInfo = true;
            this.DueDate = dueDate;
            this.Designer = designer;
            this.ToolMaker = toolMaker;
            this.RoughProgrammer = roughProgrammer;
            this.ElectrodeProgrammer = electrodeProgrammer;
            this.FinishProgrammer = finishProgrammer;
            this.KanBanWorkbookPath = kanBanWorkbookPath;
        }

        public void SetHasProjectInfo(bool hasInfo)
        {
            this.HasProjectInfo = hasInfo;
        }

        public void SetProjectNumber(string projectNumber)
        {
            bool isInteger = int.TryParse(projectNumber, out int n);

            if(isInteger)
            {
                this.ProjectNumber = n;
            }
            else
            {
                MessageBox.Show("Project Number needs to be a whole number.");
            }
        }

        public void SetOldProjectNumber(int oldProjectNumber)
        {
            this.OldProjectNumber = oldProjectNumber;
        }

        public void SetProjectInfo(string jobNumber, string projectNumber, DateTime dueDate, object toolMaker, object designer, object roughProgrammer, object electrodeProgrammer, object finishProgrammer)
        {
            int projectNumberResult;
            this.HasProjectInfo = true;
            this.JobNumber = jobNumber;
            if(projectNumber != this.OldProjectNumber.ToString())
            {
                this.ProjectNumberChanged = true;
            }

            if(int.TryParse(projectNumber, out projectNumberResult))
            {
                
            }
            else
            {

            }

            this.ProjectNumber = projectNumberResult;
            this.DueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
            this.ToolMaker = ConvertObjectToString(toolMaker);
            this.Designer = ConvertObjectToString(designer);
            this.RoughProgrammer = ConvertObjectToString(roughProgrammer);
            this.ElectrodeProgrammer = ConvertObjectToString(electrodeProgrammer);
            this.FinishProgrammer = ConvertObjectToString(finishProgrammer);
        }

        public void SetProjectInfo(object jobNumber, object projectNumber, object dueDate, object toolMaker, object designer, object roughProgrammer, object electrodeProgrammer, object finishProgrammer)
        {
            this.HasProjectInfo = true;
            this.JobNumber = ConvertObjectToString(jobNumber);
            this.ProjectNumber = ConvertObjectToInt(projectNumber);
            this.DueDate = ConvertObjectToDateTime(dueDate);
            this.ToolMaker = ConvertObjectToString(toolMaker);
            this.Designer = ConvertObjectToString(designer);
            this.RoughProgrammer = ConvertObjectToString(roughProgrammer);
            this.ElectrodeProgrammer = ConvertObjectToString(electrodeProgrammer);
            this.FinishProgrammer = ConvertObjectToString(finishProgrammer);
        }

        public void SetProjectDueDate(DateTime dueDate)
        {
            this.DueDate = dueDate;
        }

        public bool AddComponent(string name)
        {
            if(!ComponentNameExists(name))
            {
                if (name.Length > ComponentModel.ComponentCharacterLimit)
                {
                    MessageBox.Show($"Component: '{name}' is greater than {ComponentModel.ComponentCharacterLimit} characters. \n\nPlease shorten name.");
                    return false;
                }

                ComponentList.Add(new ComponentModel(name));
                return true;
            }
            else
            {
                return false;
            }

            //printComponentList();
        }

        public bool AddComponent(ComponentModel component)
        {
            if (!ComponentNameExists(component.Name))
            {
                ComponentList.Add(component);

                return true;
            }
            else
            {
                return false;
            }

            //printComponentList();
        }

        public void AddComponentList(List<ComponentModel> componentList)
        {
            ComponentList = new List<ComponentModel>();

            this.ComponentList = componentList;
        }

        public void RemoveComponent(string name)
        {
            ComponentModel component = ComponentList.Where(x => x.Name == name).First();
            ComponentList.Remove(component);

            //printComponentList();
        }

        public void SetQuoteInfo(QuoteModel quoteInfo)
        {
            QuoteInfo = quoteInfo;
        }

        public void MoveComponentUp(int promotedComponentIndex)
        {
            ComponentModel promotedComponent;

            if (promotedComponentIndex > 0)
            {
                promotedComponent = ComponentList.ElementAt(promotedComponentIndex);
                
                ComponentList.RemoveAt(promotedComponentIndex);
                ComponentList.Insert(promotedComponentIndex - 1, promotedComponent);
            }
            else
            {
                MessageBox.Show("Cannot move component any higher.");
            }

            //printComponentList();
        }

        public void MoveComponentDown(int demotedComponentIndex)
        {
            ComponentModel demotedComponent;

            if (demotedComponentIndex < ComponentList.Count - 1)
            {
                demotedComponent = ComponentList.ElementAt(demotedComponentIndex);

                ComponentList.RemoveAt(demotedComponentIndex);
                ComponentList.Insert(demotedComponentIndex + 1, demotedComponent);
            }
            else
            {
                MessageBox.Show("Cannot move component any lower.");
            }

            //printComponentList();
        }

        public bool ComponentNameExists(string name)
        {
            // Perhaps try and find a linq statement that does this.
            foreach(ComponentModel component in ComponentList)
            {
                if(component.Name == name)
                {
                    MessageBox.Show("Component name already exists.");
                    return true;
                }
            }

            return false;
        }

        private void PrintComponentList()
        {
            foreach (ComponentModel component in ComponentList)
            {
                Console.WriteLine(component.Name);
            }

            Console.WriteLine("");
        }

        public void SetDefaultCopiedProjectInfo(int projectNumber)
        {
            this.JobNumber = "XXXXXX";
            this.ProjectNumber = projectNumber;
            this.DueDate = DateTime.Today;
            this.Designer = "";
            this.ToolMaker = "";
            this.RoughProgrammer = "";
            this.ElectrodeProgrammer = "";
            this.FinishProgrammer = "";
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

        private int ConvertObjectToInt(object obj)
        {
            if (obj != null && obj.ToString() != "")
            {
                return Convert.ToInt32(obj);
            }
            else
            {
                return 0;
            }
        }

        private DateTime ConvertObjectToDateTime(object obj)
        {
            DateTime dueDate;

            if (obj.ToString() == "")
            {
                dueDate = DateTime.Today;
                dueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
            }
            else
            {
                dueDate = Convert.ToDateTime(obj);
            }

            return new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
        }
    }
}
