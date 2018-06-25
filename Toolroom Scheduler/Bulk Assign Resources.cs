using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Toolroom_Scheduler
{
    public partial class Bulk_Assign_Resources_Form : Form
    {
        Database db;
        List<string> DesignerList = new List<string> { "Phil Morris", "Brian Yoder", "Lee Meservey", "Jim Schmidt", " " };
        List<string> ProgrammerList = new List<string> { "Josh Meservey", "Shawn Swiggum", "Alex Anderson", "Rod Shilts", "Ben Meservey", "Derek Timm", "Micah Bruins", " " };

        public string Designer {get;private set;}
        public string RoughProgrammer {get;private set;}
        public string FinishProgrammer {get;private set;}
        public string ElectrodeProgrammer {get;private set;}

        public Bulk_Assign_Resources_Form()
        {
            InitializeComponent();
        }

		private void populateComboBox(ComboBox cb)
		{
			Database db = new Database();

			if (cb.Name == "ToolMakerComboBox")
			{
				cb.DataSource = db.GetResourceList("Tool Maker");
			}
			else if (cb.Name == "DesignerComboBox")
			{
				cb.DataSource = db.GetResourceList("Designer");
			}
			else if (cb.Name == "RoughProgrammerComboBox")
			{
				cb.DataSource = db.GetResourceList("Rough Programmer");
			}
			else if (cb.Name == "FinishProgrammerComboBox")
			{
				cb.DataSource = db.GetResourceList("Finish Programmer");
			}
			else if (cb.Name == "ElectrodeProgrammerComboBox")
			{
				cb.DataSource = db.GetResourceList("Electrode Programmer");
			}
		}

		private void CancelButton_Click(object sender, EventArgs e)
        {
            
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            Designer = DesignerComboBox.Text;
            RoughProgrammer = RoughProgrammerComboBox.Text;
            FinishProgrammer = FinishProgrammerComboBox.Text;
            ElectrodeProgrammer = ElectrodeProgrammerComboBox.Text;
        }

        private void Bulk_Assign_Resources_Form_Load(object sender, EventArgs e)
        {
            DesignerComboBox.DataSource = DesignerList.ToList();
            RoughProgrammerComboBox.DataSource = ProgrammerList.ToList();
            FinishProgrammerComboBox.DataSource = ProgrammerList.ToList();
            ElectrodeProgrammerComboBox.DataSource = ProgrammerList.ToList();

            DesignerComboBox.SelectedItem = " ";
            RoughProgrammerComboBox.SelectedItem = " ";
            FinishProgrammerComboBox.SelectedItem = " ";
            ElectrodeProgrammerComboBox.SelectedItem = " ";
        }

		private void DesignerComboBox_DropDown(object sender, EventArgs e)
		{
			populateComboBox((ComboBox)sender);
		}

		private void RoughProgrammerComboBox_DropDown(object sender, EventArgs e)
		{
			populateComboBox((ComboBox)sender);
		}

		private void FinishProgrammerComboBox_DropDown(object sender, EventArgs e)
		{
			populateComboBox((ComboBox)sender);
		}

		private void ElectrodeProgrammerComboBox_DropDown(object sender, EventArgs e)
		{
			populateComboBox((ComboBox)sender);
		}
	}
}
