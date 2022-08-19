using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Toolroom_Project_Viewer;
using ClassLibrary;
using Xunit;
using System.Data;

namespace MainWindow.Tests
{
    public class ProjectCreationFormTests
    {
        private global::DevExpress.XtraScheduler.SchedulerStorage schedulerStorage;
        [Fact]
        public void SetPersonnel_ShouldWork()
        {
            ProjectModel project = Template.ReadFromXmlFile<ProjectModel>(@"X:\TOOLROOM\Workload Tracking System\Templates\100000-TEST - #12.xml");

            ProjectCreationForm pcf = new ProjectCreationForm(project, schedulerStorage);

            System.Windows.Forms.ComboBox combo = new System.Windows.Forms.ComboBox();

            combo.Name = "RoughProgrammerComboBox";

            combo.Text = "Ethan Brey";

            pcf.SetPersonnnel(combo);

            foreach (var component in project.Components)
            {                
                if (component.Component == "A-Angle Pin")
                {
                    Assert.True(component.Tasks.Find(x => x.TaskName == "Program Rough").Personnel == "Micah Bruns", $"Rough Programmer Incorrect for {component.Component} is '{component.Tasks.Find(x => x.TaskName == "Program Rough").Personnel}'");
                }
                else
                {
                    Assert.True(component.Tasks.Find(x => x.TaskName == "Program Rough").Personnel == "Ethan Brey", $"Rough Programmer Incorrect for {component.Component} is '{component.Tasks.Find(x => x.TaskName == "Program Rough").Personnel}'");
                }
            }
        }
        [Fact]
        public void FindMatchingDepartment_ShouldWork()
        {

            Assert.True(GeneralOperations.FindMatchingDepartment("RoughProgrammerComboBox") == "Program Rough", $"Failed to select Program Rough Department.");
            Assert.True(GeneralOperations.FindMatchingDepartment("ElectrodeProgrammerComboBox") == "Program Electrodes", $"Failed to select Program Electrodes Department.");
            Assert.True(GeneralOperations.FindMatchingDepartment("FinishProgrammerComboBox") == "Program Finish", $"Failed to select Program Finish Department.");
            Assert.True(GeneralOperations.FindMatchingDepartment("EDMSinkerOperatorComboBox") == "EDM Sinker", $"Failed to select EDM Sinker Department.");
            Assert.True(GeneralOperations.FindMatchingDepartment("RoughCNCOperatorComboBox") == "CNC Rough", $"Failed to select CNC Rough Department.");
            Assert.True(GeneralOperations.FindMatchingDepartment("ElectrodeCNCOperatorComboBox") == "CNC Electrodes", $"Failed to select CNC Electrodes Department.");
            Assert.True(GeneralOperations.FindMatchingDepartment("FinishCNCOperatorComboBox") == "CNC Finish", $"Failed to select CNC Finish Department.");
            Assert.True(GeneralOperations.FindMatchingDepartment("EDMWireOperaterComboBox") == "EDM Wire (In-House)", $"Failed to select EDM Wire Department.");
        }
    }
}
