using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClassLibrary;
using Xunit;

namespace DemoLibrary.Tests
{
    public class TemplateTests
    {
        // "X:\TOOLROOM\Workload Tracking System\Templates\100000-TEST - #1.txt"
        [Fact]
        public void ReadProjectFromTextFile_InvalidNameShouldFail()
        {
            // Arrange
            Template template = new Template();
            //ProjectInfo project = template.ReadProjectFromTextFile(@"");

            // Act 


            // Assert
            Assert.Throws<FileNotFoundException>(() => template.ReadProjectFromTextFile(@"X:\TOOLROOM\Workload Tracking System\Templates\test.txt"));
        }

        [Fact]
        public void ReadProjectFromTextFile_EmptyNameShouldFail()
        {
            // Arrange
            Template template = new Template();
            //ProjectInfo project = template.ReadProjectFromTextFile(@"");

            // Act 

            // Assert
            Assert.Throws<ArgumentException>(() => template.ReadProjectFromTextFile(@""));
        }

        [Fact]
        public void ReadProjectFromTextFile_TestFileReadShouldWork()
        {
            // Arrange
            Template template = new Template();

            // Act 
            ProjectModel project = template.ReadProjectFromTextFile(@"X:\TOOLROOM\Workload Tracking System\Templates\100000-TEST - #1.txt");

            // Assert
            Assert.True(project.JobNumber == "100000-TEST", "Job Number incorrect.");
            Assert.True(project.ProjectNumber == 1, "Project Number incorrect");
            Assert.True(project.DueDate == new DateTime(2018,6,15), "Due Date incorrect.");
            Assert.True(project.ToolMaker == "Barry Black", "Tool Maker incorrect.");
            Assert.True(project.Designer == "Phil Morris", "Designer incorrect");
            Assert.True(project.RoughProgrammer == "Micah Bruns", "Rough programmer incorrect.");
            Assert.True(project.ElectrodeProgrammer == "Rod Shilts", "Electrode programmer incorrect.");
            Assert.True(project.FinishProgrammer == "Alex Anderson", "Finish programmer incorrect.");

            Assert.True(project.ComponentList.Count == 2, "Incorrect number of components read.");

            Assert.True(project.ComponentList[0].Name == "A-Cavity", "Component name incorrect.");
            Assert.True(project.ComponentList[0].Quantity == 2, "Quantity incorrect.");
            Assert.True(project.ComponentList[0].Spares == 1, "Spares incorrect.");
            Assert.True(project.ComponentList[0].Material == "S7", "Material incorrect.");
            Assert.True(project.ComponentList[0].Finish == "80 Grit", "Finish incorrect.");
            Assert.True(project.ComponentList[0].Notes == "Make it nice.", "Notes incorrect.");

            Assert.True(project.ComponentList[0].TaskList.Count == 9, "Incorrect number of tasks read.");

            Assert.Equal(5, project.ComponentList[0].TaskList[0].Hours);
            Assert.Equal("1 Day(s)", project.ComponentList[0].TaskList[0].Duration);
            Assert.Equal("None", project.ComponentList[0].TaskList[0].Machine);
            Assert.Equal("Micah Bruns", project.ComponentList[0].TaskList[0].Personnel);
            Assert.Equal("", project.ComponentList[0].TaskList[0].Predecessors);
            Assert.Equal("Leave .01\" stock.", project.ComponentList[0].TaskList[0].Notes);

            //Assert.True(project.ComponentList[0].TaskList[0].Duration == "1 Day(s)", "Task Duration incorrect.");
            //Assert.True(project.ComponentList[0].TaskList[0].Machine == "None", "Machine incorrect.");
            //Assert.True(project.ComponentList[0].TaskList[0].Personnel == "Micah Bruns", "Personnel incorrect.");
            //Assert.True(project.ComponentList[0].TaskList[0].Predecessors == "", "Predecessors incorrect.");
            //Assert.True(project.ComponentList[0].TaskList[0].Notes == "Leave .01\" stock.", "Notes incorrect.");
        }
    }
}
