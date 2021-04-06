using System;
using ClassLibrary;

namespace Toolroom_Project_Viewer
{
    public partial class MonthlyXtraSchedulerReport : DevExpress.XtraScheduler.Reporting.XtraSchedulerReport
    {
        public ProjectModel Project { get; set; }
        public MonthlyXtraSchedulerReport()
        {
            InitializeComponent();
        }

        public MonthlyXtraSchedulerReport(ProjectModel project, string component)
        {
            InitializeComponent();
            Project = project;
            xrTableCell2.Text = $"{Project.JobNumber} - #{Project.ProjectNumber}";
            xrTableCell3.Text = $"{component}";
            xrTableCell6.Text = $"{Project.DatePulled}";
        }
    }
}
