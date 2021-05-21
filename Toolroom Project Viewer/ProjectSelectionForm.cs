using System;
using System.ComponentModel;
using System.Windows.Forms;
using ClassLibrary;

namespace Toolroom_Project_Viewer
{
    public partial class ProjectSelectionForm : DevExpress.XtraEditors.XtraForm
    {
        public ProjectModel Project { get; set; } = new ProjectModel();

        public ProjectSelectionForm()
        {
            InitializeComponent();
        }

        private void ProjectSelectionForm_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'workload_Tracking_System_DBDataSet.WorkLoad' table. You can move, or remove it, as needed.
            //this.workLoadTableAdapter.Fill(this.workload_Tracking_System_DBDataSet.WorkLoad);

            BindingList<WorkLoadModel> workloads = new BindingList<WorkLoadModel>(Database.GetWorkloads());

            gridControl1.DataSource = workloads;
        }

        private void UseSelectedProjectButton_Click(object sender, EventArgs e)
        {
            if (bandedGridView1.FocusedRowHandle >= 0)
            {
                Project = new ProjectModel(
                                            jobNumber: bandedGridView1.GetFocusedRowCellValue("ToolNumber"),
                                            projectNumber: bandedGridView1.GetFocusedRowCellValue("ProjectNumber"),
                                            mwoNumber: bandedGridView1.GetFocusedRowCellValue("MWONumber"),
                                            customer: bandedGridView1.GetFocusedRowCellValue("Customer"),
                                            project: bandedGridView1.GetFocusedRowCellValue("Project"),
                                            dueDate: bandedGridView1.GetFocusedRowCellValue("FinishDate"),
                                            toolMaker: bandedGridView1.GetFocusedRowCellValue("ToolMaker"),
                                            designer: bandedGridView1.GetFocusedRowCellValue("Designer"),
                                            roughProgrammer: bandedGridView1.GetFocusedRowCellValue("RoughProgrammer"),
                                            electrodProgrammer: bandedGridView1.GetFocusedRowCellValue("ElectrodeProgrammer"),
                                            finishProgrammer: bandedGridView1.GetFocusedRowCellValue("FinishProgrammer"),
                                            apprentice: bandedGridView1.GetFocusedRowCellValue("Apprentice")
                                         );

                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("Please select a project from the grid.");
            }
        }
    }
}