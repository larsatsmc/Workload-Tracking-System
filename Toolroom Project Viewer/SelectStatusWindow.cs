using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Toolroom_Project_Viewer
{
    public partial class SelectStatusWindow : Form
    {
        List<string> PersonnelStatuses = new List<string> { "Task Completed", "In-Progress", "Have yet to start", "Hold", "Clear"};
        List<string> ProjectStatuses = new List<string> { "Waiting for customer", "Waiting for engineer", "Job is a go", "Completed & need to run final report" };
        List<object> PersonnelColorList = new List<object> { Color.Gray, Color.LightGreen, Color.Orange, Color.Yellow };
        List<string> OtherStatuses = new List<string> { "Highlight", "Clear" };
        List<object> OtherColorList = new List<object> { Color.Yellow };

        public Color SelectedColor { get; private set; }

        public SelectStatusWindow(string columnType, Color rowColor)
        {
            int i = 0;

            List<string> radioGroupList = new List<string>();

            List<object> colorList = new List<object>();

            InitializeComponent();

            radioGroup1.Properties.Items.Clear();

            colorList.Clear();

            if (columnType == "Personnel")
            {
                radioGroupList = PersonnelStatuses;
                colorList = PersonnelColorList;
            }
            else if (columnType == "Other")
            {
                radioGroupList = OtherStatuses;
                colorList = OtherColorList;
            }

            colorList.Add(rowColor);

            foreach (string status in radioGroupList)
            {
                radioGroup1.Properties.Items.Add(new DevExpress.XtraEditors.Controls.RadioGroupItem(colorList.ElementAt(i++), status));
            }
        }

        private void radioGroup1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedColor = (Color)radioGroup1.Properties.Items[radioGroup1.SelectedIndex].Value;

            this.DialogResult = DialogResult.OK;
        }
    }
}
