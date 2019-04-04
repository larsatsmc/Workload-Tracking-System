using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Toolroom_Project_Viewer
{
    public partial class ForwardDateWindow : DevExpress.XtraEditors.XtraForm
    {
        public DateTime ForwardDate { get; private set; }

        public ForwardDateWindow(string title, DateTime preselectedDate)
        {
            InitializeComponent();

            this.Text = title;

            if (preselectedDate != null)
            {
                calendarControl.DateTime = preselectedDate;
            }
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            ForwardDate = calendarControl.DateTime;
        }
    }
}
