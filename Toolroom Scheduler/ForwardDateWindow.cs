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
    public partial class ForwardDateWindow : Form
    {
        public DateTime ForwardDate { get; private set; }

        public ForwardDateWindow()
        {
            InitializeComponent();
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            ForwardDate = calendarControl1.DateTime;
        }
    }
}
