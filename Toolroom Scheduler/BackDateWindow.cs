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
    public partial class BackDateWindow : Form
    {
        public DateTime DueDate { get; set; }
        public DateTime BackDate { get; private set; }

        public BackDateWindow(DateTime date)
        {
            InitializeComponent();
            DueDate = date;
            label1.Text = "Due Date: " + DueDate.ToShortDateString();
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            Database db = new Database();
            BackDate = db.SubtractBusinessDays(DueDate, daysNumericUpDown.Value.ToString() + " Day(s)");
        }
    }
}
