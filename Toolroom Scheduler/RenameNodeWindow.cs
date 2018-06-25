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
    public partial class RenameNodeWindow : Form
    {
        public RenameNodeWindow()
        {
            InitializeComponent();
        }

        public void setPrefilledText(string prefilledText)
        {
            this.PrefilledText = prefilledText; 
        }

        public string PrefilledText { get; private set; }

        private void RenameNodeWindow_Load(object sender, EventArgs e)
        {
            NewNameTextBox.Text = this.PrefilledText;
        }
    }
}
