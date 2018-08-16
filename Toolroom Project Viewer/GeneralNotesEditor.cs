using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Toolroom_Project_Viewer
{
    public partial class GeneralNotesEditor : DevExpress.XtraEditors.XtraForm
    {
        public string RTFText { get; private set; }

        public GeneralNotesEditor(string rtfText)
        {
            InitializeComponent();
            richEditControl1.Text = rtfText;
            //richEditControl1.RtfText = rtfText;
        }

        private void richEditControl1_TextChanged(object sender, EventArgs e)
        {
            this.RTFText = richEditControl1.RtfText;
            Console.WriteLine(RTFText);
            Console.WriteLine(richEditControl1.Text);
        }

        private void GeneralNotesEditor_Load(object sender, EventArgs e)
        {
            
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void textEdit1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}