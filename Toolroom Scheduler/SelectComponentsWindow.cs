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
    public partial class SelectComponentsWindow : Form
    {
        public List<string> ComponentList { get; private set; }

        public SelectComponentsWindow()
        {
            InitializeComponent();
        }

        public SelectComponentsWindow(List<Component> componentList)
        {
            InitializeComponent();

            foreach (Component component in componentList)
            {
                ComponentCheckedListBox.Items.Add(component.Name);
            }
        }

        public SelectComponentsWindow(List<string> componentList)
        {
            InitializeComponent();

            foreach (string component in componentList)
            {
                ComponentCheckedListBox.Items.Add(component);
            }
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            ComponentList = new List<string>();

            foreach (var item in ComponentCheckedListBox.CheckedItems)
            {
                ComponentList.Add(item.ToString());
            }
        }
    }
}
