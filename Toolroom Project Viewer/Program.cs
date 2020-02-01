using System;
using System.Windows.Forms;

namespace Toolroom_Project_Viewer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainWindow());
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
            }
        }
    }
}
