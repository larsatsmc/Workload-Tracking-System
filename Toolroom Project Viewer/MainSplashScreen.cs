using System;
using System.Diagnostics;
using DevExpress.XtraSplashScreen;

namespace Toolroom_Project_Viewer
{
    public partial class MainSplashScreen : SplashScreen
    {
        public MainSplashScreen()
        {
            InitializeComponent();
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            this.labelCopyright.Text = $"v.{fvi.FileVersion} Copyright © 1998-{DateTime.Now.Year}";
        }

        #region Overrides

        public override void ProcessCommand(Enum cmd, object arg)
        {
            base.ProcessCommand(cmd, arg);
        }

        #endregion

        public enum SplashScreenCommand
        {
        }
    }
}