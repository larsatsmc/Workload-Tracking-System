﻿using ClassLibrary;
using DevExpress.XtraReports.UI;
using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;

namespace Toolroom_Project_Viewer
{
    public partial class KanBanXtraReport : DevExpress.XtraReports.UI.XtraReport
    {
        public KanBanXtraReport()
        {
            InitializeComponent();
        }

        public KanBanXtraReport(ProjectModel project)
        {
            InitializeComponent();

            objectDataSource1.DataSource = project;
        }
    }
}
