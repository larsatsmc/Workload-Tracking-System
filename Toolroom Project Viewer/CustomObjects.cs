﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Toolroom_Project_Viewer
{
    #region #customappointment
    public class CustomAppointment
    {
        private int m_AppointmentId;
        private DateTime m_Start;
        private DateTime m_End;
        private string m_Subject;
        private int m_Status;
        private string m_Description;
        private int m_Label;
        private string m_Location;
        private bool m_Allday;
        private int m_EventType;
        private string m_RecurrenceInfo;
        private string m_ReminderInfo;
        private object m_OwnerId;
        private int m_TaskId;

        
        public int AppointmentId { get { return m_AppointmentId; } set { m_AppointmentId = value; } }
        public DateTime StartDate { get { return m_Start; } set { m_Start = value; } }
        public DateTime FinishDate { get { return m_End; } set { m_End = value; } }
        public string Subject { get { return m_Subject; } set { m_Subject = value; } }
        public int Status { get { return m_Status; } set { m_Status = value; } }
        public string Description { get { return m_Description; } set { m_Description = value; } }
        public int Label { get { return m_Label; } set { m_Label = value; } }
        public string Location { get { return m_Location; } set { m_Location = value; } }
        public bool AllDay { get { return m_Allday; } set { m_Allday = value; } }
        public int EventType { get { return m_EventType; } set { m_EventType = value; } }
        public string RecurrenceInfo { get { return m_RecurrenceInfo; } set { m_RecurrenceInfo = value; } }
        public string ReminderInfo { get { return m_ReminderInfo; } set { m_ReminderInfo = value; } }
        public object OwnerId { get { return m_OwnerId; } set { m_OwnerId = value; } }
        public int TaskId { get { return m_TaskId; } set { m_TaskId = value; } }

        public CustomAppointment()
        {
        }
    }
    #endregion  #customappointment

    #region #customresource
    public class CustomResource
    {
        private string m_name;
        private int m_res_id;
        private int m_par_id;

        public string Name { get { return m_name; } set { m_name = value; } }
        public int ResID { get { return m_res_id; } set { m_res_id = value; } }
        public int ParentID { get { return m_par_id; } set { m_par_id = value; } }

        public CustomResource()
        {
        }
    }
    #endregion #customresource

    #region #customdependency
    public class CustomDependency
    {
        private int dep_id;
        private int par_id;

        public int DepID { get { return dep_id; } set { dep_id = value; } }
        public int ParentID { get { return par_id; } set { par_id = value; } }

        public CustomDependency()
        {
        }
    }
    #endregion #customdependency
}
