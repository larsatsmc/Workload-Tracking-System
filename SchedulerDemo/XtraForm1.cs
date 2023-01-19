using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.Drawing;
using System.Diagnostics;

namespace SchedulerDemo
{
    public partial class XtraForm1 : DevExpress.XtraEditors.XtraForm
    {
        private Appointment DraggedAppointment;
        private SchedulerHitInfo HitInfo;
        public XtraForm1()
        {
            InitializeComponent();
        }

        private void XtraForm1_Load(object sender, EventArgs e)
        {
            //Appointment apt = 
        }

        private void schedulerControl1_PopupMenuShowing(object sender, PopupMenuShowingEventArgs e)
        {
            //Point pos = schedulerControl1.PointToClient(Cursor.Position);
            //SchedulerViewInfoBase viewInfo = schedulerControl1.ActiveView.ViewInfo;
            //HitInfo = viewInfo.CalcHitInfo(pos, false);

            e.Menu.Items.Remove(e.Menu.Items.FirstOrDefault(x => x.Caption == "&Copy"));

            if (e.Menu.Items.Count(x => x.Caption == "Move Subsequent Component Tasks with Lock Spacing") == 0)
            {
                e.Menu.Items.Insert(1, new SchedulerMenuItem("Move Subsequent Component Tasks with Lock Spacing", schedulerControl_MoveAndDoSomeOtherStuff));
            }
        }

        private void schedulerControl_MoveAndDoSomeOtherStuff(object sender, EventArgs e)
        {
            //MessageBox.Show($"Final Start: {HitInfo.ViewInfo.Interval.Start.ToShortDateString()} Finish: {HitInfo.ViewInfo.Interval.End.ToShortDateString()}");

            Debug.WriteLine($"Init. Start: {DraggedAppointment.Start.ToShortDateString()} Finish: {DraggedAppointment.End.ToShortDateString()}");
            Debug.WriteLine($"Final Start: {HitInfo.ViewInfo.Interval.Start.ToShortDateString()} Finish: {HitInfo.ViewInfo.Interval.End.ToShortDateString()}");
            DraggedAppointment.Start = HitInfo.ViewInfo.Interval.Start;
            DraggedAppointment.End = HitInfo.ViewInfo.Interval.End;
        }

        private void schedulerControl1_MouseDown(object sender, MouseEventArgs e)
        {
            var scheduler = sender as DevExpress.XtraScheduler.SchedulerControl;
            var hitInfo = scheduler.ActiveView.CalcHitInfo(e.Location, false);

            if (e.Button == MouseButtons.Right)
            {
                //RightMouseButtonPressed = true;

                if (hitInfo.HitTest == SchedulerHitTest.AppointmentContent)
                {
                    Appointment apt = ((AppointmentViewInfo)hitInfo.ViewInfo).Appointment;
                    DraggedAppointment = apt;
                }
            }
        }

        private void schedulerControl1_DragOver(object sender, DragEventArgs e)
        {
            var scheduler = sender as DevExpress.XtraScheduler.SchedulerControl;
            Point pos = schedulerControl1.PointToClient(Cursor.Position);
            HitInfo = scheduler.ActiveView.CalcHitInfo(pos, false);

            if (HitInfo.HitTest == SchedulerHitTest.AppointmentContent)
            {
                //MessageBox.Show($"Final Start: {hitInfo.ViewInfo.Interval.Start.ToShortDateString()} Finish: {hitInfo.ViewInfo.Interval.End.ToShortDateString()}");
                Debug.WriteLine($"Final Start: {HitInfo.ViewInfo.Interval.Start.ToShortDateString()} Finish: {HitInfo.ViewInfo.Interval.End.ToShortDateString()}");
            }

        }
    }
}