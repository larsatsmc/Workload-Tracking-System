using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using ClassLibrary;
using DevExpress.XtraEditors;
using DevExpress.Xpf.Grid;

namespace Toolroom_Project_Viewer
{
    class Database
    {
        static readonly string ConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=X:\TOOLROOM\Workload Tracking System\Database\Workload Tracking System DB.accdb";
        OleDbConnection Connection = new OleDbConnection(ConnString);

        DataTable taskIDKey = new DataTable();

        public DataTable GetAppointmentData()
        {
            DataTable dt = new DataTable();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            //string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resources FROM Tasks WHERE TaskName LIKE 'CNC Finish'";
            string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resource, ToolMaker, Notes FROM Tasks WHERE TaskName = 'CNC Rough'";

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            try
            {
                adapter.Fill(dt);
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message, "OledbException Error");
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message, "getAppointmentsData");
            }

            return dt;
        }

        public DataTable GetAppointmentData(string department)
        {
            DataTable dt = new DataTable();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            //string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resources FROM Tasks WHERE TaskName LIKE 'CNC Finish'";
            string queryString = SetQueryString(department);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);
            int i = 1;

            //adapter.SelectCommand.Parameters.AddWithValue("@department", setQueryString(department);

            try
            {
                adapter.Fill(dt);

                //foreach (DataRow nrow in dt.Rows)
                //{
                //    nrow["ID"] = i++;
                //    if(nrow["Resource"].ToString() == "")
                //    {
                //        nrow["Resource"] = "None";
                //    }
                //    Console.WriteLine($"{nrow["ID"]} {nrow["Subject"]} {nrow["Location"]} {nrow["StartDate"]}");
                //}
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message, "OledbException Error");
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message, "getAppointmentsData");
            }
            

            return dt;
        }

        public DataTable GetAppointmentData(string jobNumber, int projectNumber)
        {
            DataTable dt = new DataTable();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            //string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resources FROM Tasks WHERE TaskName LIKE 'CNC Finish'";
            string queryString = "SELECT TaskID, JobNumber & ' ' & Component & ' ' & TaskName As Subject, TaskName, StartDate, FinishDate, Machine, ToolMaker, Notes " +
                                 "FROM Tasks " +
                                 "WHERE JobNumber LIKE @jobNumber AND ProjectNumber LIKE @projectNumber";

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            adapter.SelectCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

            adapter.Fill(dt);

            return dt;
        }

        public bool UpdateTask(string jobNumber, int projectNumber, string component, int taskID, DateTime startDate, DateTime finishDate)
        {
            try
            {
                string queryString = "UPDATE Tasks SET StartDate = @startDate, FinishDate = @finishDate " +
                                     "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";
                OleDbConnection Connection = new OleDbConnection(ConnString);
                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

                adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

                string predecessors = GetTaskPredecessors(jobNumber, projectNumber, component, taskID);

                if (predecessors != "" && GetLatestPredecessorFinishDate(jobNumber, projectNumber, component, predecessors) > startDate)
                {
                    MessageBox.Show("You cannot put a task start date before its predecessor's finish date.");
                    return false;
                }

                adapter.UpdateCommand.Parameters.AddWithValue("@startDate", startDate);
                adapter.UpdateCommand.Parameters.AddWithValue("@finishDate", finishDate);
                adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@component", component);
                adapter.UpdateCommand.Parameters.AddWithValue("@taskID", taskID);

                Connection.Open();
                adapter.UpdateCommand.ExecuteNonQuery();
                Connection.Close();

                Connection.Dispose();

                MoveDescendents(jobNumber, projectNumber, component, finishDate, taskID);
                return true;
            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message);
                return false;
            }
        }

        public (string jobNumber, int projectNumber, int taskID, string predecessors) GetTaskInfo(int id)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable dt = new DataTable();
            string jobNumber = "";
            string predecessors = "";
            int projectNumber = 0;
            int taskID = 0;
            string queryString;

            queryString = "SELECT * FROM Tasks WHERE ID = @id";
            OleDbConnection Connection = new OleDbConnection(ConnString);
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            adapter.SelectCommand.Parameters.AddWithValue("@id", id);

            try
            {
                Connection.Open();
                using (var reader = adapter.SelectCommand.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        int ord = reader.GetOrdinal("JobNumber");
                        jobNumber = reader.GetString(ord);

                        ord = reader.GetOrdinal("ProjectNumber");
                        projectNumber = reader.GetInt32(ord);

                        ord = reader.GetOrdinal("TaskID");
                        taskID = reader.GetInt32(ord);

                        ord = reader.GetOrdinal("Predecessors");
                        predecessors = reader.GetString(ord);
                    }
                }
                Connection.Close();
                Connection.Dispose();
            }
            catch(Exception e)
            {
                Connection.Close();
                Connection.Dispose();
                MessageBox.Show(e.Message);
            }

            //MessageBox.Show(projectInfo.Item1 + " " + projectInfo.Item2);

            return (jobNumber, projectNumber, taskID, predecessors);
        }

        public string GetTaskPredecessors(string jobNumber, int projectNumber, string component, int taskID)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable dt = new DataTable();
            string predecessors = "";
            string queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";

            OleDbConnection Connection = new OleDbConnection(ConnString);

            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            adapter.SelectCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            adapter.SelectCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
            adapter.SelectCommand.Parameters.AddWithValue("@component", component);
            adapter.SelectCommand.Parameters.AddWithValue("@taskID", taskID);

            try
            {
                Connection.Open();

                using (var reader = adapter.SelectCommand.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        int ord = reader.GetOrdinal("Predecessors");
                        predecessors = reader.GetString(ord);
                    }
                }
                Connection.Close();
                Connection.Dispose();
            }
            catch (Exception e)
            {
                Connection.Close();
                Connection.Dispose();
                MessageBox.Show(e.Message);
            }

            //MessageBox.Show(projectInfo.Item1 + " " + projectInfo.Item2);

            return predecessors;
        }

        public DateTime GetFinishDate(string jobNumber, int projectNumber, string component, int taskID)
        {
            DateTime FinishDate = DateTime.Today;
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbCommand sqlCommand = new OleDbCommand("SELECT FinishDate from Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID", Connection);

            sqlCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            sqlCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
            sqlCommand.Parameters.AddWithValue("@component", component);
            sqlCommand.Parameters.AddWithValue("@taskID", taskID);

            try
            {
                Connection.Open();

                FinishDate = (DateTime)sqlCommand.ExecuteScalar();

                Connection.Close();
                Connection.Dispose();
            }
            catch
            {
                Connection.Close();
                Connection.Dispose();
                MessageBox.Show("A predecessor has no finish date.");
            }

            return FinishDate;
        }

        private DateTime GetTaskFinishDate(DataTable dt, int id)
        {
            DateTime fd = new DateTime(2000, 1, 1);

            foreach (DataRow nrow in dt.Rows)
            {
                //Console.WriteLine(nrow["ID"] + " ");
                if (nrow["TaskID"].ToString() == id.ToString())
                {
                    if (nrow["FinishDate"] != DBNull.Value)
                        fd = Convert.ToDateTime(nrow["FinishDate"]);

                    goto NextStep;
                }
            }
            NextStep:
            return fd;
        }

        private DateTime GetLatestPredecessorFinishDate(string jobNumber, int projectNumber, string component, string predecessors)
        {
            Database db = new Database();
            DateTime? latestFinishDate = null;
            DateTime currentDate;
            string[] predecessorArr;
            string predecessor;

            predecessorArr = predecessors.Split(',');

            foreach (string currPredecessor in predecessorArr)
            {
                predecessor = currPredecessor.Trim();
                currentDate = db.GetFinishDate(jobNumber, projectNumber, component, Convert.ToInt16(predecessor));

                if (latestFinishDate == null || latestFinishDate < currentDate)
                {
                    latestFinishDate = currentDate;
                }
            }

            return (DateTime)latestFinishDate;
        }

        private DateTime AddBusinessDays(DateTime date, string durationSt)
        {
            int days;
            string[] duration = durationSt.Split(' ');
            days = Convert.ToInt16(duration[0]);

            if (days < 0)
            {
                throw new ArgumentException("days cannot be negative", "days");
            }

            if (days == 0) return date;

            if (date.DayOfWeek == DayOfWeek.Saturday)
            {
                date = date.AddDays(2);
                days -= 1;
            }
            else if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = date.AddDays(1);
                days -= 1;
            }

            date = date.AddDays(days / 5 * 7);
            int extraDays = days % 5;

            if ((int)date.DayOfWeek + extraDays > 5)
            {
                extraDays += 2;
            }

            return date.AddDays(extraDays);

        }

        public double GetBusinessDays(DateTime startD, DateTime endD)
        {
            double calcBusinessDays =
                1 + ((endD - startD).TotalDays * 5 -
                (startD.DayOfWeek - endD.DayOfWeek) * 2) / 7;

            if (endD.DayOfWeek == DayOfWeek.Saturday) calcBusinessDays--;
            if (startD.DayOfWeek == DayOfWeek.Sunday) calcBusinessDays--;

            return calcBusinessDays;
        }

        public void ChangeTaskStartDate(string jobNumber, int projectNumber, string component, DateTime currentTaskStartDate, string duration, int taskID)
        {
            try
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter();

                DateTime currentTaskFinishDate = AddBusinessDays(currentTaskStartDate, duration);

                string queryString;

                queryString = "UPDATE Tasks " +
                              "SET StartDate = @startDate, FinishDate = @finishDate " +
                              "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";

                adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

                adapter.UpdateCommand.Parameters.AddWithValue("@startDate", currentTaskStartDate.ToShortDateString());
                adapter.UpdateCommand.Parameters.AddWithValue("@finishDate", currentTaskFinishDate.ToShortDateString());
                adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@component", component);
                adapter.UpdateCommand.Parameters.AddWithValue("@taskID", taskID);


                Connection.Open();

                adapter.UpdateCommand.ExecuteNonQuery();

                MoveDescendents(jobNumber, projectNumber, component, currentTaskFinishDate, taskID);
            }
            catch (OleDbException oleEx)
            {
                throw oleEx;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Connection.Close();
            }
        }

        public void ChangeTaskFinishDate(string jobNumber, int projectNumber, string component, DateTime currentTaskFinishDate, int taskID)
        {
            try
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter();

                string queryString;

                queryString = "UPDATE Tasks " +
                              "SET FinishDate = @finishDate " +
                              "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";

                adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

                adapter.UpdateCommand.Parameters.AddWithValue("@finishDate", currentTaskFinishDate);
                adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@component", component);
                adapter.UpdateCommand.Parameters.AddWithValue("@taskID", taskID);

                Connection.Open();

                adapter.UpdateCommand.ExecuteNonQuery();

                MoveDescendents(jobNumber, projectNumber, component, currentTaskFinishDate, taskID);

            }
            catch (OleDbException oleException)
            {
                throw oleException;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Connection.Close();
            }
        }

        public void MoveDescendents(string jobNumber, int projectNumber, string component, DateTime currentTaskFinishDate, int currentTaskID)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;

            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component ORDER BY TaskID ASC";
            OleDbConnection Connection = new OleDbConnection(ConnString);
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 20).Value = jobNumber;
            adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = projectNumber;
            adapter.SelectCommand.Parameters.Add("@component", OleDbType.VarChar, 30).Value = component;
            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

            Console.WriteLine("Move Descendents");

            adapter.Fill(datatable);

            UpdateStartAndFinishDates(currentTaskID, datatable, currentTaskFinishDate);

            adapter.Update(datatable);
        }

        public void GetTask(int id)
        {
            string taskName;
            try
            {
                string queryString = "SELECT TaskName " +
                                     "FROM Tasks " +
                                     "WHERE ID = @id";
                OleDbConnection Connection = new OleDbConnection(ConnString);
                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

                adapter.SelectCommand = new OleDbCommand(queryString, Connection);

                adapter.SelectCommand.Parameters.AddWithValue("@id", id);

                Connection.Open();
                taskName = (string)adapter.SelectCommand.ExecuteScalar();
                Connection.Close();
                Connection.Dispose();

                MessageBox.Show(taskName);
            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message);
            }
        }

        // This method is not used and is not properly set up.
        private DateTime GetProjectStartDate(string jobNumber, int projectNumber)
        {
            DateTime projectStartDate = new DateTime(2000, 1, 1);
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbCommand sqlCommand = new OleDbCommand("SELECT StartDate from Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND TaskID = @taskID", Connection);

            sqlCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            sqlCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
            sqlCommand.Parameters.AddWithValue("@taskID", 1);

            try
            {
                Connection.Open();

                projectStartDate = (DateTime)sqlCommand.ExecuteScalar();

                Connection.Close();
                Connection.Dispose();
            }
            catch
            {
                Connection.Close();
                Connection.Dispose();
                MessageBox.Show("First task has no start date.");
            }

            return projectStartDate;
        }

        public List<string> GetJobNumberComboList()
        {
            string queryString = "SELECT DISTINCT JobNumber, ProjectNumber FROM Tasks";
            DataTable dt = new DataTable();
            List<string> jobNumberList = new List<string>();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            foreach (DataRow nrow in dt.Rows)
            {
                //Console.WriteLine(nrow["JobNumber"]);
                //
                jobNumberList.Add($"{nrow["JobNumber"].ToString()} - #{nrow["ProjectNumber"].ToString()}");
            }

            return jobNumberList;
        }

        private string SetQueryString(string department)
        {
            string queryString = null;
            string selectStatment = "ID, JobNumber & ' #' & ProjectNumber & ' ' & Component & '-' & TaskID As Subject, TaskName & ' (' & Hours & ' Hours)' As Location, StartDate, FinishDate, Machine, Resource, Notes";
            string orderByStatement = " ORDER BY StartDate ASC";

            if (department == "Design")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Design%'" + orderByStatement;
            }
            else if (department == "Program Rough")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Rough'" + orderByStatement;
            }
            else if (department == "Program Finish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Finish'" + orderByStatement;
            }
            else if (department == "Program Electrodes")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Electrodes'" + orderByStatement;
            }
            else if (department == "CNC Rough")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Rough'" + orderByStatement;
            }
            else if (department == "CNC Finish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Finish'" + orderByStatement;
            }
            else if (department == "CNC Electrodes")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Electrodes'" + orderByStatement;
            }
            else if (department == "EDM Sinker")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'EDM Sinker'" + orderByStatement;
            }
            else if (department == "Inspection")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE 'Inspection%'" + orderByStatement;
            }
            else if (department == "Grind")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Grind%'" + orderByStatement;
            }
            else if (department == "Polish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Polish%'" + orderByStatement;
            }
            else if (department == "All")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks" + orderByStatement;
            }

            return queryString;
        }

        private string SetWeeklyHoursQueryString(string weekStart, string weekEnd)
        {
            string department = "All";
            string queryString = null;
            string selectStatment = "JobNumber, ProjectNumber, TaskName, Duration, StartDate, FinishDate, Hours";
            string fromStatement = "Tasks";
            string whereStatement = "(StartDate BETWEEN #" + weekStart + "# AND #" + weekEnd + "#)";
            string orderByStatement = "ORDER BY StartDate ASC";
            //string groupByStatement = "GROUP BY ";

            if (department == "Design")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Design%' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Program Rough")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Rough' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Program Finish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Finish' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Program Electrodes")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Electrodes' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "CNC Rough")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Rough' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "CNC Finish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Finish' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "CNC Electrodes")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Electrodes' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "EDM Sinker")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'EDM Sinker' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Inspection")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE 'Inspection%' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Grind")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Grind%' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Polish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Polish%' AND " + whereStatement + " " + orderByStatement;
            }
            else if (department == "All")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE  " + whereStatement + " " + orderByStatement;
            }

            return queryString;
        }

        public List<Week> GetDayHours(string weekStart, string weekEnd)
        {
            List<Week> weeks = new List<Week>();

            string queryString = SetWeeklyHoursQueryString(weekStart, weekEnd);
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbCommand cmd = new OleDbCommand(queryString, Connection);

            string[] departmentArr = { "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "CNC Electrodes", "EDM Sinker", "EDM Wire (In-House)", "Polish (In-House)", "Inspection", "Grind" };

            foreach (string item in departmentArr)
            {
                weeks.Add(new Week(item));
            }

            Connection.Open();

            using (var rdr = cmd.ExecuteReader())
            {
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        if(rdr["TaskName"].ToString() == "Program Rough")
                        {
                            weeks[0].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }
                        else if (rdr["TaskName"].ToString() == "Program Finish")
                        {
                            weeks[1].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }
                        else if (rdr["TaskName"].ToString() == "Program Electrodes")
                        {
                            weeks[2].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }
                        else if (rdr["TaskName"].ToString() == "CNC Rough")
                        {
                            weeks[3].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }
                        else if (rdr["TaskName"].ToString() == "CNC Finish")
                        {
                            weeks[4].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }
                        else if (rdr["TaskName"].ToString() == "CNC Electrodes")
                        {
                            weeks[5].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }
                        else if (rdr["TaskName"].ToString() == "EDM Sinker")
                        {
                            weeks[6].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }
                        else if (rdr["TaskName"].ToString() == "EDM Wire (In-House)")
                        {
                            weeks[7].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }
                        else if (rdr["TaskName"].ToString() == "Polish (In-House)")
                        {
                            weeks[8].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }
                        else if (rdr["TaskName"].ToString().Contains("Inspection"))
                        {
                            weeks[9].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }
                        else if (rdr["TaskName"].ToString().Contains("Grind"))
                        {
                            weeks[10].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                        }

                        //Console.WriteLine($"{rdr["TaskName"]} {rdr["StartDate"]} {rdr["Hours"]}");
                    }
                }
                else
                {

                }

                
            }

            Connection.Close();
            Connection.Dispose();

            foreach (Week week in weeks)
            {
                Console.WriteLine("");
                Console.WriteLine(week.Department);

                foreach (ClassLibrary.Day day in week.DayList)
                {
                    Console.WriteLine($"{day.DayName} {(int)day.Hours}");
                }
            }

            return weeks;
        }

        public DataTable GetRoleCounts()
        {
            string queryString = "SELECT COUNT(*) AS RoleCount, Role FROM Roles GROUP BY Role";

            DataTable dt = new DataTable();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            return dt;
        }

        public void SetDailyDepartmentCapacities(string department)
        {
            DataTable dt = GetRoleCounts();
        }

        public DataTable GetDailyDepartmentCapacities()
        {
            string queryString = "SELECT * FROM Departments";
            DataTable dt = new DataTable();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            return dt;
        }

        // Creates a weeklist with 20 weeks for each department.
        public List<Week> InitializeDeptWeeksList(DateTime wsDate)
        {
            List<Week> weekList = new List<Week>();
            string[] departmentArr = { "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "CNC Electrodes", "EDM Sinker", "EDM Wire (In-House)", "Polish (In-House)", "Inspection", "Grind" };

            for (int i = 1; i <= 20; i++)
            {
                //wsDate = wsDate.AddDays((i - 1) * 7);
                //weDate = wsDate.AddDays(6);

                foreach (string department in departmentArr)
                {
                    weekList.Add(new Week(i, wsDate.AddDays((i - 1) * 7), wsDate.AddDays((i - 1) * 7 + 6), department));
                }
            }

            return weekList;
        }

        public List<Week> GetWeekHours(string weekStart, string weekEnd)
        {
            List<Week> weekList = new List<Week>();
            List<Week> deptWeekList = new List<Week>();
            Week weekTemp;
            DateTime wsDate = Convert.ToDateTime(weekStart);
            int weekNum, count = 0;

            string queryString = SetWeeklyHoursQueryString(weekStart, weekEnd);
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbCommand cmd = new OleDbCommand(queryString, Connection);

            weekList = InitializeDeptWeeksList(wsDate);

            //Console.WriteLine("\nLoad");

            Connection.Open();

            using (var rdr = cmd.ExecuteReader())
            {
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        //Console.WriteLine($"{++count} {rdr["JobNumber"].ToString()}-{rdr["ProjectNumber"].ToString()} {rdr["TaskName"].ToString()} {rdr["Duration"].ToString()} {rdr["Hours"].ToString()}");

                        var week = from wk in weekList
                                   where (rdr["TaskName"].ToString().StartsWith(wk.Department) || (rdr["TaskName"].ToString().Contains("Grind") && rdr["TaskName"].ToString().Contains(wk.Department))) // && Convert.ToDateTime(rdr["StartDate"]) >= wk.WeekStart && Convert.ToDateTime(rdr["StartDate"]) <= wk.WeekEnd
                                   orderby wk.WeekNum ascending
                                   select wk;

                        if (week.Any())
                        {
                            deptWeekList = week.ToList();
                            weekTemp = deptWeekList.Find(x => x.WeekStart <= Convert.ToDateTime(rdr["StartDate"]) && x.WeekEnd >= Convert.ToDateTime(rdr["StartDate"]));
                            weekNum = weekTemp.WeekNum;
                            //weekTemp.AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));

                            //Console.WriteLine(rdr["Duration"].ToString());

                            //Console.WriteLine($"{rdr["JobNumber"].ToString()}-{rdr["ProjectNumber"].ToString()} {rdr["TaskName"].ToString()} {rdr["Duration"].ToString()} {Convert.ToDateTime(rdr["StartDate"]).ToShortDateString()} {Convert.ToDateTime(rdr["FinishDate"]).ToShortDateString()} {rdr["Hours"].ToString()}");

                            double hours = Convert.ToInt32(rdr["Hours"]);
                            double days = (int)GetBusinessDays(Convert.ToDateTime(rdr["StartDate"]), Convert.ToDateTime(rdr["FinishDate"]));
                            DateTime date = Convert.ToDateTime(rdr["StartDate"]);
                            decimal dailyAVG;

                            if (days == 0)
                            {
                                dailyAVG = (decimal)hours;
                            }
                            else
                            {
                                dailyAVG = (decimal)(hours / days);
                            }

                            if (days >= 1)
                            {
                                while (days > 0)
                                {

                                    if (date.DayOfWeek == DayOfWeek.Saturday)
                                    {
                                        date = date.AddDays(1);

                                        weekNum++;

                                        if(weekNum > 20)
                                        {
                                            goto MyEnd;
                                        }

                                        //weekTemp = deptWeekList.Find(x => x.WeekNum == weekNum);
                                        weekTemp = deptWeekList[weekNum - 1];
                                        //weekTemp.AddHoursToDay((int)date.DayOfWeek, dailyAVG);
                                        //Console.WriteLine($"{weekTemp.Department} {weekTemp.WeekStart.ToShortDateString()} {date.DayOfWeek} {dailyAVG} {days}");
                                    }
                                    else
                                    {
                                        weekTemp.AddHoursToDay((int)date.DayOfWeek, dailyAVG);
                                        if(weekTemp.Department == "CNC Rough")
                                        Console.WriteLine($"{weekTemp.Department} {weekTemp.WeekStart.ToShortDateString()} {date.DayOfWeek} {dailyAVG} {days}");
                                        days -= 1;
                                    }

                                    
                                    date = date.AddDays(1);
                                }
                            }
                            else
                            {
                                weekTemp.AddHoursToDay((int)date.AddDays(days).DayOfWeek, dailyAVG);
                                if (weekTemp.Department == "CNC Rough")
                                    Console.WriteLine($"{weekTemp.Department} {weekTemp.WeekStart.ToShortDateString()} {date.AddDays(days).DayOfWeek} {dailyAVG} {days}");
                            }
                        }
                    }
                }
                else
                {

                }
            }

        MyEnd:;

            Connection.Close();
            Connection.Dispose();

            //Console.WriteLine("\nReview:");

            //foreach (Week week in weekList)
            //{
            //    Console.WriteLine($"{week.Department} {week.GetWeekHours()} {week.WeekStart.ToShortDateString()} - {week.WeekEnd.ToShortDateString()}");
            //}

            return weekList;
        }

        public DataTable GetResourceData()
        {
            DataTable dt = new DataTable();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            string queryString = "SELECT ResourceName, ID From Resources ORDER BY ResourceName ASC";
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            foreach (DataRow nrow in dt.Rows)
            {

            }

            return dt;
        }

        public ProjectModel GetProject(string jobNumber, int projectNumber)
        {
            ProjectModel project = GetProjectInfo(jobNumber, projectNumber);

            AddComponents(project);

            AddTasks(project);

            return project;
        }

        public ProjectModel GetProjectInfo(string jobNumber, int projectNumber)
        {
            OleDbCommand cmd;
            ProjectModel pi = null;
            string queryString;

            queryString = "SELECT * FROM Projects WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber";
            OleDbConnection Connection = new OleDbConnection(ConnString);
            cmd = new OleDbCommand(queryString, Connection);
            cmd.Parameters.AddWithValue("@jobNumber", jobNumber);
            cmd.Parameters.AddWithValue("@projectNumber", projectNumber);

            try
            {
                Connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            pi = new ProjectModel
                            (
                                        jobNumber: Convert.ToString(rdr["JobNumber"]),
                                    projectNumber: Convert.ToInt32(rdr["ProjectNumber"]),
                                          dueDate: Convert.ToDateTime(rdr["DueDate"]),
                                           status: Convert.ToString(rdr["Status"]),
                                        toolMaker: Convert.ToString(rdr["ToolMaker"]),
                                         designer: Convert.ToString(rdr["Designer"]),
                                  roughProgrammer: Convert.ToString(rdr["RoughProgrammer"]),
                               electrodProgrammer: Convert.ToString(rdr["ElectrodeProgrammer"]),
                                 finishProgrammer: Convert.ToString(rdr["FinishProgrammer"]),
                               kanBanWorkbookPath: Convert.ToString(rdr["KanBanWorkbookPath"])
                            );
                        }
                    }
                }

                Connection.Close();
                Connection.Dispose();
            }
            catch (Exception e)
            {
                Connection.Close();
                MessageBox.Show(e.Message, "GetProjectInfo");
            }

            return pi;
        }

        public List<ProjectModel> GetProjectInfoList()
        {
            string queryString = "SELECT * FROM Projects";
            OleDbCommand cmd;
            ProjectModel pi;
            List<ProjectModel> piList = new List<ProjectModel>();

            cmd = new OleDbCommand(queryString, Connection);

            try
            {
                Connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            pi = new ProjectModel
                            (
                                        jobNumber: Convert.ToString(rdr["JobNumber"]),
                                    projectNumber: Convert.ToInt32(rdr["ProjectNumber"]),
                                          dueDate: Convert.ToDateTime(rdr["DueDate"]),
                                           status: Convert.ToString(rdr["Status"]),
                                        toolMaker: Convert.ToString(rdr["ToolMaker"]),
                                         designer: Convert.ToString(rdr["Designer"]),
                                  roughProgrammer: Convert.ToString(rdr["RoughProgrammer"]),
                               electrodProgrammer: Convert.ToString(rdr["ElectrodeProgrammer"]),
                                 finishProgrammer: Convert.ToString(rdr["FinishProgrammer"]),
                               kanBanWorkbookPath: Convert.ToString(rdr["KanBanWorkbookPath"])
                            );

                            piList.Add(pi);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Connection.Close();
            }

            return piList;
        }

        public void AddComponents(ProjectModel project)
        {
            OleDbCommand cmd;
            ComponentModel component;

            string queryString;

            queryString = "SELECT * FROM Components WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber ORDER BY Component";

            cmd = new OleDbCommand(queryString, Connection);
            cmd.Parameters.AddWithValue("@jobNumber", project.JobNumber);
            cmd.Parameters.AddWithValue("@projectNumber", project.ProjectNumber);

            try
            {
                Connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            component = new ComponentModel
                            (
                                           name: rdr["Component"],
                                          notes: rdr["Notes"],
                                       priority: rdr["Priority"],
                                       position: rdr["Position"],
                                       quantity: rdr["Quantity"],
                                         spares: rdr["Spares"],
                                        picture: rdr["Pictures"],
                                       material: rdr["Material"],
                                         finish: rdr["Finish"],
                                    taskIDCount: rdr["TaskIDCount"]
                            );

                            project.AddComponent(component);
                        }
                    }
                    else
                    {
                        project.AddComponentList(GetComponentListFromTasksTable(project.JobNumber, project.ProjectNumber));
                    }
                }
            }
            finally
            {
                Connection.Close();
            }
        }

        public List<ComponentModel> GetComponentListFromTasksTable(string jobNumber, int projectNumber)
        {
            OleDbCommand cmd;
            List<ComponentModel> componentList = new List<ComponentModel>();

            string queryString;

            queryString = "SELECT DISTINCT Component FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber";
            OleDbConnection Connection = new OleDbConnection(ConnString);
            cmd = new OleDbCommand(queryString, Connection);
            cmd.Parameters.AddWithValue("@jobNumber", jobNumber);
            cmd.Parameters.AddWithValue("@projectNumber", projectNumber);

            try
            {
                Connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            componentList.Add(new ComponentModel
                            (
                                    name: rdr["Component"]
                            ));
                        }
                    }
                    else
                    {

                    }
                }

                Connection.Close();
            }
            catch (Exception e)
            {
                Connection.Close();
                MessageBox.Show(e.Message, "getComponentListFromTaskTable");
            }

            return componentList;
        }

        private void AddTasks(ProjectModel project)
        {
            List<TaskModel> projectTaskList = GetProjectTaskList(project.JobNumber, project.ProjectNumber);

            foreach (ComponentModel component in project.Components)
            {
                var tasks = from t in projectTaskList
                            where t.Component == component.Component
                            orderby t.ID ascending
                            select t;

                component.AddTaskList(tasks.ToList());
            }

            foreach (ComponentModel component in project.Components)
            {
                Console.WriteLine(component.Component);

                foreach (TaskModel task in component.Tasks)
                {
                    Console.WriteLine($"   {task.ID} {task.TaskName}");
                    task.HasInfo = true;
                }
            }
        }

        public List<TaskModel> GetProjectTaskList(string jobNumber, int projectNumber)
        {
            OleDbCommand cmd;
            List<TaskModel> taskList = new List<TaskModel>();

            string queryString;
            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber";

            cmd = new OleDbCommand(queryString, Connection);
            cmd.Parameters.AddWithValue("@jobNumber", jobNumber);
            cmd.Parameters.AddWithValue("@projectNumber", projectNumber);

            try
            {
                Connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            taskList.Add(new TaskModel
                            (
                                    taskName: rdr["TaskName"],
                                          id: rdr["TaskID"],
                                   component: rdr["Component"],
                                       hours: rdr["Hours"],
                                    duration: rdr["Duration"],
                                   startDate: rdr["StartDate"],
                                  finishDate: rdr["FinishDate"],
                                      status: rdr["Status"],
                               dateCompleted: rdr["DateCompleted"],
                                    initials: rdr["Initials"],
                                     machine: rdr["Machine"],
                                   personnel: rdr["Resource"],
                                predecessors: rdr["Predecessors"],
                                       notes: rdr["Notes"]
                            ));
                        }
                    }
                }
            }
            finally
            {
                Connection.Close();
            }

            return taskList;
        }

        public DataTable GetComponentCompletionPercents()
        {
            //  TOP(Components.TaskIDCount) AS TaskIDCount, , Components.Pictures
            string queryString = "SELECT COUNT(*) As CompletedTaskCount, MAX(Tasks.TaskID) AS LastCompletedTask, Components.TaskIDCount, Projects.JobNumber, Projects.ProjectNumber, Projects.ToolMaker, Components.Component FROM (Tasks INNER JOIN Components ON Components.Component = Tasks.Component) INNER JOIN Projects ON Projects.ProjectNumber = Components.ProjectNumber WHERE Tasks.Status = 'Completed' GROUP BY Projects.JobNumber, Projects.ProjectNumber, Projects.ToolMaker, Components.Component, Components.TaskIDCount";
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbCommand cmd = new OleDbCommand(queryString, Connection);
            List<ProjectModel> projects = new List<ProjectModel>();
            string project = "", jobNumber = "";
            int lastTaskID = 0, taskIDCount;
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            ImageConverter ic = new ImageConverter();

            dt.Columns.Add("JobNumber", typeof(string));
            dt.Columns.Add("ProjectNumber", typeof(int));
            dt.Columns.Add("ToolMaker", typeof(string));
            dt.Columns.Add("Component", typeof(string));
            //dt.Columns.Add("Pictures", typeof(Image));
            dt.Columns.Add("LastCompletedTask");
            dt.Columns.Add("TaskIDCount", typeof(int));
            dt.Columns.Add("Status", typeof(string));
            dt.Columns.Add("PercentComplete", typeof(double));

            dt2 = GetAllTasks();

            try
            {
                Connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            DataRow row = dt.NewRow();

                            row["JobNumber"] = rdr["JobNumber"].ToString();
                            row["ProjectNumber"] = Convert.ToInt32(rdr["ProjectNumber"]);
                            row["ToolMaker"] = rdr["ToolMaker"].ToString();
                            row["Component"] = rdr["Component"].ToString();
                            //row["Pictures"] = (Image)ic.ConvertFrom(rdr["Pictures"]);
                            row["PercentComplete"] = Convert.ToDouble(rdr["CompletedTaskCount"]) / Convert.ToDouble(rdr["TaskIDCount"]);

                            lastTaskID = Convert.ToInt16(rdr["LastCompletedTask"]);
                            taskIDCount = Convert.ToInt16(rdr["TaskIDCount"]);

                            row["LastCompletedTask"] = rdr["LastCompletedTask"];
                            row["TaskIDCount"] = taskIDCount;

                            if (lastTaskID < taskIDCount)
                            {
                                row["Status"] = FindTask(dt2, rdr["JobNumber"].ToString(), Convert.ToInt32(rdr["ProjectNumber"]), rdr["Component"].ToString(), lastTaskID + 1);
                            }
                            else
                            {
                                row["Status"] = "Done";
                            }

                            dt.Rows.Add(row);
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                Connection.Close();
            }

            foreach (DataRow nrow in dt.Rows)
            {
                Console.WriteLine($"{nrow["JobNumber"]} {nrow["ProjectNumber"]} {nrow["Component"]} {nrow["Status"]} {nrow["PercentComplete"]}");
            }

            return dt;
        }

        public void ForwardDateProjectTasks(string jobNumber, int projectNumber, List<string> componentList, DateTime forwardDate)
        {
            OleDbDataAdapter adapter;
            DataTable dt = new DataTable();
            string queryString;
            bool skipDatedTasks = false;

            if (forwardDate == new DateTime(2000, 1, 1))
            {
                return;
            }

            if (componentList.Count == 0)
            {
                XtraMessageBox.Show("No Components were selected.");
                return;
            }

            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber ORDER BY ID DESC";

            adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
            adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = projectNumber;

            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

            adapter.Fill(dt);

            if (componentList == null)
            {
                foreach (DataRow nrow in dt.Rows)
                {
                    nrow["StartDate"] = DBNull.Value;
                    nrow["FinishDate"] = DBNull.Value;
                }
            }
            else
            {
                foreach (string component in componentList)
                {
                    var result2 = from DataRow myRow in dt.Rows
                                  where myRow["Component"].ToString() == component
                                  select myRow;

                    foreach (DataRow nrow in result2)
                    {
                        nrow["StartDate"] = DBNull.Value;
                        nrow["FinishDate"] = DBNull.Value;
                    }
                }
            }

            var result = from DataRow myRow in dt.Rows
                         where myRow["Predecessors"].ToString() == ""
                         select myRow;

            foreach (DataRow nrow in result)
            {
                //if(skipDatedTasks = true && (nrow["StartDate"] != DBNull.Value || nrow["FinishDate"] != DBNull.Value))
                //{
                //    goto Skip;
                //}

                if (componentList == null)
                {
                    nrow["StartDate"] = forwardDate;
                    nrow["FinishDate"] = AddBusinessDays(forwardDate, nrow["Duration"].ToString());

                    //Skip:;

                    ForwardDateTask(Convert.ToInt32(nrow["TaskID"]), nrow["Component"].ToString(), skipDatedTasks, Convert.ToDateTime(nrow["FinishDate"]), dt);
                }
                else if (componentList.Exists(x => x == nrow["Component"].ToString()))
                {
                    nrow["StartDate"] = forwardDate;
                    nrow["FinishDate"] = AddBusinessDays(forwardDate, nrow["Duration"].ToString());

                    //Skip:;

                    ForwardDateTask(Convert.ToInt32(nrow["TaskID"]), nrow["Component"].ToString(), skipDatedTasks, Convert.ToDateTime(nrow["FinishDate"]), dt);
                }
            }

            //foreach (DataRow nrow in dt.Rows)
            //{
            //    Console.WriteLine($"{nrow["TaskID"].ToString()} {nrow["Component"].ToString()} {nrow["StartDate"].ToString()} {nrow["FinishDate"].ToString()}");
            //}

            adapter.UpdateCommand = builder.GetUpdateCommand();
            adapter.Update(dt);

        }

        private void ForwardDateTask(int predecessorID, string component, bool skipDatedTasks, DateTime predecessorFinishDate, DataTable projectTaskTable)
        {
            var result = from DataRow myRow in projectTaskTable.Rows
                         where myRow["Predecessors"].ToString().Contains(predecessorID.ToString()) && myRow["Component"].ToString() == component
                         select myRow;

            //Console.WriteLine(predecessorTaskID);

            foreach (DataRow nrow in result)
            {
                if (nrow["StartDate"] == DBNull.Value || Convert.ToDateTime(nrow["StartDate"]) < predecessorFinishDate)
                {
                    if (skipDatedTasks == true && (nrow["StartDate"] != DBNull.Value || nrow["FinishDate"] != DBNull.Value))
                    {
                        goto Skip;
                    }

                    nrow["StartDate"] = predecessorFinishDate;
                    //MessageBox.Show(nrow["TaskName"].ToString());
                    nrow["FinishDate"] = AddBusinessDays(Convert.ToDateTime(nrow["StartDate"]), nrow["Duration"].ToString());

                    Skip:;

                    ForwardDateTask(Convert.ToInt16(nrow["TaskID"]), nrow["Component"].ToString(), skipDatedTasks, Convert.ToDateTime(nrow["FinishDate"]), projectTaskTable);
                }
            }

            //foreach(DataRow nrow in projectTaskTable.Rows)
            //{
            //    Console.WriteLine(nrow["TaskID"].ToString() + " " + nrow["StartDate"].ToString() + " " + nrow["FinishDate"].ToString());
            //}
        }

        private void ClearProjectTaskDates(string jobNumber, int projectNumber)
        {
            OleDbDataAdapter adapter;
            DataTable dt = new DataTable();
            string queryString;

            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber ORDER BY ID DESC";

            adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
            adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = projectNumber;

            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

            adapter.Fill(dt);

            foreach (DataRow nrow in dt.Rows)
            {
                nrow["StartDate"] = DBNull.Value;
                nrow["FinishDate"] = DBNull.Value;
            }

            adapter.UpdateCommand = builder.GetUpdateCommand();
            adapter.Update(dt);
        }

        private void ClearComponentTaskDates(string jobNumber, int projectNumber, string component)
        {
            OleDbDataAdapter adapter;
            DataTable dt = new DataTable();
            string queryString;

            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component ORDER BY ID DESC";

            adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
            adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = projectNumber;
            adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = component;

            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

            adapter.Fill(dt);

            foreach (DataRow nrow in dt.Rows)
            {
                nrow["StartDate"] = DBNull.Value;
                nrow["FinishDate"] = DBNull.Value;
            }

            adapter.UpdateCommand = builder.GetUpdateCommand();
            adapter.Update(dt);
        }

        public void ForwardDateComponent(string jobNumber, int projectNumber, string component, DateTime currentTaskFinishDate, int currentTaskID)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;

            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component ORDER BY TaskID ASC";

            adapter.SelectCommand = new OleDbCommand(queryString, Connection);

            adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 20).Value = jobNumber;
            adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = projectNumber;
            adapter.SelectCommand.Parameters.Add("@component", OleDbType.VarChar, 30).Value = component;

            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

            adapter.Fill(datatable);

            UpdateStartAndFinishDates(currentTaskID, datatable, currentTaskFinishDate);

            adapter.UpdateCommand = builder.GetUpdateCommand();
            adapter.Update(datatable);
        }

        // This needs to be a separate method so that recursion can take place.
        private void UpdateStartAndFinishDates(int id, DataTable dt, DateTime fd)
        {
            string[] predecessorArr;

            Console.WriteLine("Update Start and Finish Dates");

            foreach (DataRow nrow in dt.Rows)
            {
                predecessorArr = nrow["Predecessors"].ToString().Split(',');

                //Console.WriteLine(id + " " + nrow["TaskID"] + " " + nrow["Component"] + " " + nrow["Predecessors"]);

                for (int i = 0; i < predecessorArr.Length; i++)
                {
                    if (predecessorArr[i].ToString() == id.ToString())
                    {
                        //Console.WriteLine(currentTaskID + " " + nrow2["TaskID"] + " " + nrow2["Component"] + " " + predecessorArr[i2].ToString() + " " + nrow2["Predecessors"]);
                        if (nrow["StartDate"] == DBNull.Value)
                        {
                            Console.WriteLine(id + " " + nrow["TaskID"] + " " + nrow["Component"] + " " + nrow["StartDate"] + " " + nrow["FinishDate"] + " " + nrow["Predecessors"]);
                        }
                        else if (Convert.ToDateTime(nrow["StartDate"]) < fd) // If start date of current task comes before finish date of predecessor.
                        {
                            nrow["StartDate"] = fd;
                            nrow["FinishDate"] = AddBusinessDays(Convert.ToDateTime(nrow["StartDate"]), nrow["Duration"].ToString());
                            Console.WriteLine(id + " " + nrow["TaskID"] + " " + nrow["Component"] + " " + Convert.ToDateTime(nrow["StartDate"]).ToShortDateString() + " " + Convert.ToDateTime(nrow["FinishDate"]).ToShortDateString() + " " + nrow["Predecessors"]);
                            //Console.WriteLine(currentTaskID + " " + currentTaskFinishDate + " " + nrow2["TaskID"] + " " + predecessorArr[i2].ToString() + " " + nrow2["Predecessors"]);
                        }
                        else if (Convert.ToDateTime(nrow["StartDate"]) > fd) // If start date of current task comes after the finish date of predecessor.
                        {
                            // Do nothing.  Otherwise you won't have any of the separation that may be necessary when scheduling tasks.
                            //nrow["StartDate"] = fd;
                            //nrow["FinishDate"] = AddBusinessDays(Convert.ToDateTime(nrow["StartDate"]), nrow["Duration"].ToString());
                            Console.WriteLine(id + " " + nrow["TaskID"] + " " + nrow["Component"] + " " + Convert.ToDateTime(nrow["StartDate"]).ToShortDateString() + " " + Convert.ToDateTime(nrow["FinishDate"]).ToShortDateString() + " " + nrow["Predecessors"]);
                        }

                        if (nrow["FinishDate"] != DBNull.Value)
                            UpdateStartAndFinishDates(Convert.ToInt16(nrow["TaskID"]), dt, Convert.ToDateTime(nrow["FinishDate"]));

                        goto NextStep;
                    }
                }

                predecessorArr = null;

                NextStep:;

                //Console.WriteLine(nrow["Component"] + " " + nrow["Predecessors"]);
            }
        }

        // This needs to be a separate method so that recursion can take place.
        private void UpdateStartAndFinishDates2(int id, DataTable dt, DateTime fd)
        {
            string[] predecessorArr;

            foreach (DataRow nrow in dt.Rows)
            {
                predecessorArr = nrow["Predecessors"].ToString().Split(',');

                //Console.WriteLine(id + " " + nrow["TaskID"] + " " + nrow["Component"] + " " + nrow["Predecessors"]);

                for (int i = 0; i < predecessorArr.Length; i++)
                {
                    if (predecessorArr[i].ToString() == id.ToString())
                    {
                        //Console.WriteLine(currentTaskID + " " + nrow2["TaskID"] + " " + nrow2["Component"] + " " + predecessorArr[i2].ToString() + " " + nrow2["Predecessors"]);
                        if (nrow["StartDate"] == DBNull.Value)
                        {
                            Console.WriteLine(id + " " + nrow["TaskID"] + " " + nrow["Component"] + " " + nrow["StartDate"] + " " + nrow["FinishDate"] + " " + nrow["Predecessors"]);
                        }
                        else if (Convert.ToDateTime(nrow["StartDate"]) < fd) // If start date of current task comes before finish date of predecessor.
                        {
                            nrow["StartDate"] = fd;
                            nrow["FinishDate"] = AddBusinessDays(Convert.ToDateTime(nrow["StartDate"]), nrow["Duration"].ToString());
                            Console.WriteLine(id + " " + nrow["TaskID"] + " " + nrow["Component"] + " " + Convert.ToDateTime(nrow["StartDate"]).ToShortDateString() + " " + Convert.ToDateTime(nrow["FinishDate"]).ToShortDateString() + " " + nrow["Predecessors"]);
                            //Console.WriteLine(currentTaskID + " " + currentTaskFinishDate + " " + nrow2["TaskID"] + " " + predecessorArr[i2].ToString() + " " + nrow2["Predecessors"]);
                        }
                        else if (Convert.ToDateTime(nrow["StartDate"]) > fd) // If start date of current task comes after the finish date of predecessor.
                        {
                            // Do nothing.  Otherwise you won't have any of the separation that may be necessary when scheduling tasks.
                            //nrow["StartDate"] = fd;
                            //nrow["FinishDate"] = AddBusinessDays(Convert.ToDateTime(nrow["StartDate"]), nrow["Duration"].ToString());
                            Console.WriteLine(id + " " + nrow["TaskID"] + " " + nrow["Component"] + " " + Convert.ToDateTime(nrow["StartDate"]).ToShortDateString() + " " + Convert.ToDateTime(nrow["FinishDate"]).ToShortDateString() + " " + nrow["Predecessors"]);
                        }

                        if (nrow["FinishDate"] != DBNull.Value)
                            UpdateStartAndFinishDates(Convert.ToInt16(nrow["TaskID"]), dt, Convert.ToDateTime(nrow["FinishDate"]));

                        goto NextStep;
                    }
                }

                predecessorArr = null;

                NextStep:;

                //Console.WriteLine(nrow["Component"] + " " + nrow["Predecessors"]);
            }
        }

        public void BackDateProjectTasks(string jobNumber, int projectNumber, List<string> componentList, DateTime backDate)
        {
            OleDbDataAdapter adapter;
            DataTable dt = new DataTable();
            string queryString;
            bool skipDatedTasks = false;

            if (backDate == new DateTime(2000, 1, 1))
            {
                return;
            }

            if (componentList.Count == 0)
            {
                XtraMessageBox.Show("No components selected.");
                return;
            }

            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber ORDER BY TaskID DESC";

            adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
            adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = projectNumber;

            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

            adapter.Fill(dt);

            var result = from myRow in dt.AsEnumerable()
                         group myRow by new { component = myRow.Field<string>("Component") } into components
                         select new { highestID = components.Max(c => c.Field<int>("TaskID")), component = components.Key.component };

            foreach (var lastTask in result)
            {
                if (componentList == null)
                {
                    Console.WriteLine($"{lastTask.highestID} {lastTask.component}");
                    BackDateTask(lastTask.highestID, lastTask.component, skipDatedTasks, backDate, dt);
                }
                else if (componentList.Exists(x => x == lastTask.component))
                {
                    Console.WriteLine($"{lastTask.highestID} {lastTask.component}");
                    BackDateTask(lastTask.highestID, lastTask.component, skipDatedTasks, backDate, dt);
                }

            }

            //foreach (DataRow nrow in dt.Rows)
            //{
            //    Console.WriteLine($"{nrow["TaskID"].ToString()} {nrow["Component"].ToString()} {nrow["StartDate"].ToString()} {nrow["FinishDate"].ToString()}");
            //}

            adapter.UpdateCommand = builder.GetUpdateCommand();

            adapter.Update(dt);
        }

        private void BackDateTask(int taskID, string component, bool skipDatedTasks, DateTime descendantStartDate, DataTable projectTaskTable)
        {
            string[] predecessors;

            var result = from DataRow myRow in projectTaskTable.Rows
                         where Convert.ToInt32(myRow["TaskID"]) == taskID && myRow["Component"].ToString() == component
                         select myRow;

            //Console.WriteLine(predecessorTaskID);

            foreach (DataRow nrow in result)
            {
                if (skipDatedTasks == true && (nrow["FinishDate"] != DBNull.Value || nrow["StartDate"] != DBNull.Value))
                {
                    goto Skip;
                }

                nrow["FinishDate"] = descendantStartDate;
                //MessageBox.Show(nrow["TaskName"].ToString());
                nrow["StartDate"] = SubtractBusinessDays(Convert.ToDateTime(nrow["FinishDate"]), nrow["Duration"].ToString());

                Skip:;

                // If a task has more than one predecessor.
                // Backdate each predecessor.
                if (nrow["Predecessors"].ToString().Contains(','))
                {
                    predecessors = nrow["Predecessors"].ToString().Split(',');

                    foreach (string id in predecessors)
                    {
                        BackDateTask(Convert.ToInt32(id), component, skipDatedTasks, Convert.ToDateTime(nrow["StartDate"]), projectTaskTable);
                    }
                }
                // If a task has one predecessor.
                // Backdate the one predecessor.
                else if (nrow["Predecessors"].ToString() != "")
                {
                    BackDateTask(Convert.ToInt32(nrow["Predecessors"]), component, skipDatedTasks, Convert.ToDateTime(nrow["StartDate"]), projectTaskTable);
                }
                // If a task has no predecessors.
                // Exit method.
                else if (nrow["Predecessors"].ToString() == "")
                {
                    return;
                }
            }

            //foreach(DataRow nrow in projectTaskTable.Rows)
            //{
            //    Console.WriteLine(nrow["TaskID"].ToString() + " " + nrow["StartDate"].ToString() + " " + nrow["FinishDate"].ToString());
            //}
        }

        public DateTime SubtractBusinessDays(DateTime finishDate, string durationSt)
        {
            int days;
            string[] duration = durationSt.Split(' ');
            days = Convert.ToInt16(duration[0]);

            if (days < 0)
            {
                throw new ArgumentException("Days cannot be negative.", "days");
            }

            if (days == 0) return finishDate;

            if (finishDate.DayOfWeek == DayOfWeek.Saturday)
            {
                finishDate = finishDate.AddDays(-1);
                days -= 1;
            }
            else if (finishDate.DayOfWeek == DayOfWeek.Sunday)
            {
                finishDate = finishDate.AddDays(-2);
                days -= 1;
            }

            finishDate = finishDate.AddDays(-days / 5 * 7);

            int extraDays = days % 5;

            if ((int)finishDate.DayOfWeek - extraDays < 1)
            {
                extraDays += 2;
            }

            return finishDate.AddDays(-extraDays);
        }

        private DataTable GetAllTasks()
        {
            DataTable dt = new DataTable();
            string queryString = "SELECT * FROM Tasks";

            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            return dt;
        }

        private string FindTask(DataTable dataTable, string jobNumber, int projectNumber, string component, int taskID)
        {
            DataRow task = dataTable.Rows.Cast<DataRow>().FirstOrDefault(x => (string)x["JobNumber"] == jobNumber && (int)x["ProjectNumber"] == projectNumber && (int)x["TaskID"] == taskID);

            return task["TaskName"].ToString();
        }

        public bool LoadProjectToDB(ProjectModel project)
        {

            //if(result == DialogResult.Yes)
            //{
            //    //int baseIDNumber = getHighestProjectTaskID(project.JobNumber, project.ProjectNumber);
            //    //updateProjectData(pi);
            //    //foreach (Component component in project.ComponentList)
            //    //{
            //    //    foreach (TaskInfo task in component.Tasks)
            //    //    {
            //    //        task.ChangeIDs(baseIDNumber);
            //    //    }
            //    //}

            //    addTaskDataTableToDatabase(createDataTableFromTaskList(project));
            //}
            //else if (result == DialogResult.No)
            //{
            //    return;
            //}

            if (ProjectExists(project.ProjectNumber))
            {
                MessageBox.Show("There is another project with that same project number. Enter a different project number");
                return false;
            }
            else
            {
                if (AddProjectDataToDatabase(project) &&
                AddComponentDataTableToDatabase(CreateDataTableFromComponentList(project)) &&
                AddTaskDataTableToDatabase(CreateDataTableFromTaskList(project)))
                {
                    MessageBox.Show("Project Loaded!");
                    return true;
                }
                else
                {
                    MessageBox.Show("Project load failed.");
                    return false;
                }
            }
        }

        private bool AddProjectDataToDatabase(ProjectModel project)
        {
            var adapter = new OleDbDataAdapter();
            string queryString;

            // Keep query in queryString to make query more visible.
            queryString = "INSERT INTO Projects (JobNumber, ProjectNumber, DueDate, Designer, ToolMaker, RoughProgrammer, ElectrodeProgrammer, FinishProgrammer) " +
                            "VALUES (@jobNumber, @projectNumber, @DueDate, @Designer, @ToolMaker, @RoughProgrammer, @electrodeProgrammer, @finishProgrammer)";

            adapter.InsertCommand = new OleDbCommand(queryString, Connection);

            adapter.InsertCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 20).Value = project.JobNumber;
            adapter.InsertCommand.Parameters.AddWithValue("@projectNumber", project.ProjectNumber);
            adapter.InsertCommand.Parameters.AddWithValue("@dueDate", project.DueDate);
            adapter.InsertCommand.Parameters.AddWithValue("@designer", project.Designer);
            adapter.InsertCommand.Parameters.AddWithValue("@toolMaker", project.ToolMaker);
            adapter.InsertCommand.Parameters.AddWithValue("@roughProgrammer", project.RoughProgrammer);
            adapter.InsertCommand.Parameters.AddWithValue("@electrodeProgrammer", project.ElectrodeProgrammer);
            adapter.InsertCommand.Parameters.AddWithValue("@finishProgrammer", project.FinishProgrammer);

            try
            {
                Connection.Open();
                adapter.InsertCommand.ExecuteNonQuery();
                Connection.Close();
                Console.WriteLine("Project loaded.");
                //MessageBox.Show("Project Loaded!");
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message, "OledbException Error");
                return false;
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "Exception Error");
                return false;
            }

            return true;
        }

        private bool AddComponentDataTableToDatabase(DataTable dt)
        {
            var adapter = new OleDbDataAdapter();
            adapter.SelectCommand = new OleDbCommand("SELECT * FROM Components", Connection);

            var cbr = new OleDbCommandBuilder(adapter);

            cbr.QuotePrefix = "[";
            cbr.QuoteSuffix = "]";
            cbr.GetDeleteCommand();
            cbr.GetInsertCommand();
            adapter.UpdateCommand = cbr.GetUpdateCommand();
            //Console.WriteLine(cbr.GetInsertCommand().CommandText);

            try
            {
                Connection.Open();
                adapter.Update(dt);
                Connection.Close();
                Console.WriteLine("Components Loaded.");
            }
            catch (OleDbException ex)
            {
                Connection.Close();
                MessageBox.Show(ex.Message, "OledbException Error");
                return false;
            }
            catch (Exception x)
            {
                Connection.Close();
                MessageBox.Show(x.Message, "Exception Error");
                return false;
            }

            return true;
        }

        private bool AddTaskDataTableToDatabase(DataTable dt)
        {
            var adapter = new OleDbDataAdapter();
            adapter.SelectCommand = new OleDbCommand("SELECT * FROM Tasks", Connection);

            var cbr = new OleDbCommandBuilder(adapter);

            cbr.GetDeleteCommand();
            cbr.GetInsertCommand();
            adapter.UpdateCommand = cbr.GetUpdateCommand();
            //Console.WriteLine(cbr.GetInsertCommand().CommandText);

            try
            {
                Connection.Open();
                adapter.Update(dt);
                Connection.Close();
                Console.WriteLine("Tasks Loaded.");
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message, "OledbException Error");
                return false;
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "Exception Error");
                return false;
            }

            return true;
        }


        private DataTable CreateDataTableFromComponentList(ProjectModel project)
        {
            DataTable dt = new DataTable();
            int position = 0;

            dt.Columns.Add("JobNumber", typeof(string));
            dt.Columns.Add("ProjectNumber", typeof(int));
            dt.Columns.Add("Component", typeof(string));
            dt.Columns.Add("Notes", typeof(string));
            dt.Columns.Add("Position", typeof(int));
            dt.Columns.Add("Priority", typeof(int));
            dt.Columns.Add("Pictures", typeof(byte[]));
            dt.Columns.Add("Material", typeof(string));
            dt.Columns.Add("Finish", typeof(string));
            dt.Columns.Add("TaskIDCount", typeof(int));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("Spares", typeof(int));
            dt.Columns.Add("Status", typeof(string));
            dt.Columns.Add("PercentComplete", typeof(int));

            foreach (ComponentModel component in project.Components)
            {
                DataRow row = dt.NewRow();

                row["JobNumber"] = project.JobNumber;
                row["ProjectNumber"] = project.ProjectNumber;
                row["Component"] = component.Component;
                row["Quantity"] = component.Quantity;
                row["Spares"] = component.Spares;
                row["Material"] = component.Material;
                row["Finish"] = component.Finish;
                if (component.GetPictureByteArray() != null)
                {
                    row["Pictures"] = component.GetPictureByteArray();
                }
                else
                {
                    row["Pictures"] = DBNull.Value;
                }
                row["Notes"] = component.Notes;
                row["Position"] = position++;
                row["Priority"] = component.Priority;
                row["TaskIDCount"] = component.TaskIDCount;

                dt.Rows.Add(row);
            }

            foreach (DataRow nrow in dt.Rows)
            {
                Console.WriteLine($"{nrow["JobNumber"]} {nrow["ProjectNumber"]} {nrow["Component"]} {nrow["Position"]}");
            }

            Console.WriteLine("Component DataTable Created.");

            return dt;
        }

        private DataTable CreateDataTableFromTaskList(ProjectModel project)
        {
            DataTable dt = new DataTable();
            int i;

            dt.Columns.Add("ProjectNumber", typeof(int));
            dt.Columns.Add("JobNumber", typeof(string));
            dt.Columns.Add("Component", typeof(string));
            dt.Columns.Add("TaskID", typeof(int));
            dt.Columns.Add("TaskName", typeof(string));
            dt.Columns.Add("Duration", typeof(string));
            dt.Columns.Add("StartDate", typeof(DateTime));
            dt.Columns.Add("FinishDate", typeof(DateTime));
            dt.Columns.Add("EarliestStartDate", typeof(DateTime));
            dt.Columns.Add("Predecessors", typeof(string));
            dt.Columns.Add("Machines", typeof(string));
            dt.Columns.Add("Machine", typeof(string));
            dt.Columns.Add("Resources", typeof(string));
            dt.Columns.Add("Resource", typeof(string));
            dt.Columns.Add("Hours", typeof(int));
            dt.Columns.Add("ToolMaker", typeof(string));
            dt.Columns.Add("Operator", typeof(string));
            dt.Columns.Add("Priority", typeof(string));
            dt.Columns.Add("Status", typeof(string));
            dt.Columns.Add("DateAdded", typeof(DateTime));
            dt.Columns.Add("Notes", typeof(string));
            dt.Columns.Add("Initials", typeof(string));
            dt.Columns.Add("DateCompleted", typeof(string));

            foreach (ComponentModel component in project.Components)
            {
                i = 1;

                foreach (TaskModel task in component.Tasks)
                {
                    DataRow row = dt.NewRow();

                    row["ProjectNumber"] = project.ProjectNumber;
                    row["JobNumber"] = project.JobNumber;
                    row["Component"] = task.Component;
                    row["TaskID"] = i++;  // Task.ID
                    row["TaskName"] = task.TaskName;
                    row["Duration"] = task.Duration;
                    row["Hours"] = task.Hours;
                    row["ToolMaker"] = task.ToolMaker;
                    row["Predecessors"] = task.Predecessors;
                    row["Resource"] = task.Resource;
                    row["Machine"] = task.Machine;
                    row["Priority"] = task.Priority;
                    row["Notes"] = task.Notes;

                    dt.Rows.Add(row);
                    //Console.WriteLine(i++);
                }
            }

            foreach (DataRow nrow in dt.Rows)
            {
                Console.WriteLine($"{nrow["ProjectNumber"]} {nrow["JobNumber"]} {nrow["Component"]} {nrow["TaskID"]} {nrow["Duration"]}");
            }

            Console.WriteLine("Task DataTable Created.");
            return dt;
        }

        public DataTable LoadProjectToDataTable(ProjectModel project)
        {
            DataTable dt = new DataTable();
            int count = 0;
            int baseCount = 0;

            dt.Columns.Add("Component", typeof(string));
            dt.Columns.Add("TaskName", typeof(string));
            dt.Columns.Add("Location", typeof(string));
            dt.Columns.Add("Subject", typeof(string));
            dt.Columns.Add("StartDate", typeof(DateTime));
            dt.Columns.Add("FinishDate", typeof(DateTime));
            dt.Columns.Add("Predecessors", typeof(string));
            dt.Columns.Add("Notes", typeof(string));
            dt.Columns.Add("PercentComplete", typeof(int));
            //dt.Columns.Add("AptID", typeof(int));
            dt.Columns.Add("TaskID", typeof(int));
            dt.Columns.Add("NewTaskID", typeof(int));

            foreach (ComponentModel component in project.Components)
            {
                count++;
                baseCount = count;

                foreach (TaskModel task in component.Tasks)
                {
                    DataRow row = dt.NewRow();

                    row["Component"] = component.Component;
                    row["TaskID"] = ++count;
                    //row["TaskID"] = task.ID;
                    row["TaskName"] = task.TaskName;
                    row["Location"] = task.TaskName + " (" + task.Hours + " Hours)";
                    row["Subject"] = project.JobNumber + " #" + project.ProjectNumber + "-" + task.ID;
                    row["StartDate"] = task.StartDate;
                    row["FinishDate"] = task.FinishDate;
                    row["PercentComplete"] = GetPercentComplete(task.Status);
                    row["Predecessors"] = task.GetNewPredecessors(baseCount);
                    row["Notes"] = task.Notes;
                    row["NewTaskID"] = count;

                    dt.Rows.Add(row);
                }
            }

            return dt;
        }

        public bool EditProjectInDB(ProjectModel project)
        {
            try
            {
                ProjectModel databaseProject = GetProject(project.JobNumber, project.ProjectNumber);
                List<ComponentModel> newComponentList = new List<ComponentModel>();
                //List<ComponentModel> updatedComponentList = new List<ComponentModel>();
                List<TaskModel> newTaskList = new List<TaskModel>();
                List<ComponentModel> deletedComponentList = new List<ComponentModel>();

                if (project.ProjectNumberChanged && ProjectExists(project.ProjectNumber))
                {
                    MessageBox.Show("There is another project with that same project number. Enter a different project number.");
                    return false;
                }

                UpdateProjectData(project);

                if (ProjectHasComponents(project.ProjectNumber))
                {
                    // Check modified project for added components.
                    foreach (ComponentModel component in project.Components)
                    {
                        component.SetPosition(project.Components.IndexOf(component));

                        bool componentExists = databaseProject.Components.Exists(x => x.Component == component.Component);

                        if (componentExists)
                        {
                            UpdateComponentData(project, component);

                            if (component.ReloadTaskList)
                            {
                                RemoveTasks(project, component);
                                newTaskList.AddRange(component.Tasks);
                            }
                            else
                            {
                                UpdateTasks(project.JobNumber, project.ProjectNumber, component.Component, component.Tasks);
                            }
                        }
                        else
                        {
                            newComponentList.Add(component);
                            newTaskList.AddRange(component.Tasks);
                        }
                    }

                    // Check modified project for deleted components.
                    foreach (ComponentModel component in databaseProject.Components)
                    {
                        bool componentExists = project.Components.Exists(x => x.Component == component.Component);

                        if (!componentExists)
                        {
                            deletedComponentList.Add(component);
                        }
                    }

                    // Check modified project for updated tasks

                    if (newComponentList.Count > 0)
                    {
                        AddComponentDataTableToDatabase(CreateDataTableFromComponentList(project, newComponentList));
                    }

                    if (newTaskList.Count > 0)
                    {
                        AddTaskDataTableToDatabase(CreateDataTableFromTaskList(project, newTaskList));
                    }

                    if (deletedComponentList.Count > 0)
                    {
                        foreach (ComponentModel component in deletedComponentList)
                        {
                            RemoveComponent(project, component);
                        }
                    }
                }
                else
                {
                    AddComponentDataTableToDatabase(CreateDataTableFromComponentList(project));
                }

                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
                return false;
            }
        }

        public ProjectModel GetProject(int projectNumber)
        {
            ProjectModel project = GetProjectInfo(projectNumber);

            AddComponents(project);

            AddTasks(project);

            return project;
        }

        public ProjectModel GetProjectInfo(int projectNumber)
        {
            OleDbCommand cmd;
            ProjectModel pi = null;
            string queryString;

            queryString = "SELECT * FROM Projects WHERE ProjectNumber = @projectNumber";

            cmd = new OleDbCommand(queryString, Connection);
            cmd.Parameters.AddWithValue("@projectNumber", projectNumber);

            try
            {
                Connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            pi = new ProjectModel
                            (
                                        jobNumber: Convert.ToString(rdr["JobNumber"]),
                                    projectNumber: Convert.ToInt32(rdr["ProjectNumber"]),
                                          dueDate: Convert.ToDateTime(rdr["DueDate"]),
                                           status: Convert.ToString(rdr["Status"]),
                                        toolMaker: Convert.ToString(rdr["ToolMaker"]),
                                         designer: Convert.ToString(rdr["Designer"]),
                                  roughProgrammer: Convert.ToString(rdr["RoughProgrammer"]),
                               electrodProgrammer: Convert.ToString(rdr["ElectrodeProgrammer"]),
                                 finishProgrammer: Convert.ToString(rdr["FinishProgrammer"]),
                               kanBanWorkbookPath: Convert.ToString(rdr["KanBanWorkbookPath"])
                            );
                        }
                    }
                }

                Connection.Close();
            }
            catch (Exception e)
            {
                Connection.Close();
                MessageBox.Show(e.Message, "GetProjectInfo");
            }

            return pi;
        }

        private void UpdateProjectData(ProjectModel project)
        {
            try
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter();

                string queryString;

                queryString = "UPDATE Projects " +
                              "SET JobNumber = @jobNumber, ProjectNumber = @newProjectNumber, DueDate = @dueDate, Designer = @designer, ToolMaker = @toolMaker, RoughProgrammer = @roughProgrammer, ElectrodeProgrammer = @electrodeProgrammer, " +
                              "FinishProgrammer = @finishProgrammer " +
                              "WHERE ProjectNumber = @oldProjectNumber";

                adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

                adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", project.JobNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@newProjectNumber", project.ProjectNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@dueDate", project.DueDate);
                adapter.UpdateCommand.Parameters.AddWithValue("@designer", project.Designer);
                adapter.UpdateCommand.Parameters.AddWithValue("@toolMaker", project.ToolMaker);
                adapter.UpdateCommand.Parameters.AddWithValue("@roughProgrammer", project.RoughProgrammer);
                adapter.UpdateCommand.Parameters.AddWithValue("@electrodeProgrammer", project.ElectrodeProgrammer);
                adapter.UpdateCommand.Parameters.AddWithValue("@finishProgrammer", project.FinishProgrammer);
                adapter.UpdateCommand.Parameters.AddWithValue("@oldProjectNumber", project.OldProjectNumber);  // By default this number is set to whatever is in the database when it was loaded to the Edit project form.

                Connection.Open();

                adapter.UpdateCommand.ExecuteNonQuery();

                MessageBox.Show("Project Updated!");
            }
            catch (OleDbException ex)
            {
                throw ex;
            }
            catch (Exception x)
            {
                throw x;
            }
            finally
            {
                Connection.Close();
            }
        }

        private void UpdateTasks(string jobNumber, int projectNumber, string component, List<TaskModel> taskList)
        {
            try
            {
                DataTable dt = new DataTable();
                string queryString;

                //queryString = "UPDATE Projects " +
                //              "SET JobNumber = @jobNumber, ProjectNumber = @newProjectNumber, DueDate = @dueDate, Designer = @designer, ToolMaker = @toolMaker, RoughProgrammer = @roughProgrammer, ElectrodeProgrammer = @electrodeProgrammer, " +
                //              "FinishProgrammer = @finishProgrammer " +
                //              "WHERE ProjectNumber = @oldProjectNumber";

                queryString = "SELECT * " +
                              "FROM Tasks " +
                              "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component";

                var adapter = new OleDbDataAdapter(queryString, Connection);

                adapter.SelectCommand = new OleDbCommand(queryString, Connection);

                adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
                adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = projectNumber;
                adapter.SelectCommand.Parameters.Add("@component", OleDbType.VarChar, 35).Value = component;

                var cbr = new OleDbCommandBuilder(adapter);

                cbr.GetDeleteCommand();
                cbr.GetInsertCommand();

                adapter.UpdateCommand = cbr.GetUpdateCommand();

                adapter.Fill(dt);

                UpdateTaskDataTable(taskList, dt);

                Connection.Open();
                adapter.Update(dt);


                //MessageBox.Show($"{component} tasks updated!");
            }
            catch (OleDbException ex)
            {
                throw ex;
            }
            catch (Exception x)
            {
                throw x;
            }
            finally
            {
                Connection.Close();
            }
        }

        private DataTable UpdateTaskDataTable(List<TaskModel> taskList, DataTable taskDataTable)
        {
            int i = 0;

            foreach (DataRow nrow in taskDataTable.Rows)
            {
                nrow["Hours"] = taskList[i].Hours;
                nrow["Duration"] = taskList[i].Duration;
                nrow["Machine"] = taskList[i].Machine;
                nrow["Resource"] = taskList[i].Personnel;
                nrow["Predecessors"] = taskList[i].Predecessors;
                nrow["Priority"] = taskList[i].Priority;
                nrow["Notes"] = taskList[i].Notes;

                i++;
            }

            return taskDataTable;
        }

        public bool ProjectHasComponents(int projectNumber)
        {
            string queryString = "SELECT COUNT(*) FROM Components WHERE ProjectNumber = @projectNumber";

            OleDbCommand cmd = new OleDbCommand(queryString, Connection);

            cmd.Parameters.AddWithValue("@projectNumber", projectNumber);

            Connection.Open();
            var count = (int)cmd.ExecuteScalar();
            Connection.Close();

            if (count > 0)
            {
                return true;
            }

            return false;
        }

        private void UpdateComponentData(ProjectModel project, ComponentModel component)
        {
            try
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                string queryString;

                queryString = "UPDATE Components " +
                              "SET Component = @name, Notes = @notes, Priority = @priority, [Position] = @position, Quantity = @quantity, Spares = @spares, Pictures = @picture, Material = @material, Finish = @finish, TaskIDCount = @taskIDCount " +
                              "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @oldName";

                adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

                adapter.UpdateCommand.Parameters.AddWithValue("@name", component.Component);
                adapter.UpdateCommand.Parameters.AddWithValue("@notes", component.Notes);
                adapter.UpdateCommand.Parameters.AddWithValue("@priority", component.Priority);
                adapter.UpdateCommand.Parameters.AddWithValue("@position", component.Position);
                adapter.UpdateCommand.Parameters.AddWithValue("@quantity", component.Quantity);
                adapter.UpdateCommand.Parameters.AddWithValue("@spares", component.Spares);

                if (component.GetPictureByteArray() != null)
                {
                    adapter.UpdateCommand.Parameters.AddWithValue("@picture", component.GetPictureByteArray());
                }
                else
                {
                    adapter.UpdateCommand.Parameters.AddWithValue("@picture", DBNull.Value);
                }

                //adapter.UpdateCommand.Parameters.AddWithValue("@pictures", component.PictureList);  // Add when database is ready to receive pictures.

                adapter.UpdateCommand.Parameters.AddWithValue("@material", component.Material);
                adapter.UpdateCommand.Parameters.AddWithValue("@finish", component.Finish);
                adapter.UpdateCommand.Parameters.AddWithValue("@taskIDCount", component.TaskIDCount);

                adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", project.JobNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", project.ProjectNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@oldName", component.OldName);

                Connection.Open();

                adapter.UpdateCommand.ExecuteNonQuery();

                Console.WriteLine($"{component.Component} Updated.");
                //MessageBox.Show("Project Updated!");
            }
            catch (OleDbException ex)
            {
                throw ex;
            }
            catch (Exception x)
            {
                throw x;
            }
            finally
            {
                Connection.Close();
            }
        }

        private DataTable CreateDataTableFromComponentList(ProjectModel project, List<ComponentModel> componentList)
        {
            DataTable dt = new DataTable();
            int position = 0;

            dt.Columns.Add("JobNumber", typeof(string));
            dt.Columns.Add("ProjectNumber", typeof(int));
            dt.Columns.Add("Component", typeof(string));
            dt.Columns.Add("Notes", typeof(string));
            dt.Columns.Add("Position", typeof(int));
            dt.Columns.Add("Priority", typeof(int));
            dt.Columns.Add("Pictures", typeof(byte[]));
            dt.Columns.Add("Material", typeof(string));
            dt.Columns.Add("Finish", typeof(string));
            dt.Columns.Add("TaskIDCount", typeof(int));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("Spares", typeof(int));
            dt.Columns.Add("Status", typeof(string));
            dt.Columns.Add("PercentComplete", typeof(int));

            foreach (ComponentModel component in componentList)
            {
                DataRow row = dt.NewRow();

                row["JobNumber"] = project.JobNumber;
                row["ProjectNumber"] = project.ProjectNumber;
                row["Component"] = component.Component;
                row["Notes"] = component.Notes;
                row["Priority"] = component.Priority;
                row["Position"] = component.Position;
                row["Quantity"] = component.Quantity;
                row["Spares"] = component.Spares;
                row["Material"] = component.Material;
                row["Finish"] = component.Finish;
                row["TaskIDCount"] = component.TaskIDCount;

                dt.Rows.Add(row);
            }

            foreach (DataRow nrow in dt.Rows)
            {
                Console.WriteLine($"{nrow["JobNumber"]} {nrow["ProjectNumber"]} {nrow["Component"]} {nrow["Position"]}");
            }

            Console.WriteLine("Component DataTable Created.");

            return dt;
        }

        private DataTable CreateDataTableFromTaskList(ProjectModel project, List<TaskModel> taskList)
        {
            DataTable dt = new DataTable();
            string component = "";
            int i = 1;

            dt.Columns.Add("ProjectNumber", typeof(int));
            dt.Columns.Add("JobNumber", typeof(string));
            dt.Columns.Add("Component", typeof(string));
            dt.Columns.Add("TaskID", typeof(int));
            dt.Columns.Add("TaskName", typeof(string));
            dt.Columns.Add("Duration", typeof(string));
            dt.Columns.Add("StartDate", typeof(DateTime));
            dt.Columns.Add("FinishDate", typeof(DateTime));
            dt.Columns.Add("EarliestStartDate", typeof(DateTime));
            dt.Columns.Add("Predecessors", typeof(string));
            dt.Columns.Add("Machines", typeof(string));
            dt.Columns.Add("Machine", typeof(string));
            dt.Columns.Add("Resources", typeof(string));
            dt.Columns.Add("Resource", typeof(string));
            dt.Columns.Add("Hours", typeof(int));
            dt.Columns.Add("ToolMaker", typeof(string));
            dt.Columns.Add("Operator", typeof(string));
            dt.Columns.Add("Priority", typeof(string));
            dt.Columns.Add("Status", typeof(string));
            dt.Columns.Add("DateAdded", typeof(DateTime));
            dt.Columns.Add("Notes", typeof(string));
            dt.Columns.Add("Initials", typeof(string));
            dt.Columns.Add("DateCompleted", typeof(string));

            foreach (TaskModel task in taskList)
            {
                DataRow row = dt.NewRow();

                if (component != task.Component)
                {
                    component = task.Component;
                    i = 1;
                }

                row["ProjectNumber"] = project.ProjectNumber;
                row["JobNumber"] = project.JobNumber;
                row["Component"] = task.Component;
                row["TaskID"] = i++;  // task.ID;  // TODO: Need to change this so that task ID's show up in order in the database and predecessors are referencing the correct tasks.
                row["TaskName"] = task.TaskName;
                row["Duration"] = task.Duration;
                row["Hours"] = task.Hours;
                row["ToolMaker"] = task.ToolMaker;
                row["Predecessors"] = task.Predecessors;
                row["Resource"] = task.Resource;
                row["Machine"] = task.Machine;
                row["Priority"] = task.Priority;
                row["Notes"] = task.Notes;

                dt.Rows.Add(row);
                //Console.WriteLine(i++);
            }

            foreach (DataRow nrow in dt.Rows)
            {
                Console.WriteLine($"{nrow["ProjectNumber"]} {nrow["JobNumber"]} {nrow["Component"]} {nrow["TaskID"]} {nrow["Duration"]}");
            }

            Console.WriteLine("Task DataTable Created.");
            return dt;
        }

        public void RemoveTasks(ProjectModel project, ComponentModel component)
        {
            var adapter = new OleDbDataAdapter();

            adapter.DeleteCommand = new OleDbCommand("DELETE FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component", Connection);

            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = project.JobNumber;
            adapter.DeleteCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = project.ProjectNumber;
            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 35).Value = component.Component;

            Connection.Open();
            adapter.DeleteCommand.ExecuteNonQuery();
            Connection.Close();
        }

        private void RemoveComponent(ProjectModel project, ComponentModel component)
        {
            var adapter = new OleDbDataAdapter();

            adapter.DeleteCommand = new OleDbCommand("DELETE FROM Components WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component", Connection);

            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = project.JobNumber;
            adapter.DeleteCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = project.ProjectNumber;
            adapter.DeleteCommand.Parameters.Add("@component", OleDbType.VarChar, 35).Value = component.Component;

            Connection.Open();
            adapter.DeleteCommand.ExecuteNonQuery();
            Connection.Close();
        }

        private int GetPercentComplete(string status)
        {
            if(status == "Completed")
            {
                return 100;
            }
            else
            {
                return 0;
            }
        }

        public DataTable GetProjectResourceData(ProjectModel project)
        {
            DataTable dt = new DataTable();
            int i = 1;
            int parentID = 0;

            dt.Columns.Add("NewTaskID", typeof(int));
            dt.Columns.Add("TaskName", typeof(string));
            dt.Columns.Add("ParentID", typeof(int));

            foreach (ComponentModel component in project.Components)
            {
                DataRow newRow1 = dt.NewRow();

                newRow1["NewTaskID"] = i;
                newRow1["TaskName"] = component.Component;
                parentID = i++;

                dt.Rows.Add(newRow1);

                foreach (TaskModel task in component.Tasks)
                {
                    DataRow newRow2 = dt.NewRow();
                    
                    newRow2["NewTaskID"] = i;
                    newRow2["TaskName"] = task.TaskName;
                    newRow2["ParentID"] = parentID;

                    Console.WriteLine(newRow2["TaskName"].ToString());

                    dt.Rows.Add(newRow2);

                    i++;
                }
            }

            return dt;
        }

        public List<string> GetResourceList(string role)
        {
            List<string> ResourceList = new List<string>();
            DataTable dt = new DataTable();

            string queryString = "SELECT DISTINCT Resources.ResourceName From Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID WHERE Role = @role OR Role LIKE @role ORDER BY Resources.ResourceName ASC";

            Stopwatch sw = new Stopwatch();
            sw.Start();

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.AddWithValue("@role", "%" + role + "%");

            adapter.Fill(dt);

            ResourceList.Add("");

            foreach (DataRow nrow in dt.Rows)
            {
                ResourceList.Add($"{nrow["ResourceName"]}");
            }

            Console.WriteLine($"GetResourceList Transaction Time: {sw.Elapsed}");

            return ResourceList;
        }

        public List<string> GetResourceList()
        {
            DataTable dt = new DataTable();
            List<string> ResourceList = new List<string>();

            string queryString = "SELECT * From Resources ORDER BY ResourceName ASC";

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            foreach (DataRow nrow in dt.Rows)
            {
                ResourceList.Add(nrow["ResourceName"].ToString());
                //Console.WriteLine($"Added: {nrow["FirstName"]} {nrow["LastName"]} {nrow["Role"]}");
            }

            return ResourceList;
        }

        private int GetResourceID(string resourceName)
        {
            DataTable dt = new DataTable();
            string queryString = "SELECT ID FROM Resources WHERE ResourceName = @resourceName";

            OleDbCommand sqlCommand = new OleDbCommand(queryString, Connection);

            sqlCommand.Parameters.AddWithValue("@resourceName", resourceName);

            Connection.Open();
            int resourceID = (int)sqlCommand.ExecuteScalar();
            Connection.Close();

            return resourceID;
        }

        public List<string> GetRoleList(string role)
        {
            DataTable dt = new DataTable();
            List<string> RoleList = new List<string>();

            string queryString = "SELECT Resources.ResourceName, Roles.Role FROM Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID WHERE Roles.Role = @role ORDER BY Resources.ResourceName ASC";

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.AddWithValue("@role", role);

            adapter.Fill(dt);

            foreach (DataRow nrow in dt.Rows)
            {
                RoleList.Add(nrow["ResourceName"].ToString());
                //Console.WriteLine($"Added: {nrow["FirstName"]} {nrow["LastName"]} {nrow["Role"]}");
            }

            return RoleList;
        }

        public void InsertResource(string resourceName)
        {
            if (resourceName != "")
            {
                OleDbCommand cmd = new OleDbCommand("INSERT INTO Resources (ResourceName) VALUES (@resourceName)", Connection);

                cmd.Parameters.AddWithValue("resourceID", resourceName);

                Connection.Open();
                cmd.ExecuteNonQuery();
                Connection.Close();
            }
            else
            {
                MessageBox.Show("You have not entered a name for a resource to add.");
            }
        }

        public void RemoveResource(string resourceName)
        {

            if (resourceName != "")
            {
                OleDbCommand cmd1 = new OleDbCommand("DELETE FROM Roles WHERE ResourceID = @resourceID ", Connection);
                OleDbCommand cmd2 = new OleDbCommand("DELETE FROM Resources WHERE ID = @resourceID ", Connection);

                cmd1.Parameters.AddWithValue("resourceID", GetResourceID(resourceName));
                cmd2.Parameters.AddWithValue("ID", GetResourceID(resourceName));

                Connection.Open();
                cmd1.ExecuteNonQuery();
                cmd2.ExecuteNonQuery();
                Connection.Close();
            }
            else
            {
                MessageBox.Show("You have not selected a resource to remove.");
            }
        }

        public void InsertResourceRole(string resourceName, string role)
        {
            if (resourceName != "")
            {
                OleDbCommand cmd = new OleDbCommand("INSERT INTO Roles (ResourceID, Role) VALUES (@resourceID, @role)", Connection);

                cmd.Parameters.AddWithValue("@resourceID", GetResourceID(resourceName));
                cmd.Parameters.AddWithValue("@role", role);

                Connection.Open();
                cmd.ExecuteNonQuery();
                Connection.Close();
            }
            else
            {
                MessageBox.Show("You have not selected a resource to add a role to.");
            }
        }

        public void RemoveResourceRole(string resourceName, string role)
        {
            if (resourceName != "")
            {
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Roles WHERE ResourceID = @resourceID AND Role = @role", Connection);

                cmd.Parameters.AddWithValue("resourceID", GetResourceID(resourceName));
                cmd.Parameters.AddWithValue("@role", role);

                Connection.Open();
                cmd.ExecuteNonQuery();
                Connection.Close();
            }
            else
            {
                MessageBox.Show("You have not selected a resource to add a role to.");
            }
        }

        public List<string> GetMachineList(string machineType)
        {
            List<string> machineList = new List<string>();
            DataTable dt = new DataTable();

            string queryString = "SELECT MachineName From Machines WHERE MachineType LIKE @machineType ORDER BY MachineName ASC";

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.AddWithValue("@machineType", "%" + machineType + "%");

            adapter.Fill(dt);

            foreach (DataRow nrow in dt.Rows)
            {
                machineList.Add($"{nrow["MachineName"]}");
            }

            return machineList;
        }

        public DataTable GetDependencyData(DataTable taskTable)
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("ParentId", typeof(int));
            dt.Columns.Add("DependentId", typeof(int));

            //taskIDKey = createTaskIDKey(taskTable);

            //Console.WriteLine("Get Dependency Data");

            foreach (DataRow nrow in taskTable.Rows)
            {
                if (nrow["Predecessors"].ToString().Contains(","))
                {
                    foreach (string predecessor in nrow["Predecessors"].ToString().Split(','))
                    {
                        DataRow row = dt.NewRow();

                        row["DependentId"] = nrow["TaskID"];
                        row["ParentId"] = Convert.ToInt32(predecessor);

                        dt.Rows.Add(row);

                        //Console.WriteLine($"{nrow["TaskID"]} {predecessor}");
                    }
                }
                else if (nrow["Predecessors"].ToString() != "")
                {
                    DataRow row = dt.NewRow();

                    row["DependentId"] = nrow["TaskID"];
                    row["ParentId"] = Convert.ToInt32(nrow["Predecessors"]);

                    dt.Rows.Add(row);

                    //Console.WriteLine($"{nrow["TaskID"]} {nrow["Predecessors"]}");
                }

            }

            foreach (DataRow nrow in dt.Rows)
            {
                //Console.WriteLine(nrow["ParentId"].ToString() + " " + nrow["DependentId"]);
            }

            return dt;
        }

        private bool ProjectExists(int projectNumber)
        {
            OleDbCommand sqlCommand = new OleDbCommand("SELECT COUNT(*) from Projects WHERE ProjectNumber = @projectNumber", Connection);

            sqlCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

            Connection.Open();
            int projectCount = (int)sqlCommand.ExecuteScalar();
            Connection.Close();

            if (projectCount > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        // Only need to delete the project from projects since the Database is set to cascade delete related records.
        public bool RemoveProject(string jobNumber, int projectNumber)
        {
            var adapter = new OleDbDataAdapter();

            adapter.DeleteCommand = new OleDbCommand("DELETE FROM Projects WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", Connection);
            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
            adapter.DeleteCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = projectNumber;

            Connection.Open();
            adapter.DeleteCommand.ExecuteNonQuery();
            Connection.Close();

            if (!ProjectExists(projectNumber))
            {
                MessageBox.Show("Project Deleted.");
                return true;
            }
            else
            {
                MessageBox.Show("Project Deletion Failed.");
                return false;
            }
        }

        public void UpdateProjectsTable(object s, CellValueChangedEventArgs ev)
        {
            try
            {
                var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

                //queryString = "UPDATE Tasks SET JobNumber = @jobNumber, Component = @component, TaskID = @taskID, TaskName = @taskName, " +
                //    "Duration = @duration, StartDate = @startDate, FinishDate = @finishDate, Predecessor = @predecessor, Machines = @machines, " +
                //    "Machine = @machine, Person = @person, Priority = @priority WHERE ID = @tID";

                using (Connection)
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = Connection;

                    if (ev.Column.FieldName == "JobNumber")
                    {
                        cmd.CommandText = "UPDATE Projects SET JobNumber = @jobNumber WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@jobNumber", ev.Value.ToString());
                    }
                    //else if (ev.Column.FieldName == "ProjectNumber")
                    //{
                    //    cmd.CommandText = "UPDATE WorkLoad SET ProjectNumber = @projectNumber WHERE (ID = @tID)";

                    //    if (ev.Value.ToString() != "")
                    //    {
                    //        cmd.Parameters.AddWithValue("@projectNumber", ev.Value.ToString());
                    //    }
                    //    else
                    //    {
                    //        cmd.Parameters.AddWithValue("@projectNumber", "");
                    //    }
                    //}
                    else if (ev.Column.FieldName == "Customer")
                    {
                        cmd.CommandText = "UPDATE Projects SET Customer = @customer WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@customer", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Project")
                    {
                        cmd.CommandText = "UPDATE Projects SET Project = @project WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@project", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "DueDate")
                    {
                        cmd.CommandText = "UPDATE Projects SET DueDate = @dueDate WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@dueDate", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@dueDate", DBNull.Value);
                        }
                    }
                    else if (ev.Column.FieldName == "Status")
                    {
                        cmd.CommandText = "UPDATE Projects SET Status = @status WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@status", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Designer")
                    {
                        cmd.CommandText = "UPDATE Projects SET Designer = @designer WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@designer", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "ToolMaker")
                    {
                        cmd.CommandText = "UPDATE Projects SET ToolMaker = @toolMaker WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@toolMaker", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "RoughProgrammer")
                    {
                        cmd.CommandText = "UPDATE Projects SET RoughProgrammer = @roughProgrammer WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@roughProgrammer", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "FinishProgrammer")
                    {
                        cmd.CommandText = "UPDATE Projects SET FinishProgrammer = @finishProgrammer WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@finishProgrammer", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "ElectrodeProgrammer")
                    {
                        cmd.CommandText = "UPDATE Projects SET ElectrodeProgrammer = @electrodeProgrammer WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@electrodeProgrammer", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "KanBanWorkbookPath")
                    {
                        cmd.CommandText = "UPDATE Projects SET KanBanWorkbookPath = @kanBanWorkbookPath WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@kanBanWorkbookPath", ev.Value.ToString());
                    }

                    cmd.Parameters.AddWithValue("@tID", (grid.GetRowCellValue(ev.RowHandle, grid.Columns["ID"])));

                    Connection.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Connection.Close();
            }
        }

        public void UpdateComponentsTable(object s, CellValueChangedEventArgs ev)
        {
            try
            {
                var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

                //queryString = "UPDATE Tasks SET JobNumber = @jobNumber, Component = @component, TaskID = @taskID, TaskName = @taskName, " +
                //    "Duration = @duration, StartDate = @startDate, FinishDate = @finishDate, Predecessor = @predecessor, Machines = @machines, " +
                //    "Machine = @machine, Person = @person, Priority = @priority WHERE ID = @tID";

                using (Connection)
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = Connection;

                    if (ev.Column.FieldName == "Component")
                    {
                        cmd.CommandText = "UPDATE Components SET Component = @component WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@component", ev.Value);
                    }
                    else if (ev.Column.FieldName == "Pictures")
                    {
                        cmd.CommandText = "UPDATE Components SET Pictures = @pictures WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@pictures", ev.Value);
                    }
                    //else if (ev.Column.FieldName == "ProjectNumber")
                    //{
                    //    cmd.CommandText = "UPDATE WorkLoad SET ProjectNumber = @projectNumber WHERE (ID = @tID)";

                    //    if (ev.Value.ToString() != "")
                    //    {
                    //        cmd.Parameters.AddWithValue("@projectNumber", ev.Value.ToString());
                    //    }
                    //    else
                    //    {
                    //        cmd.Parameters.AddWithValue("@projectNumber", "");
                    //    }
                    //}
                    else if (ev.Column.FieldName == "Material")
                    {
                        cmd.CommandText = "UPDATE Components SET Material = @material WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@material", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Finish")
                    {
                        cmd.CommandText = "UPDATE Components SET Finish = @finish WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@finish", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Notes")
                    {
                        cmd.CommandText = "UPDATE Components SET Notes = @notes WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@notes", ev.Value.ToString());

                    }
                    else if (ev.Column.FieldName == "Priority")
                    {
                        cmd.CommandText = "UPDATE Components SET Priority = @priority WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@priority", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@priority", 0);
                        }
                    }
                    else if (ev.Column.FieldName == "Quantity")
                    {
                        cmd.CommandText = "UPDATE Components SET Quantity = @quantity WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@quantity", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@quantity", 0);
                        }
                    }
                    else if (ev.Column.FieldName == "Spares")
                    {
                        cmd.CommandText = "UPDATE Components SET Spares = @spares WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@spares", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@spares", 0);
                        }
                    }
                    else if (ev.Column.FieldName == "Status")
                    {
                        cmd.CommandText = "UPDATE Components SET Status = @status WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@status", ev.Value.ToString());
                    }

                    cmd.Parameters.AddWithValue("@tID", (grid.GetRowCellValue(ev.RowHandle, grid.Columns["ID"])));

                    Connection.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Connection.Close();
            }
        }

        public void UpdateTasksTable(object s, CellValueChangedEventArgs ev)
        {
            try
            {
                var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

                //queryString = "UPDATE Tasks SET JobNumber = @jobNumber, Component = @component, TaskID = @taskID, TaskName = @taskName, " +
                //    "Duration = @duration, StartDate = @startDate, FinishDate = @finishDate, Predecessor = @predecessor, Machines = @machines, " +
                //    "Machine = @machine, Person = @person, Priority = @priority WHERE ID = @tID";

                using (Connection)
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = Connection;

                    //if (ev.Column.FieldName == "TaskName")
                    //{
                    //    cmd.CommandText = "UPDATE Tasks SET TaskName = @taskName WHERE (ID = @tID)";

                    //    cmd.Parameters.AddWithValue("@taskName", ev.Value.ToString());
                    //}
                    //else if (ev.Column.FieldName == "ProjectNumber")
                    //{
                    //    cmd.CommandText = "UPDATE WorkLoad SET ProjectNumber = @projectNumber WHERE (ID = @tID)";

                    //    if (ev.Value.ToString() != "")
                    //    {
                    //        cmd.Parameters.AddWithValue("@projectNumber", ev.Value.ToString());
                    //    }
                    //    else
                    //    {
                    //        cmd.Parameters.AddWithValue("@projectNumber", "");
                    //    }
                    //}
                    if (ev.Column.FieldName == "Notes")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Notes = @notes WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@notes", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Hours")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Hours = @hours WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@hours", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@hours", 0);
                        }
                    }
                    else if (ev.Column.FieldName == "Duration")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Duration = @duration WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@duration", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Predecessors")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Predecessors = @predecessors WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@predecessors", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Machine")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Machine = @machine WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@machine", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Resource")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Resource = @resource WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@resource", ev.Value.ToString());
                    }
                    else
                    {
                        MessageBox.Show(ev.Column.ToString() + " column is not editable.");
                        return;
                    }

                    cmd.Parameters.AddWithValue("@tID", (grid.GetRowCellValue(ev.RowHandle, grid.Columns["ID"])));

                    Connection.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Connection.Close();
            }
        }

        public void UpdateWorkloadTable(object s, CellValueChangedEventArgs ev)
        {
            try
            {
                var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

                //queryString = "UPDATE Tasks SET JobNumber = @jobNumber, Component = @component, TaskID = @taskID, TaskName = @taskName, " +
                //    "Duration = @duration, StartDate = @startDate, FinishDate = @finishDate, Predecessor = @predecessor, Machines = @machines, " +
                //    "Machine = @machine, Person = @person, Priority = @priority WHERE ID = @tID";

                using (Connection)
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = Connection;

                    if (ev.Column.FieldName == "ToolNumber")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET ToolNumber = @toolNumber WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@toolNumber", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@toolNumber", "");
                        }
                    }
                    else if (ev.Column.FieldName == "MWONumber")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET MWONumber = @mwoNumber WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@mwoNumber", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@mwoNumber", "");
                        }
                    }
                    else if (ev.Column.FieldName == "ProjectNumber")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET ProjectNumber = @projectNumber WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@projectNumber", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@projectNumber", "");
                        }
                    }
                    else if (ev.Column.FieldName == "Stage")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET Stage = @stage WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@stage", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@stage", "");
                        }
                    }
                    else if (ev.Column.FieldName == "Customer")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET Customer = @customer WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@customer", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@customer", "");
                        }
                    }
                    else if (ev.Column.FieldName == "PartName")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET PartName = @partName WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@partName", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@partName", "");
                        }
                    }
                    else if (ev.Column.FieldName == "Engineer")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET Engineer = @engineer WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@engineer", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@engineer", "");
                        }
                    }
                    else if (ev.Column.FieldName == "DeliveryInWeeks")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET DeliveryInWeeks = @deliveryInWeeks WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@deliveryInWeeks", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@deliveryInWeeks", "0");
                        }
                    }
                    else if (ev.Column.FieldName == "StartDate")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET StartDate = @startDate WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@startDate", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@startDate", DBNull.Value);
                        }
                    }
                    else if (ev.Column.FieldName == "FinishDate")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET FinishDate = @finishDate WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@finishDate", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@finishDate", DBNull.Value);
                        }
                    }
                    else if (ev.Column.FieldName == "AdjustedDeliveryDate")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET AdjustedDeliveryDate = @adjustedDeliveryDate WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@adjustedDeliveryDate", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@adjustedDeliveryDate", DBNull.Value);
                        }
                    }
                    else if (ev.Column.FieldName == "MoldCost")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET MoldCost = @moldCost WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@moldCost", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@moldCost", "0");
                        }
                    }
                    else if (ev.Column.FieldName == "Designer")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET Designer = @designer WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@designer", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@designer", "");
                        }
                    }
                    else if (ev.Column.FieldName == "Designer")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET Designer = @designer WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@designer", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@designer", "");
                        }
                    }
                    else if (ev.Column.FieldName == "ToolMaker")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET ToolMaker = @toolMaker WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@toolMaker", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@toolMaker", "");
                        }
                    }
                    else if (ev.Column.FieldName == "RoughProgrammer")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET RoughProgrammer = @roughProgrammer WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@roughProgrammer", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@roughProgrammer", "");
                        }
                    }
                    else if (ev.Column.FieldName == "FinishProgrammer")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET FinishProgrammer = @finishProgrammer WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@finishProgrammer", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@finishProgrammer", "");
                        }
                    }
                    else if (ev.Column.FieldName == "ElectrodeProgrammer")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET ElectrodeProgrammer = @electrodeProgrammer WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@electrodeProgrammer", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@electrodeProgrammer", "");
                        }
                    }
                    else if (ev.Column.FieldName == "Manifold")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET Manifold = @manifold WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@manifold", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@manifold", "");
                        }
                    }
                    else if (ev.Column.FieldName == "MoldBase")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET MoldBase = @moldBase WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@moldBase", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@moldBase", "");
                        }
                    }
                    else if (ev.Column.FieldName == "GeneralNotes")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET GeneralNotes = @generalNotes WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@generalNotes", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@generalNotes", "");
                        }
                    }
                    else if (ev.Column.FieldName == "GeneralNotesRTF")
                    {
                        cmd.CommandText = "UPDATE WorkLoad SET GeneralNotesRTF = @generalNotesRTF WHERE (ID = @tID)";

                        if (ev.Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@generalNotesRTF", ev.Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@generalNotesRTF", "");
                        }
                    }

                    cmd.Parameters.AddWithValue("@tID", (grid.GetRowCellValue(ev.RowHandle, grid.Columns["ID"])));

                    //Console.WriteLine(connectionString);
                    //Console.WriteLine(queryString);
                    //Console.WriteLine((grid.Rows[ev.RowIndex]).Cells[0].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[1].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[2].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[3].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[4].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[5].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[6].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[7].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[8].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[9].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[10].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[11].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[12].Value.ToString() + " ");
                    Connection.Open();
                    cmd.ExecuteNonQuery();

                    {
                        //if(ev.ColumnIndex != 5)
                        //    MessageBox.Show("Update Success!");
                        Connection.Close();
                    }

                }
            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message);
            }
        }

        public void AddWorkLoadEntry(WorkLoadModel wli)
        {
            try
            {
                OleDbCommand cmd = new OleDbCommand("INSERT INTO WorkLoad (ToolNumber, MWONumber, ProjectNumber, Stage, Customer, Project, DeliveryInWeeks, StartDate, FinishDate, AdjustedDeliveryDate, MoldCost, Engineer, Designer, ToolMaker, RoughProgrammer, FinishProgrammer, ElectrodeProgrammer, Manifold, MoldBase, GeneralNotes) VALUES " + 
                                                                       "(@toolNumber, @mwoNumber, @projectNumber, @stage, @customer, @project, @deliveryInWeeks, @startDate, @finishDate, @adjustedDeliveryDate, @moldCost, @engineer, @designer, @toolMaker, @roughProgrammer, @finishProgrammer, @electrodeProgrammer, @manifold, @moldBase, @generalNotes)", Connection);

                cmd.Parameters.AddWithValue("@toolNumber", wli.ToolNumber);
                cmd.Parameters.AddWithValue("@mwoNumber", wli.MWONumber);
                cmd.Parameters.AddWithValue("@projectNumber", wli.ProjectNumber);
                cmd.Parameters.AddWithValue("@stage", wli.Stage);
                cmd.Parameters.AddWithValue("@customer", wli.Customer);
                cmd.Parameters.AddWithValue("@project", wli.Project);
                cmd.Parameters.AddWithValue("@deliveryInWeeks", wli.DeliveryInWeeks);

                if(wli.StartDate != null)
                {
                    cmd.Parameters.AddWithValue("@startDate", wli.StartDate);
                }
                else
                {
                    cmd.Parameters.AddWithValue("@startDate", DBNull.Value);
                }

                if (wli.FinishDate != null)
                {
                    cmd.Parameters.AddWithValue("@finishDate", wli.FinishDate);
                }
                else
                {
                    cmd.Parameters.AddWithValue("@finishDate", DBNull.Value);
                }

                if (wli.AdjustedDeliveryDate != null)
                {
                    cmd.Parameters.AddWithValue("@adjustedDeliveryDate", wli.AdjustedDeliveryDate);
                }
                else
                {
                    cmd.Parameters.AddWithValue("@adjustedDeliveryDate", DBNull.Value);
                }

                cmd.Parameters.AddWithValue("@moldCost", wli.MoldCost);
                cmd.Parameters.AddWithValue("@engineer", wli.Engineer);
                cmd.Parameters.AddWithValue("@designer", wli.Designer);
                cmd.Parameters.AddWithValue("@toolMaker", wli.ToolMaker);
                cmd.Parameters.AddWithValue("@roughProgrammer", wli.RoughProgrammer);
                cmd.Parameters.AddWithValue("@finishProgrammer", wli.FinishProgrammer);
                cmd.Parameters.AddWithValue("@electrodeProgrammer", wli.ElectrodeProgrammer);
                cmd.Parameters.AddWithValue("@manifold", wli.Manifold);
                cmd.Parameters.AddWithValue("@moldBase", wli.MoldBase);
                cmd.Parameters.AddWithValue("@generalNotes", wli.GeneralNotes);

                Connection.Open();
                cmd.ExecuteNonQuery();
                Connection.Close();
            }
            catch (Exception er)
            {
                Connection.Close();
                MessageBox.Show(er.Message);
            }
        }

        public bool DeleteWorkLoadEntry(int id)
        {
            OleDbCommand cmd = new OleDbCommand("DELETE FROM WorkLoad WHERE ID = @id", Connection);

            cmd.Parameters.AddWithValue("@id", id);

            Connection.Open();
            cmd.ExecuteNonQuery();
            Connection.Close();

            return true;
        }

        public void AddColorEntry(int projectID, string column, int aRGBColor)
        {
            try
            {
                OleDbCommand cmd = new OleDbCommand("INSERT INTO WorkLoadColors (ProjectID, ColumnFieldName, ARGBColor) VALUES (@projectID, @columnFieldName, @aRGBColor)", Connection);

                cmd.Parameters.AddWithValue("@projectID", projectID);
                cmd.Parameters.AddWithValue("@columnFieldName", column);
                cmd.Parameters.AddWithValue("@aRGBColor", aRGBColor);


                Connection.Open();
                cmd.ExecuteNonQuery();
                Connection.Close();

            }
            catch (Exception e)
            {
                Connection.Close();
                MessageBox.Show(e.Message);
            }
        }

        public void UpdateColorEntry(int projectID, string column, int aRGBColor)
        {
            using (Connection)
            {
                try
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = Connection;

                    cmd.CommandText = "UPDATE WorkLoadColors SET ARGBColor = @aRGBColor WHERE (ProjectID = @projectID AND ColumnFieldName = @column)";

                    cmd.Parameters.AddWithValue("@aRGBColor", aRGBColor);
                    cmd.Parameters.AddWithValue("@projectID", projectID);
                    cmd.Parameters.AddWithValue("@column", column);

                    Connection.Open();
                    cmd.ExecuteNonQuery();
                    Connection.Close();
                }
                catch (Exception e)
                {
                    Connection.Close();
                    MessageBox.Show(e.Message);
                }
            }
        }

        public void DeleteColorEntries(int projectID)
        {
            OleDbCommand cmd = new OleDbCommand("DELETE FROM WorkLoadColors WHERE ProjectID = @projectID", Connection);

            cmd.Parameters.AddWithValue("@projectID", projectID);

            Connection.Open();
            cmd.ExecuteNonQuery();
            Connection.Close();
        }

        public List<ColorStruct> GetColorEntries()
        {
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM WorkLoadColors", Connection);
            List<ColorStruct> colorList = new List<ColorStruct>();

            try
            {
                Connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            colorList.Add(new ColorStruct
                            (
                                   projectID: rdr["ProjectID"],
                                      column: rdr["ColumnFieldName"],
                                   aRGBColor: rdr["ARGBColor"]
                            ));
                        }
                    }
                }

                Connection.Close();

                return colorList;
            }
            catch (Exception e)
            {
                Connection.Close();
                MessageBox.Show(e.Message);
                return null;
            }
        }

        public string GetKanBanWorkbookPath(string jobNumber, int projectNumber)
        {
            string kanBanWorkbookPath;
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbCommand sqlCommand = new OleDbCommand("SELECT KanBanWorkbookPath from Projects WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", Connection);

            sqlCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            sqlCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

            try
            {
                Connection.Open();

                object kanBanWorkbookPathObject = sqlCommand.ExecuteScalar();

                if (kanBanWorkbookPathObject != null && kanBanWorkbookPathObject != DBNull.Value)
                {
                    kanBanWorkbookPath = kanBanWorkbookPathObject.ToString();
                }
                else
                {
                    kanBanWorkbookPath = "";
                }
            }
            finally
            {
                Connection.Close();
                Connection.Dispose();
            }


            return kanBanWorkbookPath;
        }

        public void SetKanBanWorkbookPath(string path, string jn, int pn)
        {
            try
            {
                string queryString = "UPDATE Projects SET KanBanWorkbookPath = @path " +
                                     "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber";

                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

                adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

                adapter.UpdateCommand.Parameters.AddWithValue("@path", path);
                adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", jn);
                adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", pn);

                Connection.Open();
                adapter.UpdateCommand.ExecuteNonQuery();

            }
            finally
            {
                Connection.Close();
            }
        }
    }
}
