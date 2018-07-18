using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using DevExpress.Spreadsheet;
using System.Globalization;
using System.Runtime.InteropServices;
using Toolroom_Scheduler;

namespace Toolroom_Project_Viewer
{
    class Database
    {
        static readonly string ConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=X:\TOOLROOM\Workload Tracking System\Database\Workload Tracking System DB.accdb";

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

                foreach (DataRow nrow in dt.Rows)
                {
                    nrow["ID"] = i++;
                    Console.WriteLine($"{nrow["ID"]} {nrow["Subject"]} {nrow["Location"]} {nrow["StartDate"]}");
                }
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

                string predecessors = GetTaskPredecessors(jobNumber, projectNumber, taskID);

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

        public string GetTaskPredecessors(string jobNumber, int projectNumber, int taskID)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable dt = new DataTable();
            string predecessors = "";
            string queryString;

            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND TaskID = @taskID";
            OleDbConnection Connection = new OleDbConnection(ConnString);
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);

            adapter.SelectCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            adapter.SelectCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
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

        public void MoveDescendents(string jn, int pn, string component, DateTime currentTaskFinishDate, int currentTaskID)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;

            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jn AND ProjectNumber = @pn AND Component = @component ORDER BY TaskID ASC";
            OleDbConnection Connection = new OleDbConnection(ConnString);
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            adapter.SelectCommand.Parameters.Add("@jn", OleDbType.VarChar, 20).Value = jn;
            adapter.SelectCommand.Parameters.Add("@pn", OleDbType.Integer, 12).Value = pn;
            adapter.SelectCommand.Parameters.Add("@jn", OleDbType.VarChar, 30).Value = component;
            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

            Console.WriteLine("Move Descendents");

            adapter.Fill(datatable);

            UpdateStartAndFinishDates(currentTaskID, datatable, currentTaskFinishDate);

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
            string selectStatment = "ID, JobNumber & ' #' & ProjectNumber & ' ' & Component & '-' & TaskID As Subject, TaskName & ' (' & Hours & ' Hours)' As Location, StartDate, FinishDate, Machine, Resource, ToolMaker, Notes";

            if (department == "Design")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Design%'";
            }
            else if (department == "Program Rough")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Rough'";
            }
            else if (department == "Program Finish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Finish'";
            }
            else if (department == "Program Electrodes")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Electrodes'";
            }
            else if (department == "CNC Rough")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Rough'";
            }
            else if (department == "CNC Finish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Finish'";
            }
            else if (department == "CNC Electrodes")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Electrodes'";
            }
            else if (department == "EDM Sinker")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'EDM Sinker'";
            }
            else if (department == "Inspection")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE 'Inspection%'";
            }
            else if (department == "Grind")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Grind%'";
            }
            else if (department == "Polish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Polish%'";
            }
            else if (department == "All")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks";
            }

            return queryString;
        }

        private string SetMonthlyHoursQueryString(string department)
        {
            string queryString = null;
            string selectStatment = "TOP 3, MONTH(StartDate) as mo, YEAR(StartDate) AS yr, SUM(Hours) AS total";
            string fromStatement = "Tasks";
            string whereStatement = "AND MONTH(StartDate) >= MONTH(DATE()) AND YEAR(StartDate) >= YEAR(DATE())";
            string orderByStatement = "GROUP BY YEAR(StartDate), MONTH(StartDate) ORDER BY YEAR(StartDate), MONTH(StartDate)";

            if (department == "Design")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Design%' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Program Rough")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Rough' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Program Finish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Finish' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Program Electrodes")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'Program Electrodes' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "CNC Rough")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Rough' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "CNC Finish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Finish' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "CNC Electrodes")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'CNC Electrodes' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "EDM Sinker")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName = 'EDM Sinker' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Inspection")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE 'Inspection%' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Grind")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Grind%' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Polish")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Polish%' " + whereStatement + " " + orderByStatement;
            }
            else if (department == "All")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE  " + whereStatement + " " + orderByStatement;
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

        public DataTable GetAppointmentDataList()
        {
            DataTable dt = new DataTable();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            //string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resources FROM Tasks WHERE TaskName LIKE 'CNC Finish'";
            string queryString = "SELECT JobNumber & ' #' & ProjectNumber & ' ' & Component As Subject, TaskName As Location, StartDate, FinishDate, Machine, Resource, Resources, ToolMaker, Notes FROM Tasks";
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            //TODO: #1 Create a list of objects containing appointment data.

            foreach(DataRow nrow in dt.Rows)
            {

            }

            return dt;
        }

        public DataTable GetNextThreeMonthsHours(string department, string timeUnits)
        {
            DataTable dt = new DataTable();
            string queryString = "";

            if(timeUnits == "Months")
            {
                queryString = SetMonthlyHoursQueryString(department);
            }
            else if(timeUnits == "Weeks")
            {
                
            }

            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            foreach (DataRow nrow in dt.Rows)
            {
                Console.WriteLine($"{nrow["yr"].ToString()} {nrow["mo"].ToString()} {nrow["total"].ToString()} ");
            }

            return dt;
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

                foreach (Day day in week.DayList)
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

                            Console.WriteLine($"{rdr["JobNumber"].ToString()}-{rdr["ProjectNumber"].ToString()} {rdr["TaskName"].ToString()} {rdr["Duration"].ToString()} {Convert.ToDateTime(rdr["StartDate"]).ToShortDateString()} {Convert.ToDateTime(rdr["FinishDate"]).ToShortDateString()} {rdr["Hours"].ToString()}");

                            double hours = Convert.ToInt32(rdr["Hours"]);
                            double days = (int)GetBusinessDays(Convert.ToDateTime(rdr["StartDate"]), Convert.ToDateTime(rdr["FinishDate"]));
                            DateTime date = Convert.ToDateTime(rdr["StartDate"]);
                            decimal dailyAVG = (decimal)(hours / days);

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

        //public bool ProjectHasDates()
        //{
        //    string queryString;

        //    queryString = "";
        //}

        public ProjectInfo GetProject(string jobNumber, int projectNumber)
        {
            ProjectInfo project = GetProjectInfo(jobNumber, projectNumber);

            AddComponents(project);

            AddTasks(project);

            return project;
        }

        public ProjectInfo GetProjectInfo(string jobNumber, int projectNumber)
        {
            OleDbCommand cmd;
            ProjectInfo pi = null;
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
                            pi = new ProjectInfo
                            (
                                    jn: Convert.ToString(rdr["JobNumber"]),
                                    pn: Convert.ToInt32(rdr["ProjectNumber"]),
                                    dd: Convert.ToDateTime(rdr["DueDate"]),
                                    tm: Convert.ToString(rdr["ToolMaker"]),
                                     d: Convert.ToString(rdr["Designer"]),
                                    rp: Convert.ToString(rdr["RoughProgrammer"]),
                                    ep: Convert.ToString(rdr["ElectrodeProgrammer"]),
                                    fp: Convert.ToString(rdr["FinishProgrammer"])
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

        public void AddComponents(ProjectInfo project)
        {
            OleDbCommand cmd;
            Component component;

            string queryString;

            queryString = "SELECT * FROM Components WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber";
            OleDbConnection Connection = new OleDbConnection(ConnString);
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
                            component = new Component
                            (
                                    name: rdr["Component"],
                                    notes: rdr["Notes"],
                                    priority: rdr["Priority"],
                                    position: rdr["Position"],
                                    material: rdr["Material"],
                                    taskIDCount: rdr["TaskIDCount"]
                            );

                            project.AddComponent(component);
                        }

                        Connection.Close();
                        Connection.Dispose();
                    }
                    else
                    {
                        Connection.Close();
                        Connection.Dispose();

                        project.AddComponentList(GetComponentListFromTasksTable(project.JobNumber, project.ProjectNumber));
                    }
                }

            }
            catch (Exception e)
            {
                Connection.Close();
                Connection.Dispose();
                MessageBox.Show(e.Message, "AddComponents");
            }
        }

        public List<Component> GetComponentListFromTasksTable(string jobNumber, int projectNumber)
        {
            OleDbCommand cmd;
            List<Component> componentList = new List<Component>();

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
                            componentList.Add(new Component
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

        private void AddTasks(ProjectInfo project)
        {
            List<TaskInfo> projectTaskList = GetProjectTaskList(project.JobNumber, project.ProjectNumber);

            foreach (Component component in project.ComponentList)
            {
                var tasks = from t in projectTaskList
                            where t.Component == component.Name
                            orderby t.ID ascending
                            select t;

                component.AddTaskList(tasks.ToList());
            }

            foreach (Component component in project.ComponentList)
            {
                Console.WriteLine(component.Name);

                foreach (TaskInfo task in component.TaskList)
                {
                    Console.WriteLine($"   {task.ID} {task.TaskName}");
                }
            }
        }

        public List<TaskInfo> GetProjectTaskList(string jobNumber, int projectNumber)
        {
            OleDbCommand cmd;
            List<TaskInfo> taskList = new List<TaskInfo>();

            string queryString;
            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber";
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
                            taskList.Add(new TaskInfo
                            (
                                    taskName: rdr["TaskName"],
                                          id: rdr["TaskID"],
                                   component: rdr["Component"],
                                       hours: rdr["Hours"],
                                    duration: rdr["Duration"],
                                   startDate: rdr["StartDate"],
                                  finishDate: rdr["FinishDate"],
                                      status: rdr["Status"],
                                     machine: rdr["Machine"],
                                   personnel: rdr["Resource"],
                                predecessors: rdr["Predecessors"],
                                       notes: rdr["Notes"]
                            ));
                        }
                    }
                }

                Connection.Close();
            }
            catch (Exception e)
            {
                Connection.Close();
                throw e;
            }

            return taskList;
        }

        public DataTable LoadProjectToDataTable(ProjectInfo project)
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

            foreach (Component component in project.ComponentList)
            {
                count++;
                baseCount = count;

                foreach (TaskInfo task in component.TaskList)
                {
                    DataRow row = dt.NewRow();

                    row["Component"] = component.Name;
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

        public DataTable GetProjectData(string jobNumber, int projectNumber)
        {
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            int i = 1;
            string component = null;

            string queryString = "SELECT Component, TaskID, TaskName, JobNumber & ' #' & ProjectNumber & ' ' & Component As Subject, TaskName As Location, StartDate, FinishDate, Predecessors " +
                                 "FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber " +
                                 "ORDER BY TaskID";
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            adapter.SelectCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

            adapter.Fill(dt);

            dt.Columns.Add("NewTaskID", typeof(int));

            foreach (DataRow nrow in dt.Rows)
            {
                if(component != nrow["Component"].ToString())
                {
                    component = nrow["Component"].ToString();

                    if(nrow["Component"].ToString() != "")
                    {
                        i++;
                    }
                        
                }

                nrow["TaskID"] = i;
                nrow["NewTaskID"] = i;
                i++;
            }

            Console.WriteLine("Get Project Data");

            foreach (DataRow nrow in dt.Rows)
            {
                Console.WriteLine($"{nrow["NewTaskID"]} {nrow["TaskName"]}");
            }

            return dt;
        }

        public DataTable GetProjectResourceData(string jobNumber, int projectNumber)
        {
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            int i = 1;
            int parentID = 0;
            string component = null;

            string queryString = "SELECT Component, TaskID, TaskName From Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber ORDER BY TaskID";
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            adapter.SelectCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

            adapter.Fill(dt1);

            dt2.Columns.Add("TaskName", typeof(string));
            dt2.Columns.Add("NewTaskID", typeof(Int32));
            dt2.Columns.Add("ParentID", typeof(Int32));

            foreach (DataRow nrow in dt1.Rows)
            {
                if (component != nrow["Component"].ToString())
                {
                    component = nrow["Component"].ToString();

                    if(nrow["Component"].ToString() != "")
                    {
                        DataRow newRow1 = dt2.NewRow();
                        parentID = i;
                        newRow1["NewTaskID"] = i;
                        newRow1["TaskName"] = component;

                        dt2.Rows.Add(newRow1);

                        i++;
                    }
                }

                DataRow newRow2 = dt2.NewRow();
                //Console.WriteLine(nrow["TaskName"].ToString());
                newRow2["NewTaskID"] = i;
                newRow2["ParentID"] = parentID;
                newRow2["TaskName"] = nrow["TaskName"];

                dt2.Rows.Add(newRow2);

                i++;
            }

            return dt2;
        }

        public DataTable GetProjectResourceData(ProjectInfo project)
        {
            DataTable dt = new DataTable();
            int i = 1;
            int parentID = 0;

            dt.Columns.Add("NewTaskID", typeof(int));
            dt.Columns.Add("TaskName", typeof(string));
            dt.Columns.Add("ParentID", typeof(int));

            foreach (Component component in project.ComponentList)
            {
                DataRow newRow1 = dt.NewRow();

                newRow1["NewTaskID"] = i;
                newRow1["TaskName"] = component.Name;
                parentID = i++;

                dt.Rows.Add(newRow1);

                foreach (TaskInfo task in component.TaskList)
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

        public DataTable GetProjectResourceData(DataTable taskTable)
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("TaskID");
            dt.Columns.Add("TaskName");

            Console.WriteLine("Get Project Resource Data");

            foreach (DataRow nrow in taskTable.Rows)
            {
                DataRow row = dt.NewRow();

                row["TaskID"] = GetNewTaskID(Convert.ToInt32(nrow["TaskID"]));
                row["TaskName"] = nrow["TaskName"];
            }

            return dt;
        }

        public DataTable GetDependencyData(string jobNumber, int projectNumber)
        {
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();

            //string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resources FROM Tasks WHERE TaskName LIKE 'CNC Finish'";
            string queryString = "SELECT TaskID, Predecessors " +
                                 "FROM Tasks " +
                                 "WHERE JobNumber LIKE @jobNumber AND ProjectNumber LIKE @projectNumber";

            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            adapter.SelectCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

            adapter.Fill(dt1);

            dt2.Columns.Add("ParentId", typeof(int));
            dt2.Columns.Add("DependentId", typeof(int));

            foreach (DataRow nrow in dt1.Rows)
            {
                if (nrow["Predecessors"].ToString().Contains(","))
                {
                    foreach (string predecessor in nrow["Predecessors"].ToString().Split(','))
                    {
                        DataRow row = dt2.NewRow();

                        row["DependentId"] = Convert.ToInt32(nrow["TaskId"]);
                        row["ParentId"] = Convert.ToInt32(predecessor);

                        dt2.Rows.Add(row);
                    }
                }
                else if (nrow["Predecessors"].ToString() != "")
                {
                    DataRow row = dt2.NewRow();

                    row["DependentId"] = nrow["TaskId"];
                    row["ParentId"] = nrow["Predecessors"];

                    dt2.Rows.Add(row);
                }

            }

            foreach (DataRow nrow in dt2.Rows)
            {
                Console.WriteLine(nrow["ParentId"].ToString() + " " + nrow["DependentId"]);
            }

            return dt2;
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

        private DataTable CreateTaskIDKey(DataTable taskTable)
        {
            DataTable dt = new DataTable();
            int i = 1;

            dt.Columns.Add("NewTaskID", typeof(int));
            dt.Columns.Add("OldTaskID", typeof(int));

            //Console.WriteLine("Create Task Key");

            foreach (DataRow nrow in taskTable.Rows)
            {
                DataRow row = dt.NewRow();

                row["NewTaskID"] = i;
                row["OldTaskID"] = nrow["TaskID"];

                //Console.WriteLine($"{row["NewTaskID"]} {row["OldTaskID"]}");

                dt.Rows.Add(row);
                i++;
            }
            
            return dt;
        }

        private int GetNewTaskID(int id)
        {
            DataRow selectedRow;

            selectedRow = (DataRow)taskIDKey.Rows.Cast<DataRow>().Where(r => r.Field<int>("OldTaskID") == id).First();

            //Console.WriteLine("Get New Task ID");
            //Console.WriteLine($"{id} {selectedRow["NewTaskID"]} ");

            return (int)selectedRow["NewTaskID"];
        }

        private DataTable GetTranslatedTaskIDTable(DataTable projectResourceDataTable)
        {
            DataTable dt = new DataTable();
            DataTable keyTable = CreateTaskIDKey(projectResourceDataTable);

            dt.Columns.Add("TaskID", typeof(int));
            dt.Columns.Add("TaskName", typeof(string));

            foreach (DataRow nrow in projectResourceDataTable.Rows)
            {
                DataRow row = dt.NewRow();

                //var results = from DataRow myRow in keyTable.Rows
                //              where (int)myRow["OldTaskID"] == (int)nrow["TaskID"]
                //              select myRow;

                DataRow selectedRow = (DataRow)keyTable.Rows.Cast<DataRow>().Where(r => (int)r["OldTaskID"] == (int)nrow["TaskID"]);


                row["TaskID"] = selectedRow["NewTaskID"];
                row["TaskName"] = nrow["TaskName"];

                dt.Rows.Add(row);
            }

            return dt;
        }

        public void OpenKanBanWorkbook(string filepath, string component)
        {
            //Excel.Worksheet ws;

            if (filepath != null)
            {
                FileInfo fi = new FileInfo(filepath);

                if (fi.Exists)
                {
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook = excelApp.Workbooks.Open(fi.FullName);

                    try
                    {
                        //var attributes = File.GetAttributes(fi.FullName);    

                        foreach (Excel.Worksheet ws in workbook.Worksheets)
                        {
                            if (ws.Name.Trim() == component)
                            {
                                workbook.Sheets[ws.Index].Select();
                                workbook.Save();
                            }
                        }

                        workbook.Close();
                        excelApp.Quit();

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        Marshal.ReleaseComObject(workbook);

                        //Marshal.ReleaseComObject(ws);
                        Marshal.ReleaseComObject(excelApp);

                        var res = Process.Start("EXCEL.EXE", "\"" + fi.FullName + "\"");

                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);

                        //workbook.Close();
                        excelApp.Quit();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        Marshal.ReleaseComObject(workbook);

                        //Marshal.ReleaseComObject(ws);
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                else
                {
                    MessageBox.Show("Can't find a Kan Ban Workbook with path " + filepath + ".");
                }
            }
            else
            {
                MessageBox.Show("There is no Kan Ban Workbook for this project.");
            }

        }

        public string GetKanBanWorkbookPath(string jn, int pn)
        {
            string kanBanWorkbookPath = "";
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbCommand sqlCommand = new OleDbCommand("SELECT KanBanWorkbookPath from Projects WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", Connection);

            sqlCommand.Parameters.AddWithValue("@jobNumber", jn);
            sqlCommand.Parameters.AddWithValue("@projectNumber", pn);

            try
            {
                Connection.Open();
                kanBanWorkbookPath = (string)sqlCommand.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Connection.Close();
            Connection.Dispose();

            return kanBanWorkbookPath;
        }

        //public DataTable getTaskData()
        //{

        //}
    }
}
