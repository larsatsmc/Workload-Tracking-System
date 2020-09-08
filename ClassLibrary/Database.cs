using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using DevExpress.XtraScheduler;
using Dapper;
using System.Data.SqlClient;

namespace ClassLibrary
{
    public class Database
    {
        // WorkloadTrackingSystemDB, LocalSqlServerDB See App.config for list of connection names.
        static readonly string DatabaseType = "SQL Server"; // Either 'Access' or 'SQL Server'.
        static readonly string SQLClientConnectionName = "SQLServerToolRoomSchedulerDB";  // LocalSqlServerDB
        static readonly string OLEDBConnectionName = "LocalOLEDBSqlServerDB";

        #region Projects Table Operations

        #region Create

        public static bool CreateProject(ProjectModel project)
        {
            if (ProjectExists(project.ProjectNumber))
            {
                MessageBox.Show("There is another project with that same project number. Enter a different project number");
            }
            else
            {
                using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    string queryString1 = "INSERT INTO Projects (JobNumber, ProjectNumber, Customer, Project, DueDate, Priority, Designer, ToolMaker, RoughProgrammer, ElectrodeProgrammer, FinishProgrammer, EDMSinkerOperator, RoughCNCOperator, ElectrodeCNCOperator, FinishCNCOperator, EDMWireOperator, Apprentice, OverlapAllowed, IncludeHours, DateCreated, DateModified) " + // 
                                          "VALUES (@JobNumber, @ProjectNumber, @Customer, @Project, @DueDate, @Priority, @Designer, @ToolMaker, @RoughProgrammer, @ElectrodeProgrammer, @FinishProgrammer, @EDMSinkerOperator, @RoughCNCOperator, @ElectrodeCNCOperator, @FinishCNCOperator, @EDMWireOperator, @Apprentice, @OverlapAllowed, @IncludeHours, GETDATE(), GETDATE())"; // 

                    string queryString2 = "INSERT INTO Components (JobNumber, ProjectNumber, Component, Notes, Priority, [Position], Material, TaskIDCount, Quantity, Spares, Picture, Finish) " + // 
                                          "VALUES (@JobNumber, @ProjectNumber, @Component, @Notes, @Priority, @Position, @Material, @TaskIDCount, @Quantity, @Spares, @Picture, @Finish)"; // 

                    string queryString3 = "INSERT INTO Tasks (JobNumber, ProjectNumber, Component, TaskID, TaskName, Duration, StartDate, FinishDate, Predecessors, Machine, Resources, Personnel, Hours, Priority, Notes) " +
                                          "VALUES (@JobNumber, @ProjectNumber, @Component, @TaskID, @TaskName, @Duration, @StartDate, @FinishDate, @Predecessors, @Machine, @Resources, @Personnel, @Hours, @Priority, @Notes)";

                    connection.Open();

                    using (var trans = connection.BeginTransaction())
                    {
                        // OleDBConnection doesn't like it when I feed the object directly into the DynamicParameters constructor.
                        //var p1 = new DynamicParameters(project);

                        var p1 = new
                        {
                            project.JobNumber,
                            project.ProjectNumber,
                            project.Customer,
                            project.Project,
                            project.DueDate,
                            project.Priority,
                            project.Designer,
                            project.ToolMaker,
                            project.RoughProgrammer,
                            project.ElectrodeProgrammer,
                            project.FinishProgrammer,
                            project.EDMSinkerOperator,
                            project.RoughCNCOperator,
                            project.ElectrodeCNCOperator,
                            project.FinishCNCOperator,
                            project.EDMWireOperator,
                            project.Apprentice,
                            project.OverlapAllowed,
                            project.IncludeHours
                        };

                        connection.Execute(queryString1, p1, trans);

                        foreach (ComponentModel component in project.Components)
                        {
                            var p2 = new
                            {
                                project.JobNumber,
                                project.ProjectNumber,
                                component.Component,
                                component.Notes,
                                component.Priority,
                                component.Position,
                                component.Material,
                                component.TaskIDCount,
                                component.Quantity,
                                component.Spares,
                                component.Picture,
                                component.Finish,
                                component.Status
                            };

                            Console.WriteLine($"{project.JobNumber} {project.ProjectNumber} {component.Component}");

                            connection.Execute(queryString2, p2, trans);

                            foreach (TaskModel task in component.Tasks)
                            {
                                var p3 = new 
                                { 
                                    project.JobNumber,
                                    project.ProjectNumber,
                                    component.Component,
                                    task.TaskID,
                                    task.TaskName,
                                    task.Duration,
                                    task.StartDate,
                                    task.FinishDate,
                                    task.Predecessors,
                                    task.Machine,
                                    task.Resources,
                                    task.Personnel,
                                    task.Hours,
                                    task.Priority,
                                    task.Notes
                                };

                                connection.Execute(queryString3, p3, trans);

                                //MessageBox.Show(connection.ExecuteScalar("SELECT @@IDENTITY", transaction: trans).ToString());

                                task.ID = int.Parse(connection.ExecuteScalar("SELECT @@IDENTITY", transaction: trans).ToString());
                            }
                        }

                        trans.Commit();
                    }

                    return true;
                }
            }

            return false;
        }

        #endregion

        #region Read

        public static (List<ProjectModel> projects, List<ComponentModel> components, List<TaskModel> tasks) GetProjects()
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                List<ProjectModel> projects = connection.Query<ProjectModel>("SELECT * FROM Projects").ToList();
                List<ComponentModel> components = connection.Query<ComponentModel>("SELECT * FROM Components").ToList();
                List<TaskModel> tasks = connection.Query<TaskModel>("dbo.spGetAllTasks").ToList();

                return (projects, components, tasks);
            }
        }
        public static List<ProjectModel> GetAllProjects()
        {
            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                List<ProjectModel> projects = connection.Query<ProjectModel>("SELECT * FROM Projects").ToList();

                return projects;
            }
        }
        public static bool ProjectExists(int projectNumber)
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                int projectCount = int.Parse(connection.ExecuteScalar("SELECT COUNT(*) from Projects WHERE ProjectNumber = @ProjectNumber", new { ProjectNumber = projectNumber } ).ToString());
                
                if (projectCount > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public static ProjectModel GetProject(int projectNumber)
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                var p = new { ProjectNumber = projectNumber };

                ProjectModel project = connection.QueryFirst<ProjectModel>("dbo.spGetProject @ProjectNumber", p);
                List<ComponentModel> components = connection.Query<ComponentModel>("dbo.spGetProjectComponents @ProjectNumber", p).ToList();
                List<TaskModel> tasks = connection.Query<TaskModel>("dbo.spGetProjectTasks @ProjectNumber", p).ToList();

                project.Components = components;

                foreach (var component in project.Components)
                {
                    component.Tasks = tasks.FindAll(x => x.Component == component.Component).OrderBy(x => x.TaskID).ToList();

                    component.Tasks.ForEach(x => x.HasInfo = true);
                }

                project.HasProjectInfo = true;

                return project;
            }
        }

        public static List<ProjectModel> GetProjectInfoList()
        {
            string queryString = "SELECT * FROM Projects";
            SqlCommand cmd;
            ProjectModel pi;
            List<ProjectModel> piList = new List<ProjectModel>();

            try
            {
                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    cmd = new SqlCommand(queryString, connection);

                    connection.Open();

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
                                           apprentice: Convert.ToString(rdr["Apprentice"]),
                                   kanBanWorkbookPath: Convert.ToString(rdr["KanBanWorkbookPath"])
                                );

                                piList.Add(pi);
                            }
                        }
                    } 
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            return piList;
        }

        public static DataTable LoadProjectToDataTable(ProjectModel project)
        {
            DataTable dt = new DataTable();
            int count = 0;
            int baseCount;

            dt.Columns.Add("JobNumber", typeof(string));
            dt.Columns.Add("ProjectNumber", typeof(int));
            dt.Columns.Add("Component", typeof(string));
            dt.Columns.Add("TaskName", typeof(string));
            dt.Columns.Add("Hours", typeof(int));
            dt.Columns.Add("Location", typeof(string));
            dt.Columns.Add("Subject", typeof(string));
            dt.Columns.Add("StartDate", typeof(DateTime));
            dt.Columns.Add("FinishDate", typeof(DateTime));
            dt.Columns.Add("Predecessors", typeof(string));
            dt.Columns.Add("Notes", typeof(string));
            dt.Columns.Add("PercentComplete", typeof(int));
            dt.Columns.Add("AptID", typeof(int));
            dt.Columns.Add("TaskID", typeof(int));
            dt.Columns.Add("NewTaskID", typeof(int));

            foreach (ComponentModel component in project.Components)
            {
                count++;
                baseCount = count;

                foreach (TaskModel task in component.Tasks)
                {
                    DataRow row = dt.NewRow();

                    row["JobNumber"] = project.JobNumber;
                    row["ProjectNumber"] = project.ProjectNumber;
                    row["Component"] = component.Component;
                    row["AptID"] = ++count;
                    row["TaskID"] = task.TaskID;
                    row["TaskName"] = task.TaskName;
                    row["Hours"] = task.Hours;
                    //row["Location"] = task.TaskName + " (" + task.Hours + " Hours)";
                    row["Subject"] = $"{project.JobNumber} #{project.ProjectNumber}";
                    if (task.StartDate == null)
                    {
                        row["StartDate"] = DBNull.Value;
                    }
                    else
                    {
                        row["StartDate"] = task.StartDate;
                    }

                    if (task.FinishDate == null)
                    {
                        row["FinishDate"] = DBNull.Value;
                    }
                    else
                    {
                        row["FinishDate"] = task.FinishDate;
                    }
                    
                    row["PercentComplete"] = task.PercentComplete;
                    row["Predecessors"] = task.GetNewPredecessors(baseCount);
                    row["Notes"] = task.Notes;
                    row["NewTaskID"] = count;

                    dt.Rows.Add(row);
                }
            }

            return dt;
        }

        public string GetKanBanWorkbookPath(string jobNumber, int projectNumber)
        {
            string kanBanWorkbookPath;

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                SqlCommand sqlCommand = new SqlCommand("SELECT KanBanWorkbookPath from Projects WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", connection);

                sqlCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
                sqlCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
                connection.Open();

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

            return kanBanWorkbookPath;
        }

        #endregion

        #region Update

        public static bool UpdateWholeProject(ProjectModel project)
        {
            if (project.ProjectNumberChanged && ProjectExists(project.ProjectNumber))
            {
                MessageBox.Show("There is another project with that same project number. Enter a different project number.");
                return false;
            }

            ProjectModel databaseProject = GetProject(project.OldProjectNumber);

            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                List<ComponentModel> newComponentList = new List<ComponentModel>();
                //List<Component> updatedComponentList = new List<Component>();
                List<TaskModel> taskList = new List<TaskModel>();
                List<TaskModel> databaseTaskList = new List<TaskModel>();
                List<ComponentModel> deletedComponentList = new List<ComponentModel>();

                connection.Execute("dbo.spUpdateProject @JobNumber, @ProjectNumber, @Customer, @Project, @DueDate, @Status, @PercentComplete, @Designer, @ToolMaker, " +
                    "@RoughProgrammer, @ElectrodeProgrammer, @FinishProgrammer, @Apprentice, @EDMSinkerOperator, @RoughCNCOperator, @ElectrodeCNCOperator, @FinishCNCOperator, " +
                    "@EDMWireOperator, @OverlapAllowed, @IncludeHours, @KanBanWorkbookPath, @ID", project);

                var componentsToAdd = from component in project.Components
                                      where component.ID == 0
                                      select component;

                var componentsToUpdate = from component in project.Components
                                         where component.ID != 0
                                         select component;

                var componentsToRemove = from component in databaseProject.Components
                                         where !project.Components.Exists(x => x.ID == component.ID)
                                         select component;

                connection.Execute("dbo.spCreateComponent @JobNumber, @ProjectNumber, @Component, @Notes, @Priority, @Position, @Material, @TaskIDCount, @Quantity, @Spares, " +
                    "@Picture, @Finish, @Status, @PercentComplete", componentsToAdd.ToList());
                connection.Execute("dbo.spUpdateComponent @Component, @Notes, @Priority, @Position, @Quantity, @Spares, @Picture, @Material, @Finish, @TaskIDCount, @ID", componentsToUpdate.ToList());
                connection.Execute("DELETE FROM Components WHERE ID = @ID", componentsToRemove.ToList());

                int idIndex;

                foreach (ComponentModel component in project.Components)
                {
                    idIndex = 1;

                    foreach (TaskModel task in component.Tasks)
                    {
                        task.JobNumber = project.JobNumber;
                        task.ProjectNumber = project.ProjectNumber;
                        task.Component = component.Component;
                        task.TaskID = idIndex++;
                    }

                    taskList.AddRange(component.Tasks);
                }

                databaseProject.Components.ForEach(x => databaseTaskList.AddRange(x.Tasks));
                //taskList.ForEach(x => { x.ProjectNumber = project.ProjectNumber; x.JobNumber = project.JobNumber; });
                
                var tasksToAdd = from task in taskList
                                 where task.ID == 0
                                 select task;
                
                var tasksToUpdate = from task in taskList
                                    where task.ID != 0
                                    select task;
                
                var tasksToDelete = from task in databaseTaskList
                                    where !taskList.Exists(x => x.ID == task.ID)
                                    select task;

                // TODO: Verify Task object properties will map to stored procedure.
                connection.Execute("dbo.spCreateTask @JobNumber, @ProjectNumber, @Component, @TaskID, @TaskName, @Hours, @Duration, @Machine, @Resources, @Resource, @Predecessors, @Priority, @Notes", tasksToAdd.ToList());
                connection.Execute("dbo.spUpdateTask @TaskID, @TaskName, @Hours, @Duration, @Machine, @Resources, @Personnel, @Predecessors, @Priority, @Notes, @ID", tasksToUpdate.ToList());
                connection.Execute("DELETE FROM Tasks WHERE ID = @ID", tasksToDelete.ToList());
            }

            return true;
        }
        public static bool UpdateProjectRecord(ProjectModel project, CellValueChangedEventArgs ev)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    if (ev.Column.FieldName == "ProjectNumber" && connection.Execute("dbo.spGetProjectCount @ProjectNumber", project) > 0)
                    {
                        MessageBox.Show("There is a project with that same project number.");
                        return false;
                    }

                    string queryString = $"UPDATE Project SET {ev.Column.FieldName} = @{ev.Column.FieldName} WHERE ID = @ID";
                    
                    connection.Execute(queryString, project);

                    return true;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                throw e;
            }
        }
        public static bool UpdateProjectField(WorkLoadModel workLoad, CellValueChangedEventArgs ev)
        {
            try
            {
                string queryString, dateModified = "";

                if (workLoad.MWONumber != null)
                {
                    workLoad.ProjectNumber = workLoad.MWONumber;
                }

                List<string> CriticalProjectFieldsInKanBan = new List<string> { "JobNumber", "ProjectNumber" };

                if (CriticalProjectFieldsInKanBan.Contains(ev.Column.FieldName))
                {
                    dateModified = ", DateModified = GETDATE()";
                }

                using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    // This if statement will not be true so long as this method is only called when the programmer fields are changed.

                    if (ev.Column.FieldName == "ProjectNumber" && connection.Execute("dbo.spGetProjectCount @ProjectNumber", workLoad.ProjectNumber) > 0)
                    {
                        MessageBox.Show("There is a project with that same project number.");
                        return false;
                    }

                    queryString = $"UPDATE Projects SET {ev.Column.FieldName} = @{ev.Column.FieldName}{dateModified} WHERE ProjectNumber = @ProjectNumber";

                    connection.Execute(queryString, workLoad);

                    return true;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                throw e;
            }
        }

        public static void SetKanBanWorkbookPath(string path, int projectNumber)
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string queryString = "UPDATE Projects SET KanBanWorkbookPath = @Path, LastKanBanGenerationDate = GETDATE() " +
                                     "WHERE ProjectNumber = @ProjectNumber";

                var p = new { Path = path, ProjectNumber = projectNumber };

                connection.Execute(queryString, p);
            }

        }

        public static void SetJobFolderPath(int id, string jobFolderPath)
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string queryString = "UPDATE WorkLoad SET JobFolderPath = @JobFolderPath " +
                                     "WHERE ID = @ID";

                var p = new
                {
                    JobFolderPath = jobFolderPath,
                    ID = id
                };

                connection.Execute(queryString, p);
            }
        }

        #endregion

        #region Delete

        // Only need to delete the project from projects since the Database is set to cascade delete related records.
        public static bool RemoveProject(string jobNumber, int projectNumber)
        {
            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                //var adapter = new OleDbDataAdapter();

                //adapter.DeleteCommand = new OleDbCommand("DELETE FROM Projects WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", connection);
                //adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
                //adapter.DeleteCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = projectNumber;

                connection.Execute("DELETE FROM Projects WHERE ProjectNumber = @projectNumber", new { ProjectNumber = projectNumber });

                //Connection.Open();
                //adapter.DeleteCommand.ExecuteNonQuery();
                //Connection.Close();

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
        }

        #endregion

        #endregion // Project Operations

        #region Components Table Operations

        #region Create


        #endregion

        #region Read

        #endregion

        #region Update

        public static void UpdateComponent(ComponentModel component, CellValueChangedEventArgs ev)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    connection.Execute($"Update Components SET {ev.Column.FieldName} = @{ev.Column.FieldName} WHERE (ID = @ID)", component);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion

        #region Delete

        #endregion

        #endregion // Component operations.

        #region Tasks Table Operations

        #region Create

        #endregion

        #region Read

        public string GetTaskPredecessors(string jobNumber, int projectNumber, string component, int taskID)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            string predecessors = "";

            try
            {
                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    string queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";

                    adapter.SelectCommand = new SqlCommand(queryString, connection);
                    adapter.SelectCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
                    adapter.SelectCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
                    adapter.SelectCommand.Parameters.AddWithValue("@component", component);
                    adapter.SelectCommand.Parameters.AddWithValue("@taskID", taskID);

                    connection.Open();

                    using (var reader = adapter.SelectCommand.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            int ord = reader.GetOrdinal("Predecessors");
                            predecessors = reader.GetString(ord);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            //MessageBox.Show(projectInfo.Item1 + " " + projectInfo.Item2);

            return predecessors;
        }

        public static DateTime? GetFinishDate(string jobNumber, int projectNumber, string component, int taskID)
        {
            DateTime? FinishDate = null;

            try
            {
                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    SqlCommand sqlCommand = new SqlCommand("SELECT FinishDate from Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID", connection);

                    sqlCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
                    sqlCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
                    sqlCommand.Parameters.AddWithValue("@component", component);
                    sqlCommand.Parameters.AddWithValue("@taskID", taskID);

                    connection.Open();

                    FinishDate = (DateTime?)sqlCommand.ExecuteScalar(); 
                }
            }
            catch (Exception)
            {
                MessageBox.Show("A predecessor has no finish date.");
            }

            return FinishDate;
        }

        private string FindTask(DataTable dataTable, string jobNumber, int projectNumber, string component, int taskID)
        {
            DataRow task = dataTable.Rows.Cast<DataRow>().FirstOrDefault(x => (string)x["JobNumber"] == jobNumber && (int)x["ProjectNumber"] == projectNumber && (int)x["TaskID"] == taskID);

            return task["TaskName"].ToString();
        }

        // Helper method for setting appointment resources.
        public DataTable GetTasksWithChangedResources(int projectNumber, string taskName)
        {
            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                DataTable dt = new DataTable();

                string queryString = "SELECT * FROM Tasks WHERE ProjectNumber = @projectNumber AND TaskName = @taskName";

                SqlDataAdapter adapter = new SqlDataAdapter(queryString, connection);

                adapter.SelectCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
                adapter.SelectCommand.Parameters.AddWithValue("@taskName", taskName);

                adapter.Fill(dt);

                return dt;
            }
        }

        public static List<string> GetJobNumberComboList()
        {
            string queryString = "SELECT DISTINCT JobNumber, ProjectNumber FROM Tasks";
            DataTable dt = new DataTable();
            List<string> jobNumberList = new List<string>();

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                SqlDataAdapter adapter = new SqlDataAdapter(queryString, connection);

                adapter.Fill(dt);

                foreach (DataRow nrow in dt.Rows)
                {
                    //Console.WriteLine(nrow["JobNumber"]);
                    //
                    jobNumberList.Add($"{nrow["JobNumber"].ToString()} - #{nrow["ProjectNumber"].ToString()}");
                } 
            }

            return jobNumberList;
        }

        private static string SetWeeklyHoursQueryString(string weekStart, string weekEnd)
        {
            string department = "All";
            string queryString = null;
            string selectStatment = "Projects.JobNumber, Projects.ProjectNumber, TaskName, Duration, Tasks.StartDate, FinishDate, Personnel, Hours";
            //string fromStatement = "Tasks";
            string whereStatement = "(Tasks.StartDate BETWEEN '" + weekStart + "' AND '" + weekEnd + "') AND Projects.IncludeHours = 1";
            string orderByStatement = "ORDER BY Tasks.StartDate ASC";
            //string groupByStatement = "GROUP BY ";

            if (department == "All") // This if statement will always be true until method is changed.
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks INNER JOIN Projects ON Tasks.ProjectNumber = Projects.ProjectNumber WHERE  " + whereStatement + " " + orderByStatement;
            }
            else if (department == "Design")
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
                queryString = "SELECT " + selectStatment + " FROM Tasks INNER JOIN Projects ON Tasks.ProjectNumber = Projects.ProjectNumber WHERE  " + whereStatement + " " + orderByStatement;
            }

            return queryString;
        }

        public static List<Week> GetDayHours(string weekStart, string weekEnd)
        {
            List<Week> weeks = new List<Week>();

            string queryString = SetWeeklyHoursQueryString(weekStart, weekEnd);

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                SqlCommand cmd = new SqlCommand(queryString, connection);

                string[] departmentArr = { "Design", "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "CNC Electrodes", "EDM Sinker", "EDM Wire (In-House)", "Polish (In-House)", "Inspection", "Grind" };

                foreach (string item in departmentArr)
                {
                    weeks.Add(new Week(item));
                }

                connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            if (rdr["TaskName"].ToString() == "Design")
                            {
                                weeks[0].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString() == "Program Rough")
                            {
                                weeks[1].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString() == "Program Finish")
                            {
                                weeks[2].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString() == "Program Electrodes")
                            {
                                weeks[3].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString() == "CNC Rough")
                            {
                                weeks[4].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString() == "CNC Finish")
                            {
                                weeks[5].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString() == "CNC Electrodes")
                            {
                                weeks[6].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString() == "EDM Sinker")
                            {
                                weeks[7].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString() == "EDM Wire (In-House)")
                            {
                                weeks[8].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString() == "Polish (In-House)")
                            {
                                weeks[9].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString().Contains("Inspection"))
                            {
                                weeks[10].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }
                            else if (rdr["TaskName"].ToString().Contains("Grind"))
                            {
                                weeks[11].AddDayHours(Convert.ToInt16(rdr["Hours"]), Convert.ToDateTime(rdr["StartDate"]));
                            }

                            //Console.WriteLine($"{rdr["TaskName"]} {rdr["StartDate"]} {rdr["Hours"]}");
                        }
                    }
                    else
                    {

                    }


                } 
            }

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

        public static List<string> GetDepartments()
        {
            return new List<string>() { "Design", "Program Rough", "Program Finish", "Program Electrodes", "CNC Rough", "CNC Finish", "CNC Electrodes", "EDM Sinker", "EDM Wire (In-House)", "Polish (In-House)", "Inspection", "Grind" };
        }

        public static List<Week> GetWeekHours(string weekStart, string weekEnd, List<string> departmentList, string resourceType)
        {
            List<Week> weekList = new List<Week>();
            List<Week> deptWeekList = new List<Week>();
            //List<string> departmentList = new List<string>();
            Week weekTemp;
            DateTime wsDate = Convert.ToDateTime(weekStart);
            int weekNum;
            Stopwatch stopwatch = new Stopwatch();

            string queryString = SetWeeklyHoursQueryString(weekStart, weekEnd);

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                SqlCommand cmd = new SqlCommand(queryString, connection);

                weekList = InitializeDeptWeeksList(wsDate, departmentList);

                //Console.WriteLine("\nLoad");

                connection.Open();

                stopwatch.Start();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            //Console.WriteLine($"{++count} {rdr["JobNumber"].ToString()}-{rdr["ProjectNumber"].ToString()} {rdr["TaskName"].ToString()} {rdr["Duration"].ToString()} {rdr["Hours"].ToString()}");

                            if (resourceType == "Department")
                            {
                                var weeks = from wk in weekList
                                            where (rdr["TaskName"].ToString().StartsWith(wk.Department) || (rdr["TaskName"].ToString().Contains("Grind") && rdr["TaskName"].ToString().Contains(wk.Department))) // && Convert.ToDateTime(rdr["StartDate"]) >= wk.WeekStart && Convert.ToDateTime(rdr["StartDate"]) <= wk.WeekEnd
                                            orderby wk.WeekNum ascending
                                            select wk;

                                deptWeekList = weeks.ToList();
                            }
                            else if (resourceType == "Personnel")
                            {
                                var weeks = from wk in weekList
                                            where (rdr["Resource"].ToString().Contains(wk.Department)) // && Convert.ToDateTime(rdr["StartDate"]) >= wk.WeekStart && Convert.ToDateTime(rdr["StartDate"]) <= wk.WeekEnd
                                            orderby wk.WeekNum ascending
                                            select wk;

                                deptWeekList = weeks.ToList();
                            }



                            if (deptWeekList.Any())
                            {
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

                                            if (weekNum > 20)
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
                                            if (weekTemp.Department == "Design")
                                                Console.WriteLine($"{weekTemp.Department} {weekTemp.WeekStart.ToShortDateString()} {date.DayOfWeek} Daily AVG. {dailyAVG} Hrs {hours} Days {days}");
                                            days -= 1;
                                        }


                                        date = date.AddDays(1);
                                    }
                                }
                                else
                                {
                                    weekTemp.AddHoursToDay((int)date.AddDays(days).DayOfWeek, dailyAVG);
                                    if (weekTemp.Department == "Design")
                                        Console.WriteLine($"{weekTemp.Department} {weekTemp.WeekStart.ToShortDateString()} {date.AddDays(days).DayOfWeek} {dailyAVG} {days}");
                                }
                            }
                        }
                    }
                    else
                    {

                    }
                } 
            }

            TimeSpan ts = stopwatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            Console.WriteLine("RunTime " + elapsedTime);

            stopwatch.Stop();

            MyEnd:;

            //Console.WriteLine("\nReview:");

            //foreach (Week week in weekList)
            //{
            //    Console.WriteLine($"{week.Department} {week.GetWeekHours()} {week.WeekStart.ToShortDateString()} - {week.WeekEnd.ToShortDateString()}");
            //}

            return weekList;
        }

        private static string SetQueryString(string department)
        {
            string queryString = null;
            string selectStatment = "";

            if (DatabaseType == "Access")
            {
                selectStatment = "ID, JobNumber & ' #' & ProjectNumber & ' ' & Component As Subject, TaskName As Location, JobNumber, ProjectNumber, " +
                                 "TaskID, TaskName, Component, Hours, StartDate, FinishDate, Machine, Resources, Resource, Status, Notes, Predecessors";
            }
            else if (DatabaseType == "SQL Server")
            {
                selectStatment = "ID, CONCAT(JobNumber, ' #', ProjectNumber, ' ', Component) As Subject, TaskName As Location, JobNumber, ProjectNumber, " +
                                 "TaskID, TaskName, Component, Hours, StartDate, FinishDate, Machine, Resources, Resource, Status, Notes, Predecessors";
            }

            string orderByStatement = " ORDER BY StartDate ASC";

            if (department == "Design")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Design%'" + orderByStatement;
            }
            else if (department == "Programming")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE '%Program%'" + orderByStatement;
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
            else if (department == "CNC")
            {
                queryString = "SELECT " + selectStatment + " FROM Tasks WHERE TaskName LIKE 'CNC%'" + orderByStatement;
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
        public static List<TaskModel> GetAppointments(string department) // This is meant to replace the above GetAppointmentsData method.
        {
            string queryString = SetQueryString(department);

            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                List<TaskModel> appointments = connection.Query<TaskModel>(queryString).ToList();

                return appointments;
            }
        }
        public DataTable GetAppointmentData()
        {
            DataTable dt = new DataTable();

            try
            {
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    //string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resources FROM Tasks WHERE TaskName LIKE 'CNC Finish'";
                    string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resource, ToolMaker, Notes FROM Tasks WHERE TaskName = 'CNC Rough'";

                    OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

                    adapter.Fill(dt);
                }
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace, "OledbException Error");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\n" + e.StackTrace, "getAppointmentsData");
            }

            return dt;
        }

        #endregion

        #region Update
        public static void UpdateTask(TaskModel task)
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string queryString;

                queryString = "UPDATE Tasks SET StartDate = @StartDate, FinishDate = @finishDate, Machine = @Machine, Personnel = @Personnel, Resources = @Resources " +
                              "WHERE ID = @ID";

                connection.Execute(queryString, task);
            }
        }
        public static void UpdateTask(TaskModel task, CellValueChangedEventArgs ev)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    string queryString, resourceQueryString = "";

                    if (ev.Column.FieldName == "Machine" || ev.Column.FieldName == "Personnel")
                    {
                        resourceQueryString = ", Resources = @Resources";
                    }

                    queryString = $"UPDATE Tasks SET {ev.Column.FieldName} = @{ev.Column.FieldName}{resourceQueryString}, DateModified = GETDATE() WHERE (ID = @ID)";

                    connection.Execute(queryString, task);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        // This means both machines and personnel.
        public static void SetTaskResources(object s, CellValueChangedEventArgs ev, SchedulerStorage schedulerStorage)
        {
            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string taskName;
                int? projectNumber;

                var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

                //WorkLoadModel wli = grid.GetRow(ev.RowHandle) as WorkLoadModel;

                DataTable dt = new DataTable();

                //string queryString = "UPDATE Tasks SET Resource = @resource " +
                //                     "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND TaskName = @taskName";

                string queryString = "SELECT * FROM Tasks WHERE ProjectNumber = @projectNumber AND TaskName = @taskName";

                var command = new SqlCommand(queryString);

                command.Connection = connection;
                command.CommandType = CommandType.Text;

                if (grid.GetRowCellValue(ev.RowHandle, grid.Columns["MWONumber"]) != null) // If there is an MWONumber use it for the project number.
                {
                    projectNumber = (int)grid.GetRowCellValue(ev.RowHandle, grid.Columns["MWONumber"]);
                }
                else
                {
                    projectNumber = (int)grid.GetRowCellValue(ev.RowHandle, grid.Columns["ProjectNumber"]);
                }

                if (projectNumber == null)
                {
                    return;
                }

                command.Parameters.Add("@projectNumber", SqlDbType.Int, 12).Value = projectNumber;

                taskName = "Program " + ev.Column.FieldName.Remove(ev.Column.FieldName.Length - 10, 10);

                if (taskName.Contains("Electrode"))
                {
                    taskName = taskName + "s";
                }

                command.Parameters.Add("@taskName", SqlDbType.VarChar, 20).Value = taskName;

                SqlDataAdapter adapter = new SqlDataAdapter(command);

                SqlCommandBuilder builder = new SqlCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

                adapter.Fill(dt);

                foreach (DataRow nrow in dt.Rows)
                {
                    nrow["Personnel"] = ev.Value.ToString();
                    nrow["Resources"] = TaskModel.GenerateResourceIDsString(nrow["Machine"].ToString(), nrow["Personnel"].ToString(), schedulerStorage);
                }

                adapter.Update(dt);
            }
        }
        public static void UpdateTaskDates(List<TaskModel> tasks)
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string queryString = "UPDATE Tasks " +
                                     "SET StartDate = @StartDate, FinishDate = @FinishDate " +
                                     "WHERE ID = @ID";

                connection.Execute(queryString, tasks);
            }
        }
        public static void ForwardDateTasks(List<ComponentModel> selectedComponents, DateTime forwardDate)
        {
            if (forwardDate == new DateTime(2000, 1, 1))
            {
                return;
            }

            if (selectedComponents.Count == 0)
            {
                XtraMessageBox.Show("No Components were selected.");
                return;
            }

            ClearSelectedComponentDates(selectedComponents);

            List<TaskModel> tasksToUpdate = new List<TaskModel>();

            foreach (ComponentModel component in selectedComponents)
            {
                component.ForwardDate(forwardDate);

                tasksToUpdate.AddRange(component.Tasks);
            }

            UpdateTaskDates(tasksToUpdate);
        }
        private static void ClearSelectedComponentDates(List<ComponentModel> selectedComponents)
        {
            foreach (ComponentModel component in selectedComponents)
            {
                component.ClearTaskDates();
            }
        }
        public static void BackDateTasks(List<ComponentModel> selectedComponents, DateTime backDate)
        {
            List<TaskModel> tasksToUpdate = new List<TaskModel>();

            if (backDate == new DateTime(2000, 1, 1))
            {
                return;
            }

            if (selectedComponents.Count == 0)
            {
                XtraMessageBox.Show("No components selected.");
                return;
            }

            foreach (ComponentModel component in selectedComponents)
            {
                component.BackDate(backDate);

                tasksToUpdate.AddRange(component.Tasks);
            }

            UpdateTaskDates(tasksToUpdate);
        }

        #endregion

        #region Delete

        #endregion

        private static int GetPercentComplete(string status)
        {
            if (status == "Completed")
            {
                return 100;
            }
            else
            {
                return 0;
            }
        }

        public static DataTable GetDependencyData(DataTable taskTable)
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

                        row["DependentId"] = nrow["AptID"];
                        row["ParentId"] = Convert.ToInt32(predecessor);

                        dt.Rows.Add(row);

                        //Console.WriteLine($"{nrow["TaskID"]} {predecessor}");
                    }
                }
                else if (nrow["Predecessors"].ToString() != "")
                {
                    DataRow row = dt.NewRow();

                    row["DependentId"] = nrow["AptID"];
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

        #endregion // Task Operations

        #region Resources Table Operations

        #region Create

        public static void InsertResource(string resourceName, string resourceType)
        {
            if (resourceName != "")
            {
                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    SqlCommand cmd = new SqlCommand("INSERT INTO Resources (ResourceName, ResourceType) VALUES (@resourceName, @resourceType)", connection);

                    cmd.Parameters.AddWithValue("@resourceName", resourceName);
                    cmd.Parameters.AddWithValue("@resourceType", resourceType);

                    connection.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            else
            {
                MessageBox.Show("You have not entered a name for a resource to add.");
            }
        }

        #endregion

        #region Read

        public List<string> GetResourceList(string role) // This method gets a resource list 25% faster than using DataAdapter.Fill();
        {
            List<string> ResourceList = new List<string>();

            string queryString = "SELECT DISTINCT Resources.ResourceName From Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID WHERE Role = @role OR Role LIKE @role ORDER BY Resources.ResourceName ASC";

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                using (OleDbCommand cmd = new OleDbCommand(queryString, connection))
                {
                    Stopwatch sw = new Stopwatch();
                    sw.Start();

                    cmd.Parameters.AddWithValue("@role", "%" + role + "%");

                    connection.Open();

                    using (var rdr = cmd.ExecuteReader())
                    {
                        ResourceList.Add("");

                        while (rdr.Read())
                        {
                            ResourceList.Add($"{rdr["ResourceName"]}");
                            //Console.WriteLine($"{nrow["ResourceName"]}"); 
                        }
                    }

                    Console.WriteLine($"GetResourceList Transaction Time: {sw.Elapsed}");

                    return ResourceList;
                }
            }
        }

        public static DataTable GetRoleTable()
        {
            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                DataTable dt = new DataTable();

                string queryString = "SELECT * From Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID ORDER BY Resources.ResourceName ASC";

                SqlDataAdapter adapter = new SqlDataAdapter(queryString, connection);

                adapter.Fill(dt);

                return dt;
            }
        }

        public static List<string> GetResourceList()
        {
            DataTable dt = new DataTable();
            List<string> ResourceList = new List<string>();

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string queryString = "SELECT * From Resources ORDER BY ResourceName ASC";

                SqlCommand cmd = new SqlCommand(queryString, connection);

                connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            ResourceList.Add(rdr["ResourceName"].ToString());
                        }
                    }
                }
            }

            return ResourceList;
        }

        private static int GetResourceID(string resourceName)
        {
            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string queryString = "SELECT ID FROM Resources WHERE ResourceName = @resourceName";

                SqlCommand sqlCommand = new SqlCommand(queryString, connection);

                sqlCommand.Parameters.AddWithValue("@resourceName", resourceName);

                connection.Open();
                int resourceID = (int)sqlCommand.ExecuteScalar();

                return resourceID;
            }
        }

        public static List<string> GetRoleList(string role)
        {
            DataTable dt = new DataTable();
            List<string> RoleList = new List<string>();

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string queryString = "SELECT Resources.ResourceName, Roles.Role FROM Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID WHERE Roles.Role = @role ORDER BY Resources.ResourceName ASC";

                SqlCommand cmd = new SqlCommand(queryString, connection);

                cmd.Parameters.AddWithValue("@role", role);

                connection.Open();

                using (var rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        RoleList.Add(rdr["ResourceName"].ToString());
                    }
                }

                return RoleList;
            }
        }

        // This method is for populating the department schedule view with resources.
        public static DataTable GetResourceData()
        {
            DataTable dt = new DataTable();
            
            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string queryString = "SELECT ResourceName, ResourceType, Resources.ID, Role, Departments.Department From (Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID) LEFT OUTER JOIN Departments ON Roles.DepartmentID = Departments.ID ORDER BY ResourceName ASC";

                SqlDataAdapter adapter = new SqlDataAdapter(queryString, connection);

                adapter.Fill(dt);

                DataRow row1 = dt.NewRow();

                row1["ResourceName"] = "No Machine";
                row1["Role"] = "None";
                row1["Department"] = "None";
                row1["ResourceType"] = "Machine";

                dt.Rows.Add(row1);

                DataRow row2 = dt.NewRow();

                row2["ResourceName"] = "No Personnel";
                row2["Role"] = "None";
                row2["Department"] = "None";
                row2["ResourceType"] = "Personnel";

                dt.Rows.Add(row2);

                DataRow row3 = dt.NewRow();

                row3["ResourceName"] = "None";
                row3["Role"] = "None";
                row3["Department"] = "None";
                row3["ResourceType"] = "None";

                dt.Rows.Add(row3);

                //foreach (DataRow nrow in dt.Rows)
                //{
                //    Console.WriteLine($"{nrow["ID"]} {nrow["ResourceName"]} {nrow["Role"]} {nrow["Department"]}");
                //}

                return dt; 
            }
        }

        public static string GetResourceType(string resource)
        {
            string queryString = "SELECT ResourceType FROM Resources WHERE ResourceName = @resourceName";

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                SqlCommand cmd = new SqlCommand(queryString, connection);

                cmd.Parameters.AddWithValue("@resourceName", resource);

                connection.Open();

                return cmd.ExecuteScalar().ToString();
            }
        }

        public static List<string> GetAllResourcesOfType(string resourceType)
        {
            string queryString = "SELECT ResourceName FROM Resources WHERE ResourceType = @resourceType ORDER BY ResourceName";

            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                var output = connection.Query<string>(queryString, new { ResourceType = resourceType }).ToList();

                return output;
            }
        }

        #endregion

        #region Update

        public static void SetResourceType(string resourceName, string resourceType)
        {
            string queryString = "UPDATE Resources SET ResourceType = @resourceType WHERE ResourceName = @resourceName";

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                SqlCommand cmd = new SqlCommand(queryString, connection);

                cmd.Parameters.AddWithValue("@resourceType", resourceType);
                cmd.Parameters.AddWithValue("@resourceName", resourceName);

                connection.Open();
                cmd.ExecuteNonQuery();
            }
        }

        #endregion

        #region Delete

        public static void RemoveResource(string resourceName)
        {

            if (resourceName != "")
            {
                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    SqlCommand cmd1 = new SqlCommand("DELETE FROM Roles WHERE ResourceID = @resourceID ", connection);
                    SqlCommand cmd2 = new SqlCommand("DELETE FROM Resources WHERE ID = @resourceID ", connection);

                    int resourceID = GetResourceID(resourceName);

                    cmd1.Parameters.AddWithValue("@resourceID", resourceID);
                    cmd2.Parameters.AddWithValue("@resourceID", resourceID);

                    connection.Open();
                    cmd1.ExecuteNonQuery();
                    cmd2.ExecuteNonQuery();
                }
            }
            else
            {
                MessageBox.Show("You have not selected a resource to remove.");
            }
        }

        #endregion

        #endregion // Resources

        #region Roles Table Operations

        #region Create

        public static void InsertResourceRole(string resourceName, string role, int departmentID)
        {
            if (resourceName != "")
            {
                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    SqlCommand cmd = new SqlCommand("INSERT INTO Roles (ResourceID, Role, DepartmentID) VALUES (@resourceID, @role, @departmentID)", connection);

                    cmd.Parameters.AddWithValue("@resourceID", GetResourceID(resourceName));
                    cmd.Parameters.AddWithValue("@role", role);
                    cmd.Parameters.AddWithValue("@departmentID", departmentID);

                    connection.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            else
            {
                MessageBox.Show("You have not selected a resource to add a role to.");
            }
        }

        #endregion

        #region Read

        // Get ResourceList joins the resources and roles tables under the Read region under the Resources Table region.

        public DataTable GetRoleCounts()
        {
            DataTable dt = new DataTable();

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string queryString = "SELECT COUNT(*) AS RoleCount, Role FROM Roles GROUP BY Role";

                SqlDataAdapter adapter = new SqlDataAdapter(queryString, connection);

                adapter.Fill(dt);

                return dt; 
            }
        }

        #endregion

        #region Update

        #endregion

        #region Delete

        public static void RemoveResourceRole(string resourceName, string role)
        {
            if (resourceName != "")
            {
                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    SqlCommand cmd = new SqlCommand("DELETE FROM Roles WHERE ResourceID = @resourceID AND Role = @role", connection);

                    cmd.Parameters.AddWithValue("resourceID", GetResourceID(resourceName));
                    cmd.Parameters.AddWithValue("@role", role);

                    connection.Open();
                    cmd.ExecuteNonQuery(); 
                }
            }
            else
            {
                MessageBox.Show("You have not selected a resource to add a role to.");
            }
        }

        #endregion

        #endregion

        #region Departments Table Operations

        #region Read

        public static List<DepartmentModel> LoadDepartments()
        {
            string queryString = "SELECT * FROM Departments";

            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                var output = connection.Query<DepartmentModel>(queryString, new DynamicParameters()).ToList();

                return output;
            }
        }

        #endregion

        #endregion

        #region Workload Table Operations

        #region Create

        public static void CreateWorkloadEntry(WorkLoadModel workload)
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string queryString = "INSERT INTO Workload (ToolNumber, MWONumber, ProjectNumber, Stage, Customer, Project, DeliveryInWeeks, StartDate, FinishDate, AdjustedDeliveryDate, MoldCost, Engineer, Designer, ToolMaker, RoughProgrammer, FinishProgrammer, ElectrodeProgrammer, Apprentice, Manifold, MoldBase, GeneralNotes) VALUES " +
                                                          "(@toolNumber, @mwoNumber, @projectNumber, @stage, @customer, @project, @deliveryInWeeks, @startDate, @finishDate, @adjustedDeliveryDate, @moldCost, @engineer, @designer, @toolMaker, @roughProgrammer, @finishProgrammer, @electrodeProgrammer, @apprentice, @manifold, @moldBase, @generalNotes)";

                connection.Execute(queryString, workload);
            }
        }

        #endregion

        #region Read

        public static List<WorkLoadModel> GetWorkloads()
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                List<WorkLoadModel> workLoads = connection.Query<WorkLoadModel>("SELECT * FROM Workload").ToList();

                return workLoads;
            }
        }

        #endregion

        #region Update

        public static bool UpdateWorkloadColumn(WorkLoadModel workload, CellValueChangedEventArgs ev)
        {
            if (workload.ID.ToString() == "")
            {
                // I'm tricking the system here.
                return true;
            }

            string queryString;

            //queryString = "dbo.spUpdateWorkload @ToolNumber, @MWONumber, @ProjectNumber, @Stage, @Customer, @PartName, @Engineer, @DeliveryInWeeks, @StartDate, @FinishDate, @AdjustedDeliveryDate," +
            //    "@MoldCost, @Designer, @ToolMaker, @RoughProgrammer, @FinishProgrammer, @ElectrodeProgrammer, @Apprentice, @Manifold, @Moldbase, @GeneralNotes, @JobFolderPath, @ID";

            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                queryString = $"UPDATE Workload SET {ev.Column.FieldName} = @{ev.Column.FieldName} WHERE ID = @ID";

                connection.Execute(queryString, workload);

                return true;
            }
        }

        #endregion

        #region Delete

        public static bool DeleteWorkLoadEntry(int id)
        {
            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                connection.Execute("DELETE FROM WorkLoad WHERE ID = @id", new { ID = id });
            }

            return true;
        }

        #endregion

        #endregion // WorkLoad

        #region WorkloadColors Table Operations

        #region Create

        public static void AddColorEntry(int projectID, string column, int aRGBColor)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    SqlCommand cmd = new SqlCommand("INSERT INTO WorkLoadColors (ProjectID, ColumnFieldName, ARGBColor) VALUES (@projectID, @columnFieldName, @aRGBColor)", connection);

                    cmd.Parameters.AddWithValue("@projectID", projectID);
                    cmd.Parameters.AddWithValue("@columnFieldName", column);
                    cmd.Parameters.AddWithValue("@aRGBColor", aRGBColor);

                    connection.Open();
                    cmd.ExecuteNonQuery(); 
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
            }
        }

        #endregion

        #region Read

        public static List<ColorStruct> GetColorEntries()
        {
            List<ColorStruct> colorList = new List<ColorStruct>();

            try
            {
                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    SqlCommand cmd = new SqlCommand("SELECT * FROM WorkLoadColors", connection);

                    connection.Open();

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

                    return colorList; 
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
                return null;
            }
        }

        #endregion

        #region Update

        public static void UpdateColorEntry(int projectID, string column, int aRGBColor)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = connection;

                    cmd.CommandText = "UPDATE WorkLoadColors SET ARGBColor = @aRGBColor WHERE (ProjectID = @projectID AND ColumnFieldName = @column)";

                    cmd.Parameters.AddWithValue("@aRGBColor", aRGBColor);
                    cmd.Parameters.AddWithValue("@projectID", projectID);
                    cmd.Parameters.AddWithValue("@column", column);

                    connection.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
            }
        }

        #endregion

        #region Delete

        public static void DeleteColorEntries(int projectID)
        {
            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                connection.Execute("DELETE FROM WorkLoadColors WHERE ProjectID = @projectID", new { ProjectID = projectID });
            }
        }

        #endregion

        #endregion //WorkLoadColors

        #region Machines Table Operations

        #region Create

        // TODO: Create method for adding machines to the database.

        #endregion

        #region Read

        public List<string> GetMachineList(string machineType)
        {
            List<string> machineList = new List<string>();
            DataTable dt = new DataTable();

            string queryString = "SELECT MachineName From Machines WHERE MachineType LIKE @machineType ORDER BY MachineName ASC";

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

                adapter.SelectCommand.Parameters.AddWithValue("@machineType", "%" + machineType + "%");

                adapter.Fill(dt);

                foreach (DataRow nrow in dt.Rows)
                {
                    machineList.Add($"{nrow["MachineName"]}");
                } 
            }

            return machineList;
        }

        #endregion

        #region Update

        public void SetDailyDepartmentCapacities(string department)
        {
            DataTable dt = GetRoleCounts();
        }

        #endregion

        #region Delete

        // TODO: Create method for deleting machines from the database.

        #endregion

        #endregion

        #region Departments Table Operations

        #region Create
        
        // TODO: Create method for adding departments to the database.

        #endregion

        #region Read

        public static DataTable GetDailyDepartmentCapacities()
        {
            string queryString = "SELECT * FROM Departments";

            DataTable dt = new DataTable();

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                SqlDataAdapter adapter = new SqlDataAdapter(queryString, connection);

                adapter.Fill(dt);

                return dt; 
            }
        }

        #endregion

        #region Update

        #endregion

        #region Delete

        // TODO: Create method for deleting departments from the database.

        #endregion

        #endregion // Departments Table Operations
        
        public static double GetBusinessDays(DateTime startD, DateTime endD)
        {
            double calcBusinessDays =
                1 + ((endD - startD).TotalDays * 5 -
                (startD.DayOfWeek - endD.DayOfWeek) * 2) / 7;

            if (endD.DayOfWeek == DayOfWeek.Saturday) calcBusinessDays--;
            if (startD.DayOfWeek == DayOfWeek.Sunday) calcBusinessDays--;

            return calcBusinessDays;
        }

        // Creates a weeklist with 20 weeks for each department.
        public static List<Week> InitializeDeptWeeksList(DateTime wsDate, List<string> departmentArr)
        {
            List<Week> weekList = new List<Week>();

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
    }
}
