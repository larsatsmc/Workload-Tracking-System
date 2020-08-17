using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics;
using DevExpress.Spreadsheet;
using ClassLibrary;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.Xml;
using Dapper;
using System.Data.SqlClient;
using System.Threading.Tasks;
using DevExpress.XtraSpreadsheet.PrintLayoutEngine;

namespace Toolroom_Project_Viewer
{
    class Database
    {
        // WorkloadTrackingSystemDB, LocalSqlServerDB See App.config for list of connection names.
        static readonly string DatabaseType = "SQL Server"; // Either 'Access' or 'SQL Server'.
        static readonly string SQLClientConnectionName = "LocalSqlServerDB";
        static readonly string OLEDBConnectionName = "LocalOLEDBSqlServerDB";

        private static string UpdateProjectString = "dbo.spUpdateProject @JobNumber, @ProjectNumber, @Customer, @Project, @DueDate, @Status, @PercentComplete, @Designer, @ToolMaker, @RoughProgrammer, @ElectrodeProgrammer, @FinishProgrammer, @Apprentice, @EDMSinkerOperator, @RoughCNCOperator, @ElectrodeCNCOperator, @FinishCNCOperator, @EDMWireOperator, @OverlapAllowed, @IncludeHours, @KanBanWorkbookPath, @ID";

        OleDbConnection Connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName));

        DataTable TaskIDKey = new DataTable();

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
                    string queryString1 = "INSERT INTO Projects (JobNumber, ProjectNumber, Customer, Project, DueDate, Priority, Designer, ToolMaker, RoughProgrammer, ElectrodeProgrammer, FinishProgrammer, EDMSinkerOperator, RoughCNCOperator, ElectrodeCNCOperator, FinishCNCOperator, EDMWireOperator, Apprentice, OverlapAllowed, DateCreated, DateModified) " + // 
                                          "VALUES (@JobNumber, @ProjectNumber, @Customer, @Project, @DueDate, @Priority, @Designer, @ToolMaker, @RoughProgrammer, @ElectrodeProgrammer, @FinishProgrammer, @EDMSinkerOperator, @RoughCNCOperator, @ElectrodeCNCOperator, @FinishCNCOperator, @EDMWireOperator, @Apprentice, @OverlapAllowed, GETDATE(), GETDATE())"; // 

                    string queryString2 = "INSERT INTO Components (JobNumber, ProjectNumber, Component, Notes, Priority, [Position], Material, TaskIDCount, Quantity, Spares, Pictures, Finish) " + // 
                                          "VALUES (@JobNumber, @ProjectNumber, @Component, @Notes, @Priority, @Position, @Material, @TaskIDCount, @Quantity, @Spares, @Pictures, @Finish)"; // 

                    string queryString3 = "INSERT INTO Tasks (JobNumber, ProjectNumber, Component, TaskID, TaskName, Duration, StartDate, FinishDate, Predecessors, Machine, Resources, Resource, Hours, Priority, DateAdded, Notes) " +
                                          "VALUES (@JobNumber, @ProjectNumber, @Component, @TaskID, @TaskName, @Duration, @StartDate, @FinishDate, @Predecessors, @Machine, @Resources, @Resource, @Hours, @Priority, Now(), @Notes)";

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
                            project.OverlapAllowed
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
                                component.Pictures,
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
                                    task.Resource,
                                    task.Hours,
                                    task.Priority,
                                    task.Notes
                                };

                                connection.Execute(queryString3, p3, trans);

                                task.ID = (int)connection.ExecuteScalar("SELECT @@IDENTITY", transaction: trans);
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
            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                List<ProjectModel> projects = connection.Query<ProjectModel>("SELECT * FROM Projects").ToList();
                List<ComponentModel> components = connection.Query<ComponentModel>("SELECT * FROM Components").ToList();
                List<TaskModel> tasks = connection.Query<TaskModel>("SELECT * FROM Tasks").ToList();

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
                int projectCount = connection.Execute("SELECT COUNT(*) from Projects WHERE ProjectNumber = @ProjectNumber", new { ProjectNumber = projectNumber } );
                
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

        //public ProjectModel GetProject(string jobNumber, int projectNumber)
        //{
        //    // TODO: Need to make sure all these methods use the same database connection.
        //    ProjectModel project = GetProjectInfo(jobNumber, projectNumber);

        //    GetComponents(project);

        //    GetTasks(project);

        //    return project;
        //}

        public static ProjectModel GetProject4(int projectNumber)
        {
            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                var p = new { ProjectNumber = projectNumber };

                ProjectModel project = connection.QueryFirst<ProjectModel>("SELECT * FROM Projects WHERE ProjectNumber = @ProjectNumber", p);
                List<ComponentModel> components = connection.Query<ComponentModel>("SELECT * FROM Components WHERE ProjectNumber = @ProjectNumber", p).ToList();

                string taskQuery = "SELECT ID, TaskID, Component, TaskName, Duration, StartDate, FinishDate, Predecessors, Machine, Resources, Resource AS Personnel, Hours, Notes, DateCompleted, Initials, Status FROM Tasks WHERE ProjectNumber = @ProjectNumber";

                List<TaskModel> tasks = connection.Query<TaskModel>(taskQuery, p).ToList();

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

        public static ProjectModel GetProjectSQL(int projectNumber)
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                var p = new { ProjectNumber = projectNumber };

                ProjectModel project = connection.QueryFirst<ProjectModel>("dbo.spGetProject @ProjectNumber", p);
                List<ComponentModel> components = connection.Query<ComponentModel>("dbo.spGetProjectComponents @ProjectNumber", p).ToList();

                string taskQuery = "SELECT ID, TaskID, Component, TaskName, Duration, StartDate, FinishDate, Predecessors, Machine, Resources, Resource AS Personnel, Hours, Notes, DateCompleted, Initials, Status FROM Tasks WHERE ProjectNumber = @ProjectNumber";

                List<TaskModel> tasks = connection.Query<TaskModel>(taskQuery, p).ToList();

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
            OleDbCommand cmd;
            ProjectModel pi;
            List<ProjectModel> piList = new List<ProjectModel>();

            try
            {
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    cmd = new OleDbCommand(queryString, connection);

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
            int baseCount = 0;

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
                    
                    row["PercentComplete"] = GetPercentComplete(task.Status);
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

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                OleDbCommand sqlCommand = new OleDbCommand("SELECT KanBanWorkbookPath from Projects WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", connection);

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

        //public bool EditProjectInDB(ProjectModel project)
        //{
        //    // TODO: Need to make sure that all these methods are using the same database connection.
        //    try
        //    {
        //        ProjectModel databaseProject = GetProjectSQL(project.OldProjectNumber);
        //        List<ComponentModel> newComponentList = new List<ComponentModel>();
        //        //List<Component> updatedComponentList = new List<Component>();
        //        List<TaskModel> newTaskList = new List<TaskModel>();
        //        List<ComponentModel> deletedComponentList = new List<ComponentModel>();

        //        if (project.ProjectNumberChanged && ProjectExists(project.ProjectNumber))
        //        {
        //            MessageBox.Show("There is another project with that same project number. Enter a different project number.");
        //            return false;
        //        }

        //        UpdateProjectData(project);

        //        if (ProjectHasComponents(project.ProjectNumber))
        //        {
        //            // Check modified project for added components.
        //            foreach (ComponentModel component in project.Components)
        //            {
        //                component.SetPosition(project.Components.IndexOf(component));

        //                bool componentExists = databaseProject.Components.Exists(x => x.Component == component.Component);

        //                if (componentExists)
        //                {
        //                    UpdateComponentData(project, component);

        //                    foreach (TaskModel task in component.Tasks)
        //                    {
        //                        Console.WriteLine($"{task.TaskID} {task.TaskName}");
        //                    }

        //                    UpdateTasks(project.JobNumber, project.ProjectNumber, component.Component, component.Tasks);

        //                    //if (component.ReloadTaskList)
        //                    //{
        //                    //    RemoveTasks(project, component);
        //                    //    newTaskList.AddRange(component.TaskList);
        //                    //}
        //                    //else
        //                    //{
        //                    //    UpdateTasks(project.JobNumber, project.ProjectNumber, component.Name, component.TaskList);
        //                    //}
        //                }
        //                else
        //                {
        //                    newComponentList.Add(component);
        //                    newTaskList.AddRange(component.Tasks);
        //                }
        //            }

        //            // Check modified project for deleted components.
        //            foreach (ComponentModel component in databaseProject.Components)
        //            {
        //                bool componentExists = project.Components.Exists(x => x.Component == component.Component);

        //                if (!componentExists)
        //                {
        //                    deletedComponentList.Add(component);
        //                }
        //            }

        //            // Check modified project for updated tasks

        //            if (newComponentList.Count > 0)
        //            {
        //                AddComponentDataTableToDatabase(CreateDataTableFromComponentList(project, newComponentList));
        //            }

        //            if (newTaskList.Count > 0)
        //            {
        //                AddTaskDataTableToDatabase(CreateDataTableFromTaskList(project, newTaskList));
        //            }

        //            if (deletedComponentList.Count > 0)
        //            {
        //                foreach (ComponentModel component in deletedComponentList)
        //                {
        //                    RemoveComponent(project, component);
        //                }
        //            }
        //        }
        //        else
        //        {
        //            AddComponentDataTableToDatabase(CreateDataTableFromComponentList(project));
        //        }

        //        MessageBox.Show("Project Updated!");

        //        return true;
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
        //        return false;
        //    }
        //}

        //private void UpdateProjectData(ProjectModel project)
        //{
        //    try
        //    {
        //        OleDbDataAdapter adapter = new OleDbDataAdapter();

        //        string queryString;

        //        queryString = "UPDATE Projects " +
        //                      "SET JobNumber = @jobNumber, ProjectNumber = @newProjectNumber, DueDate = @dueDate, Designer = @designer, ToolMaker = @toolMaker, RoughProgrammer = @roughProgrammer, ElectrodeProgrammer = @electrodeProgrammer, " +
        //                      "FinishProgrammer = @finishProgrammer, Apprentice = @apprentice, EDMSinkerOperator = @EDMSinkerOperator, RoughCNCOperator = @RoughCNCOperator, ElectrodeCNCOperator = @ElectrodeCNCOperator, FinishCNCOperator = @FinishCNCOperator, EDMWireOperator = @EDMWireOperator, DateModified = Now() " +
        //                      "WHERE ProjectNumber = @oldProjectNumber";

        //        adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

        //        adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", project.JobNumber);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@newProjectNumber", project.ProjectNumber);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@dueDate", project.DueDate);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@designer", project.Designer);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@toolMaker", project.ToolMaker);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@roughProgrammer", project.RoughProgrammer);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@electrodeProgrammer", project.ElectrodeProgrammer);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@finishProgrammer", project.FinishProgrammer);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@EDMSinkerOperator", project.EDMSinkerOperator);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@RoughCNCOperator", project.RoughCNCOperator);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@ElectrodeCNCOperator", project.ElectrodeCNCOperator);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@FinishCNCOperator", project.FinishCNCOperator);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@EDMWireOperator", project.EDMWireOperator);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@apprentice", project.Apprentice);
        //        adapter.UpdateCommand.Parameters.AddWithValue("@oldProjectNumber", project.OldProjectNumber);  // By default this number is set to whatever is in the database when it was loaded to the Edit project form.

        //        Connection.Open();

        //        adapter.UpdateCommand.ExecuteNonQuery();
        //    }
        //    catch (OleDbException ex)
        //    {
        //        throw ex;
        //    }
        //    catch (Exception x)
        //    {
        //        throw x;
        //    }
        //    finally
        //    {
        //        Connection.Close();
        //    }
        //}

        public static bool EditProject(ProjectModel project)
        {
            if (project.ProjectNumberChanged && ProjectExists(project.ProjectNumber))
            {
                MessageBox.Show("There is another project with that same project number. Enter a different project number.");
                return false;
            }

            ProjectModel databaseProject = GetProjectSQL(project.OldProjectNumber);

            using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                List<ComponentModel> newComponentList = new List<ComponentModel>();
                //List<Component> updatedComponentList = new List<Component>();
                List<TaskModel> taskList = new List<TaskModel>();
                List<TaskModel> databaseTaskList = new List<TaskModel>();
                List<ComponentModel> deletedComponentList = new List<ComponentModel>();

                connection.Execute(UpdateProjectString, project);

                var componentsToAdd = from component in project.Components
                                      where component.ID == 0
                                      select component;

                var componentsToUpdate = from component in project.Components
                                            where component.ID != 0
                                            select component;

                var componentsToRemove = from component in databaseProject.Components
                                         where !project.Components.Exists(x => x.ID == component.ID)
                                         select component;

                connection.Execute("dbo.spCreateComponent @JobNumber, @ProjectNumber, @Component, @Notes, @Priority, @Position, @Material, @TaskIDCount, @Quantity, @Spares, @Pictures, @Finish, @Status, @PercentComplete", componentsToAdd.ToList());
                connection.Execute("dbo.spUpdateComponent @Component, @Notes, @Priority, @Position, @Quantity, @Spares, @Pictures, @Material, @Finish, @TaskIDCount, @ID", componentsToUpdate.ToList());
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
                connection.Execute("dbo.spUpdateTask @TaskID, @TaskName, @Hours, @Duration, @Machine, @Resources, @Resource, @Predecessors, @Priority, @Notes, @ID", tasksToUpdate.ToList());
                connection.Execute("DELETE FROM Tasks WHERE ID = @ID", tasksToDelete.ToList());
            }

            return true;
        }
        public static bool EditProjectRecord(ProjectModel project, CellValueChangedEventArgs ev)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    if (ev.Column.FieldName == "ProjectNumber" && connection.Execute("dbo.spGetProjectCount @ProjectNumber", project) > 0)
                    {
                        MessageBox.Show("There is a project with that same project number.");
                        goto MyEnd;
                    }

                    connection.Execute(UpdateProjectString, project);

                    return true;

                MyEnd:;
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                throw e;
            }
        }
        public static bool UpdateProjectsTable(object s, CellValueChangedEventArgs ev)
        {
            try
            {
                var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);
                string whereClause;

                if (grid.Columns["MWONumber"] == null)
                {
                    whereClause = "WHERE (ID = @tID)";
                }
                else
                {
                    whereClause = "WHERE (ProjectNumber = @projectNumber)";
                }

                //queryString = "UPDATE Tasks SET JobNumber = @jobNumber, Component = @component, TaskID = @taskID, TaskName = @taskName, " +
                //    "Duration = @duration, StartDate = @startDate, FinishDate = @finishDate, Predecessor = @predecessor, Machines = @machines, " +
                //    "Machine = @machine, Person = @person, Priority = @priority WHERE ID = @tID";

                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    var cmd = new SqlCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = connection;

                    if (ev.Column.FieldName == "JobNumber")
                    {
                        cmd.CommandText = "UPDATE Projects SET JobNumber = @jobNumber, DateModified = Now() " + whereClause;

                        cmd.Parameters.AddWithValue("@jobNumber", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "ProjectNumber")
                    {
                        if (!ProjectExists((int)ev.Value))
                        {
                            cmd.CommandText = "UPDATE Projects SET ProjectNumber = @projectNumber, DateModified = Now() " + whereClause;

                            if (ev.Value.ToString() != "")
                            {
                                cmd.Parameters.AddWithValue("@projectNumber", ev.Value.ToString());
                            }
                            else
                            {
                                cmd.Parameters.AddWithValue("@projectNumber", "");
                            }
                        }
                        else
                        {
                            MessageBox.Show("There is a project with that same project number.");
                            goto MyEnd;
                        }
                    }
                    else if (ev.Column.FieldName == "Customer")
                    {
                        cmd.CommandText = "UPDATE Projects SET Customer = @customer " + whereClause;

                        cmd.Parameters.AddWithValue("@customer", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Project")
                    {
                        cmd.CommandText = "UPDATE Projects SET Project = @project " + whereClause;

                        cmd.Parameters.AddWithValue("@project", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "DueDate")
                    {
                        cmd.CommandText = "UPDATE Projects SET DueDate = @dueDate " + whereClause;

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
                        cmd.CommandText = "UPDATE Projects SET Status = @status " + whereClause;

                        cmd.Parameters.AddWithValue("@status", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Designer")
                    {
                        cmd.CommandText = "UPDATE Projects SET Designer = @designer " + whereClause;

                        cmd.Parameters.AddWithValue("@designer", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "ToolMaker")
                    {
                        cmd.CommandText = "UPDATE Projects SET ToolMaker = @toolMaker " + whereClause;

                        cmd.Parameters.AddWithValue("@toolMaker", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "RoughProgrammer")
                    {
                        cmd.CommandText = "UPDATE Projects SET RoughProgrammer = @roughProgrammer " + whereClause;

                        cmd.Parameters.AddWithValue("@roughProgrammer", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "FinishProgrammer")
                    {
                        cmd.CommandText = "UPDATE Projects SET FinishProgrammer = @finishProgrammer " + whereClause;

                        cmd.Parameters.AddWithValue("@finishProgrammer", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "ElectrodeProgrammer")
                    {
                        cmd.CommandText = "UPDATE Projects SET ElectrodeProgrammer = @electrodeProgrammer " + whereClause;

                        cmd.Parameters.AddWithValue("@electrodeProgrammer", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Apprentice")
                    {
                        cmd.CommandText = "UPDATE Projects SET Apprentice = @apprentice " + whereClause;

                        cmd.Parameters.AddWithValue("@apprentice", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "KanBanWorkbookPath")
                    {
                        cmd.CommandText = "UPDATE Projects SET KanBanWorkbookPath = @kanBanWorkbookPath " + whereClause;

                        cmd.Parameters.AddWithValue("@kanBanWorkbookPath", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "OverlapAllowed")
                    {
                        cmd.CommandText = "UPDATE Projects SET OverlapAllowed = @overlapAllowed " + whereClause;

                        cmd.Parameters.AddWithValue("@overlapAllowed", (bool)ev.Value);
                    }
                    else if (ev.Column.FieldName == "IncludeHours")
                    {
                        cmd.CommandText = "UPDATE Projects SET IncludeHours = @includeHours " + whereClause;

                        cmd.Parameters.AddWithValue("@includeHours", (bool)ev.Value);
                    }

                    if (grid.Columns["MWONumber"] == null)
                    {
                        cmd.Parameters.AddWithValue("@tID", grid.GetRowCellValue(ev.RowHandle, grid.Columns["ID"]));
                    }
                    else
                    {
                        if (grid.Columns["MWONumber"].ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@projectNumber", grid.GetRowCellValue(ev.RowHandle, grid.Columns["MWONumber"]));
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@projectNumber", grid.GetRowCellValue(ev.RowHandle, grid.Columns["ProjectNumber"]));
                        }
                        
                    }
                    
                    connection.Open();
                    cmd.ExecuteNonQuery();
                    return true;
                    MyEnd:;
                    return false;
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
        public bool RemoveProject(string jobNumber, int projectNumber)
        {
            using (OleDbConnection connection = new OleDbConnection(OLEDBConnectionName))
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
        private bool AddComponentDataTableToDatabase(DataTable dt)
        {
            var adapter = new OleDbDataAdapter();

            //Console.WriteLine(cbr.GetInsertCommand().CommandText);

            try
            {
                adapter.SelectCommand = new OleDbCommand("SELECT * FROM Components", Connection);

                var cbr = new OleDbCommandBuilder(adapter);

                cbr.QuotePrefix = "[";
                cbr.QuoteSuffix = "]";
                cbr.GetDeleteCommand();
                cbr.GetInsertCommand();
                adapter.UpdateCommand = cbr.GetUpdateCommand();

                Connection.Open();
                adapter.Update(dt);
                Connection.Close();
                Console.WriteLine("Components Loaded."); 
                
            }
            catch (OleDbException ex)
            {
                Connection.Close();
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace, "OledbException Error");
                return false;
            }
            catch (Exception x)
            {
                Connection.Close();
                MessageBox.Show(x.Message + "\n\n" + x.StackTrace, "Exception Error");
                return false;
            }

            return true;
        }

        #endregion

        #region Read
        public static List<ComponentModel> GetAllComponents()
        {
            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                List<ComponentModel> results = connection.Query<ComponentModel>("SELECT * FROM Components").ToList();

                return results;
            }
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

        public void GetComponents(ProjectModel project)
        {
            
            ComponentModel component;

            string queryString = "SELECT * FROM Components WHERE ProjectNumber = @projectNumber";

            OleDbCommand cmd = new OleDbCommand(queryString, Connection);

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
                                      component: rdr["Component"],
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
                        project.AddComponentList(GetComponentListFromTasksTable(project.ProjectNumber));
                    }
                }
            }
            finally
            {
                Connection.Close();
            }
        }

        public List<ComponentModel> GetComponentListFromTasksTable(int projectNumber)
        {
            OleDbCommand cmd;
            List<ComponentModel> componentList = new List<ComponentModel>();

            string queryString;

            queryString = "SELECT DISTINCT Component FROM Tasks WHERE ProjectNumber = @projectNumber";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    cmd = new OleDbCommand(queryString, connection);
                    cmd.Parameters.AddWithValue("@projectNumber", projectNumber);

                    connection.Open();

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
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "getComponentListFromTaskTable");
            }

            return componentList;
        }

        #endregion

        #region Update

        private void UpdateComponentData(ProjectModel project, ComponentModel component)
        {
            try
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                string queryString;

                queryString = "UPDATE Components " +
                        "SET Component = @name, Notes = @notes, Priority = @priority, [Position] = @position, Quantity = @quantity, Spares = @spares, Pictures = @picture, Material = @material, Finish = @finish, TaskIDCount = @taskIDCount " +
                        "WHERE ID = @ID"; // JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @oldName

                adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

                adapter.UpdateCommand.Parameters.AddWithValue("@name", component.Component);
                adapter.UpdateCommand.Parameters.AddWithValue("@notes", component.Notes);
                adapter.UpdateCommand.Parameters.AddWithValue("@priority", component.Priority);
                adapter.UpdateCommand.Parameters.AddWithValue("@position", component.Position);
                adapter.UpdateCommand.Parameters.AddWithValue("@quantity", component.Quantity);
                adapter.UpdateCommand.Parameters.AddWithValue("@spares", component.Spares);

                if (component.Pictures != null)
                {
                    adapter.UpdateCommand.Parameters.AddWithValue("@picture", component.Pictures);
                }
                else
                {
                    adapter.UpdateCommand.Parameters.AddWithValue("@picture", DBNull.Value);
                }

                //adapter.UpdateCommand.Parameters.AddWithValue("@pictures", component.PictureList);  // Add when database is ready to receive pictures.

                adapter.UpdateCommand.Parameters.AddWithValue("@material", component.Material);
                adapter.UpdateCommand.Parameters.AddWithValue("@finish", component.Finish);
                adapter.UpdateCommand.Parameters.AddWithValue("@taskIDCount", component.TaskIDCount);

                //adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", project.JobNumber);
                //adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", project.ProjectNumber);
                //adapter.UpdateCommand.Parameters.AddWithValue("@oldName", component.OldName);
                adapter.UpdateCommand.Parameters.AddWithValue("@ID", component.ID);

                Connection.Open();

                adapter.UpdateCommand.ExecuteNonQuery();

                Console.WriteLine($"{component.Component} Updated.");
                //MessageBox.Show("Project Updated!"); 
                
            }
            //catch (OleDbException ex)
            //{
            //    throw ex;
            //}
            //catch (Exception x)
            //{
            //    throw x;
            //}
            finally
            {
                Connection.Close();
            }
        }

        #endregion

        #region Delete

        private void RemoveComponent(ProjectModel project, ComponentModel component)
        {
            var adapter = new OleDbDataAdapter();

            adapter.DeleteCommand = new OleDbCommand("DELETE FROM Components WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component", Connection);

            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = project.JobNumber;
            adapter.DeleteCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = project.ProjectNumber;
            adapter.DeleteCommand.Parameters.Add("@component", OleDbType.VarChar, ComponentModel.ComponentCharacterLimit).Value = component.Component;

            Connection.Open();
            adapter.DeleteCommand.ExecuteNonQuery();
            Connection.Close();
        }

        public static void UpdateComponentsTable(object s, CellValueChangedEventArgs ev)
        {
            try
            {
                var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

                //queryString = "UPDATE Tasks SET JobNumber = @jobNumber, Component = @component, TaskID = @taskID, TaskName = @taskName, " +
                //    "Duration = @duration, StartDate = @startDate, FinishDate = @finishDate, Predecessor = @predecessor, Machines = @machines, " +
                //    "Machine = @machine, Person = @person, Priority = @priority WHERE ID = @tID";

                using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
                {
                    var cmd = new SqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandType = CommandType.Text;

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

                    connection.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion

        private DataTable CreateDataTableFromComponentList(ProjectModel project)
        {
            DataTable dt = new DataTable();
            int position = 0;

            // These three lines add the necessary columns to the datatable without adding data.

            string queryString = "SELECT * FROM Components WHERE ID = 0";

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            //dt.Columns.Add("JobNumber", typeof(string));
            //dt.Columns.Add("ProjectNumber", typeof(int));
            //dt.Columns.Add("Component", typeof(string));
            //dt.Columns.Add("Notes", typeof(string));
            //dt.Columns.Add("Position", typeof(int));
            //dt.Columns.Add("Priority", typeof(int));
            //dt.Columns.Add("Pictures", typeof(byte[]));
            //dt.Columns.Add("Material", typeof(string));
            //dt.Columns.Add("Finish", typeof(string));
            //dt.Columns.Add("TaskIDCount", typeof(int));
            //dt.Columns.Add("Quantity", typeof(int));
            //dt.Columns.Add("Spares", typeof(int));
            //dt.Columns.Add("Status", typeof(string));
            //dt.Columns.Add("PercentComplete", typeof(int));

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
                if (component.Pictures != null)
                {
                    row["Pictures"] = component.Pictures;
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

        private DataTable CreateDataTableFromComponentList(ProjectModel project, List<ComponentModel> componentList)
        {
            DataTable dt = new DataTable();

            // These three lines add the necessary columns to the datatable without adding data.

            string queryString = "SELECT * FROM Components WHERE ID = 0";

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            //dt.Columns.Add("JobNumber", typeof(string));
            //dt.Columns.Add("ProjectNumber", typeof(int));
            //dt.Columns.Add("Component", typeof(string));
            //dt.Columns.Add("Notes", typeof(string));
            //dt.Columns.Add("Position", typeof(int));
            //dt.Columns.Add("Priority", typeof(int));
            //dt.Columns.Add("Pictures", typeof(byte[]));
            //dt.Columns.Add("Material", typeof(string));
            //dt.Columns.Add("Finish", typeof(string));
            //dt.Columns.Add("TaskIDCount", typeof(int));
            //dt.Columns.Add("Quantity", typeof(int));
            //dt.Columns.Add("Spares", typeof(int));
            //dt.Columns.Add("Status", typeof(string));
            //dt.Columns.Add("PercentComplete", typeof(int));

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
                if (component.Pictures != null)
                {
                    row["Pictures"] = component.Pictures;
                }
                else
                {
                    row["Pictures"] = DBNull.Value;
                }
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

        #endregion // Component operations.

        #region Tasks Table Operations

        #region Create

        private bool AddTaskDataTableToDatabase(DataTable dt)
        {
            var adapter = new OleDbDataAdapter();

            try
            {
                adapter.SelectCommand = new OleDbCommand("SELECT * FROM Tasks", Connection);

                var cbr = new OleDbCommandBuilder(adapter);

                cbr.GetDeleteCommand();
                cbr.GetInsertCommand();
                adapter.UpdateCommand = cbr.GetUpdateCommand();
                //Console.WriteLine(cbr.GetInsertCommand().CommandText);

                Connection.Open();
                adapter.Update(dt);
                Connection.Close();
                Console.WriteLine("Tasks Loaded.");
            }
            catch (OleDbException ex)
            {
                Connection.Close();
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace, "OledbException Error");
                return false;
            }
            catch (Exception x)
            {
                Connection.Close();
                MessageBox.Show(x.Message + "\n\n" + x.StackTrace, "Exception Error");
                return false;
            }

            return true;
        }

        #endregion

        #region Read
        public static List<TaskModel> GetAllTasks()
        {
            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                List<TaskModel> results = connection.Query<TaskModel>("SELECT * FROM Tasks").ToList();

                return results;
            }
        }
        private void GetTasks(ProjectModel project)
        {
            List<TaskModel> projectTaskList = GetProjectTaskList(project.JobNumber, project.ProjectNumber);

            foreach (ComponentModel component in project.Components)
            {
                var tasks = from t in projectTaskList
                            where t.Component == component.Component
                            orderby t.TaskID ascending
                            select t;

                component.AddTaskList(tasks.ToList());
            }

            foreach (ComponentModel component in project.Components)
            {
                //Console.WriteLine(component.Name);

                foreach (TaskModel task in component.Tasks)
                {
                    //Console.WriteLine($"   {task.ID} {task.TaskName}");
                    task.HasInfo = true;
                }
            }
        }

        public string GetTaskPredecessors(string jobNumber, int projectNumber, string component, int taskID)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable dt = new DataTable();
            string predecessors = "";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    string queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";

                    adapter.SelectCommand = new OleDbCommand(queryString, connection);
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
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    OleDbCommand sqlCommand = new OleDbCommand("SELECT FinishDate from Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID", connection);

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

        private static DateTime? GetLatestPredecessorFinishDate(string jobNumber, int projectNumber, string component, string predecessors)
        {
            DateTime? latestFinishDate = null;
            DateTime? currentDate = null;
            string[] predecessorArr;
            string predecessor;

            predecessorArr = predecessors.Split(',');

            foreach (string currPredecessor in predecessorArr)
            {
                predecessor = currPredecessor.Trim();
                currentDate = GetFinishDate(jobNumber, projectNumber, component, Convert.ToInt16(predecessor));

                if (latestFinishDate == null || latestFinishDate < currentDate)
                {
                    latestFinishDate = currentDate;
                }
            }

            return latestFinishDate;
        }

        public List<TaskModel> GetProjectTaskList(string jobNumber, int projectNumber)
        {
            List<TaskModel> taskList = new List<TaskModel>();

            string queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber";

            OleDbCommand cmd = new OleDbCommand(queryString, Connection);

            cmd.Parameters.AddWithValue("@jobNumber", jobNumber);
            cmd.Parameters.AddWithValue("@projectNumber", projectNumber);

            try
            {
                if (Connection.State == ConnectionState.Closed)
                {
                    Connection.Open(); 
                }

                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            taskList.Add(new TaskModel
                            (
                                    taskName: rdr["TaskName"],
                                      taskID: rdr["TaskID"],
                                          id: rdr["ID"],
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

        private string FindTask(DataTable dataTable, string jobNumber, int projectNumber, string component, int taskID)
        {
            DataRow task = dataTable.Rows.Cast<DataRow>().FirstOrDefault(x => (string)x["JobNumber"] == jobNumber && (int)x["ProjectNumber"] == projectNumber && (int)x["TaskID"] == taskID);

            return task["TaskName"].ToString();
        }

        private DataTable CreateDataTableFromTaskList(ProjectModel project)
        {
            DataTable dt = new DataTable();
            int i;

            // These three lines add the necessary columns to the datatable without adding data.

            string queryString = "SELECT * FROM Tasks WHERE ID = 0";

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            //dt.Columns.Add("ProjectNumber", typeof(int));
            //dt.Columns.Add("JobNumber", typeof(string));
            //dt.Columns.Add("Component", typeof(string));
            //dt.Columns.Add("TaskID", typeof(int));
            //dt.Columns.Add("TaskName", typeof(string));
            //dt.Columns.Add("Duration", typeof(string));
            //dt.Columns.Add("StartDate", typeof(DateTime));
            //dt.Columns.Add("FinishDate", typeof(DateTime));
            //dt.Columns.Add("EarliestStartDate", typeof(DateTime));
            //dt.Columns.Add("Predecessors", typeof(string));
            //dt.Columns.Add("Machines", typeof(string));
            //dt.Columns.Add("Machine", typeof(string));
            //dt.Columns.Add("Resources", typeof(string));
            //dt.Columns.Add("Resource", typeof(string));
            //dt.Columns.Add("Hours", typeof(int));
            //dt.Columns.Add("ToolMaker", typeof(string));
            //dt.Columns.Add("Operator", typeof(string));
            //dt.Columns.Add("Priority", typeof(string));
            //dt.Columns.Add("Status", typeof(string));
            //dt.Columns.Add("DateAdded", typeof(DateTime));
            //dt.Columns.Add("Notes", typeof(string));
            //dt.Columns.Add("Initials", typeof(string));
            //dt.Columns.Add("DateCompleted", typeof(string));

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
                    row["Resource"] = task.Personnel;
                    row["Resources"] = task.Resources;
                    row["Machine"] = task.Machine;
                    row["Priority"] = task.Priority;
                    row["DateAdded"] = task.DateAdded;
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

        // Helper method for setting appointment resources.
        public DataTable GetTasksWithChangedResources(int projectNumber, string taskName)
        {
            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                DataTable dt = new DataTable();

                string queryString = "SELECT * FROM Tasks WHERE ProjectNumber = @projectNumber AND TaskName = @taskName";

                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

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

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

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
            string selectStatment = "Projects.JobNumber, Projects.ProjectNumber, TaskName, Duration, Tasks.StartDate, FinishDate, Resource, Hours";
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

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                OleDbCommand cmd = new OleDbCommand(queryString, connection);

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

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                OleDbCommand cmd = new OleDbCommand(queryString, connection);

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
        public static DataTable GetAppointmentData(string department, bool grouped = false)
        {
            DataTable dt = new DataTable();

            try
            {
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    string queryString = SetQueryString(department);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);
                    adapter.Fill(dt);

                    dt.Columns.Add("PercentComplete", typeof(int));

                    foreach (DataRow nrow in dt.Rows)
                    {
                        if (nrow["Status"].ToString() == "Completed")
                        {
                            nrow["PercentComplete"] = 100;
                        }

                        if (nrow["Resources"].ToString() == "")
                        {
                            nrow["Resources"] = "<ResourceIds>\r\n<ResourceId Value=\"~Xtra#Base64AAEAAAD/////AQAAAAAAAAAGAQAAAAROb25lCw==\" />\r\n</ResourceIds>";
                        }
                    }

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
        // This method is for capturing changes made to tasks in the department schedule view.
        public bool UpdateTask(string jobNumber, int projectNumber, string component, int taskID, DateTime? startDate, DateTime? finishDate, string machine, string resource, string resources)
        {
            try
            {
                using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    string queryString;

                    queryString = "UPDATE Tasks SET StartDate = @startDate, FinishDate = @finishDate, Machine = @machine, Resource = @resource, Resources = @resources " +
                                    "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";

                    if (startDate == new DateTime(0001, 1, 1))
                    {
                        startDate = null;
                    }

                    if (finishDate == new DateTime(0001, 1, 1))
                    {
                        finishDate = null;
                    }

                    if (startDate == null)
                    {
                        goto SkipBackDating;
                    }

                    string predecessors = GetTaskPredecessors(jobNumber, projectNumber, component, taskID);
                    DateTime? latestPredecessorFinishDate = GetLatestPredecessorFinishDate(jobNumber, projectNumber, component, predecessors);

                    if (predecessors != "" && latestPredecessorFinishDate > startDate)
                    {
                        DialogResult dialogResult = MessageBox.Show("There is overlap between this task and one or more predecessors.  \n" +
                                                                    "Do you wish to push the overlapping predecessors back?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                        if (dialogResult == DialogResult.Yes)
                        {
                            // TODO: Validate this process.
                            BackDatePredecessorTasks(projectNumber, component, taskID, startDate);
                        }
                        else
                        {
                            return false;
                        }
                    }

                SkipBackDating:

                    var p = new
                    {
                        startDate,
                        finishDate,
                        machine,
                        resource,
                        resources,
                        jobNumber,
                        projectNumber,
                        component,
                        taskID
                    };

                    connection.Execute(queryString, p);
                }

                MoveSuccessors(jobNumber, projectNumber, component, finishDate, taskID);
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                return false;
            } 
            
        }
        public static bool UpdateTask(TaskModel task, ComponentModel component)
        {
            try
            {
                bool batchUpdateTasks = false;

                using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    string queryString;

                    queryString = "UPDATE Tasks SET StartDate = @StartDate, FinishDate = @finishDate, Machine = @Machine, Resource = @Resource, Resources = @Resources " +
                                  "WHERE ID = @ID";

                    if (task.StartDate == new DateTime(0001, 1, 1))
                    {
                        task.StartDate = null;
                    }

                    if (task.FinishDate == new DateTime(0001, 1, 1))
                    {
                        task.FinishDate = null;
                    }

                    if (task.StartDate == null)
                    {
                        goto SkipBackDating;
                    }

                    DateTime? latestPredecessorFinishDate = component.GetLatestPredecessorFinishDate(task.Predecessors);

                    if (latestPredecessorFinishDate == null)
                    {
                        goto SkipBackDating;
                    }

                    if (task.Predecessors != "" && latestPredecessorFinishDate > task.StartDate)
                    {
                        DialogResult dialogResult = MessageBox.Show("There is overlap between this task and one or more predecessors.  \n" +
                                                                    "Do you wish to push the overlapping predecessors back?", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                        if (dialogResult == DialogResult.Yes)
                        {
                            // TODO: Validate this process.
                            BackDateTask2(task.TaskID, component.Tasks, (DateTime)task.StartDate, true);
                            batchUpdateTasks = true;
                        }
                        else if (dialogResult == DialogResult.No)
                        {

                        }
                        else if (dialogResult == DialogResult.Cancel)
                        {
                            return false;
                        }
                    }

                SkipBackDating:

                    if (component.SuccessorOverlap(task))
                    {
                        DialogResult dialogResult = MessageBox.Show("There is overlap between this task and one or more successors. \n" +
                                                                    "Do you wish to push these tasks forward?", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                        if (dialogResult == DialogResult.Yes)
                        {
                            if (task.FinishDate != null)
                            {
                                UpdateStartAndFinishDates(task.TaskID, component.Tasks, (DateTime)task.FinishDate);
                                batchUpdateTasks = true;
                            }
                        }
                        else if (dialogResult == DialogResult.No)
                        {

                        }
                        else if (dialogResult == DialogResult.Cancel)
                        {
                            return false;
                        }
                    }

                    var p = new
                    {
                        task.StartDate,
                        task.FinishDate,
                        task.Machine,
                        task.Resource,
                        task.Resources,
                        task.ID
                    };

                    connection.Execute(queryString, p);
                }

                if (batchUpdateTasks == true)
                {
                    UpdateTaskDates(component.Tasks); 
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                return false;
            }

        }
        // string machine, string resource, 
        public bool UpdateTask(string jobNumber, int projectNumber, string component, int taskID, DateTime startDate, DateTime finishDate, bool overlapAllowed)
        {
            try
            {
                string queryString;

                queryString = "UPDATE Tasks SET StartDate = @startDate, FinishDate = @finishDate " +
                              "WHERE ID = @ID";

                OleDbConnection Connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName));
                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

                adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

                string predecessors = GetTaskPredecessors(jobNumber, projectNumber, component, taskID);

                if (predecessors != "" && GetLatestPredecessorFinishDate(jobNumber, projectNumber, component, predecessors) > startDate && !overlapAllowed)
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

                MoveSuccessors(jobNumber, projectNumber, component, finishDate, taskID);
                return true;
            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message);
                return false;
            }
        }

        public bool UpdateTaskResource(string jobNumber, int projectNumber, string component, int taskID, string resource)
        {
            try
            {
                string queryString = "UPDATE Tasks SET Resource = @resource " +
                                     "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";
                OleDbConnection Connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName));
                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

                adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

                adapter.UpdateCommand.Parameters.AddWithValue("@resource", resource);
                adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@component", component);
                adapter.UpdateCommand.Parameters.AddWithValue("@taskID", taskID);

                Connection.Open();
                adapter.UpdateCommand.ExecuteNonQuery();

                return true;
            }
            catch(Exception)
            {
                return false;
            }
            finally
            {
                Connection.Close();
                Connection.Dispose();
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
                adapter.SelectCommand.Parameters.Add("@component", OleDbType.VarChar, 45).Value = component;

                var cbr = new OleDbCommandBuilder(adapter);

                adapter.DeleteCommand = cbr.GetDeleteCommand();
                adapter.InsertCommand = cbr.GetInsertCommand();
                adapter.UpdateCommand = cbr.GetUpdateCommand();

                adapter.Fill(dt);

                //foreach(DataRow nrow in UpdateTaskDataTable2(taskList, dt).Rows)
                //{
                //    Console.WriteLine(nrow["ID"].ToString() + " " + nrow["TaskName"].ToString());
                //}
                RenumberTaskIDs(taskList);

                Connection.Open();

                UpdateTaskDataTable(jobNumber, projectNumber, component, taskList, dt);

                adapter.Update(dt);

                //MessageBox.Show($"{component} tasks updated!");
            }
            finally
            {
                Connection.Close();
            }
        }

        private List<TaskModel> RenumberTaskIDs(List<TaskModel> taskList)
        {
            for (int i = 0; i < taskList.Count; i++)
            {
                taskList[i].SetTaskID(i + 1);
                Console.WriteLine(taskList[i].TaskID);
            }

            return taskList;
        }

        private DataTable UpdateTaskDataTable(string jobNumber, int projectNumber, string component, List<TaskModel> taskList, DataTable taskDataTable)
        {
            UpdateTasksInDataTable(taskList, taskDataTable);
            DeleteTasksFromDatatable(taskList, taskDataTable);
            AddTasksToDataTable(jobNumber, projectNumber, component, taskList, taskDataTable);
            
            return taskDataTable;
        }

        // DO NOT CALL THIS METHOD AFTER ADDING TASKS TO DATATABLE.  NEW TASKS DO NOT HAVE IDS ASSIGNED TO THEM AND THEREFORE CANNOT BE READ.
        private DataTable DeleteTasksFromDatatable(List<TaskModel> taskList, DataTable taskDataTable)
        {
            int id;
            foreach (DataRow row in taskDataTable.Rows)
            {
                id = (int)row["ID"];
                if (!taskList.Exists(x => x.ID == (int)row["ID"]))
                {
                    row.Delete();
                }
            }

            return taskDataTable;
        }

        private DataTable AddTasksToDataTable(string jobNumber, int projectNumber, string component, List<TaskModel> taskList, DataTable taskDataTable)
        {
            var tasksToAdd = from t in taskList
                             where t.ID == 0
                             select t;

            foreach (TaskModel task in tasksToAdd.ToList())
            {
                DataRow row = taskDataTable.NewRow();

                row["TaskID"] = task.TaskID;
                row["JobNumber"] = jobNumber;
                row["ProjectNumber"] = projectNumber;
                row["Component"] = component;
                row["TaskName"] = task.TaskName;
                row["Hours"] = task.Hours;
                row["Duration"] = task.Duration;
                row["Machine"] = task.Machine;
                row["Resources"] = task.Resources;
                row["Resource"] = task.Personnel;
                row["Predecessors"] = task.Predecessors;
                row["Priority"] = task.Priority;
                row["Notes"] = task.Notes;

                taskDataTable.Rows.Add(row);
            }

            return taskDataTable;
        }

        // DO NOT CALL THIS METHOD AFTER DELETING TASKS FROM DATATABLE. FOR SOME REASON IT THROWS AN ERROR WHEN SELECTING ROWS.
        private DataTable UpdateTasksInDataTable(List<TaskModel> taskList, DataTable taskDataTable)
        {
            var tasksToUpdate = from t in taskList
                                where t.ID != 0
                                select t;

            Console.WriteLine("Task Datatable");

            foreach (DataRow row in taskDataTable.Rows)
            {
                Console.WriteLine($"{row["ID"]} {row["TaskID"]} {row["TaskName"]}");
            }

            Console.WriteLine("Task list");

            foreach (var item in taskList)
            {
                Console.WriteLine($"{item.ID} {item.TaskID} {item.TaskName}");
            }

            foreach (TaskModel task in tasksToUpdate.ToList())
            {
                var rows = from DataRow myRow in taskDataTable.Rows
                           where (int)myRow["ID"] == task.ID
                           select myRow;

                DataRow row = rows.FirstOrDefault();

                row["TaskID"] = task.TaskID;
                row["TaskName"] = task.TaskName;
                row["Hours"] = task.Hours;
                row["Duration"] = task.Duration;
                row["Machine"] = task.Machine;
                row["Resources"] = task.Resources;
                row["Resource"] = task.Personnel;
                row["Predecessors"] = task.Predecessors;
                row["Priority"] = task.Priority;
                row["Notes"] = task.Notes;
            }

            return taskDataTable;
        }

        // This means both machines and personnel.
        public static void SetTaskResources(object s, CellValueChangedEventArgs ev, SchedulerStorage schedulerStorage)
        {
            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                string taskName;
                int? projectNumber;

                var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

                WorkLoadModel wli = grid.GetRow(ev.RowHandle) as WorkLoadModel;

                DataTable dt = new DataTable();

                //string queryString = "UPDATE Tasks SET Resource = @resource " +
                //                     "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND TaskName = @taskName";

                string queryString = "SELECT * FROM Tasks WHERE ProjectNumber = @projectNumber AND TaskName = @taskName";

                var command = new SqlCommand(queryString);

                command.Connection = connection;
                command.CommandType = CommandType.Text;

                if (wli.MWONumber != null) // If there is an MWONumber use it for the project number.
                {
                    projectNumber = wli.MWONumber;
                }
                else
                {
                    projectNumber = wli.ProjectNumber;
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

                //builder.QuotePrefix = "[";
                //builder.QuoteSuffix = "]";

                adapter.Fill(dt);

                foreach (DataRow nrow in dt.Rows)
                {
                    nrow["Resource"] = ev.Value.ToString();
                    nrow["Resources"] = GenerateResourceIDsString(nrow["Machine"].ToString(), ev.Value.ToString(), schedulerStorage);
                }

                //adapter.UpdateCommand.Parameters.AddWithValue("@resource", ev.Value.ToString());
                //adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", grid.GetRowCellValue(ev.RowHandle, grid.Columns["JobNumber"]));
                //adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", grid.GetRowCellValue(ev.RowHandle, grid.Columns["ProjectNumber"]));

                //adapter.UpdateCommand.Parameters.AddWithValue("@taskName", taskName);

                //connection.Open();
                //adapter.UpdateCommand.ExecuteNonQuery();

                adapter.Update(dt);
            }
        }

        private static string GenerateResourceIDsString(string machine, string resource, SchedulerStorage schedulerStorage)
        {
            AppointmentResourceIdCollection appointmentResourceIdCollection = new AppointmentResourceIdCollection();
            Resource res;
            int machineCount = schedulerStorage.Resources.Items.Where(x => x.Id.ToString() == machine).Count();
            int resourceCount = schedulerStorage.Resources.Items.Where(x => x.Id.ToString() == resource).Count();

            if (machineCount == 0 && resourceCount == 0)
            {
                res = schedulerStorage.Resources.Items.GetResourceById("None");
                appointmentResourceIdCollection.Add(res.Id);
            }

            if (machine != "" && machineCount == 1)
            {
                res = schedulerStorage.Resources.Items.GetResourceById(machine);
                appointmentResourceIdCollection.Add(res.Id);
            }

            if (resource != "" && resourceCount == 1)
            {
                res = schedulerStorage.Resources.Items.GetResourceById(resource);
                appointmentResourceIdCollection.Add(res.Id);
            }

            AppointmentResourceIdCollectionXmlPersistenceHelper helper = new AppointmentResourceIdCollectionXmlPersistenceHelper(appointmentResourceIdCollection);
            return helper.ToXml();
        }

        private static DateTime AddBusinessDays(DateTime date, string durationSt)
        {
            int days;
            string[] duration = null;

            if (durationSt != null)
            {
                duration = durationSt.Split(' '); 
            }
            else
            {
                MessageBox.Show("Duration cannot be null.");
            }

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

        public static void ChangeTaskDate(string fieldName, TaskModel task, ComponentModel component)
        {
            // Checks if the start date changed.
            if (fieldName == "StartDate")
            {
                if (task.StartDate == null)
                {
                    task.FinishDate = null;
                }
                else
                {
                    task.FinishDate = AddBusinessDays((DateTime)task.StartDate, task.Duration);
                }
            }

            TaskModel temp = component.Tasks.Find(x => x.ID == task.ID);

            temp.StartDate = task.StartDate;
            temp.FinishDate = task.FinishDate;


            if (task.FinishDate != null)
            {
                UpdateStartAndFinishDates(task.TaskID, component.Tasks, (DateTime)task.FinishDate); 
            }

            UpdateTaskDates(component.Tasks);
        }

        private static void UpdateTaskDates(List<TaskModel> tasks)
        {
            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                string queryString = "UPDATE Tasks " +
                                     "SET StartDate = @StartDate, FinishDate = @FinishDate " +
                                     "WHERE ID = @ID";

                foreach (var task in tasks)
                {
                    var p = new { task.StartDate, task.FinishDate, task.ID };

                    connection.Execute(queryString, p);
                }
            }
        }

        public static void MoveSuccessors(string jobNumber, int projectNumber, string component, DateTime? currentTaskFinishDate, int currentTaskID)
        {
            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                DataTable datatable = new DataTable();
                string queryString;
                queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component ORDER BY TaskID ASC";

                OleDbDataAdapter adapter = new OleDbDataAdapter();
                adapter.SelectCommand = new OleDbCommand(queryString, connection);
                adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 20).Value = jobNumber;
                adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = projectNumber;
                adapter.SelectCommand.Parameters.Add("@component", OleDbType.VarChar, 30).Value = component;
                OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

                Console.WriteLine("Move Descendents");

                adapter.Fill(datatable);

                UpdateStartAndFinishDates(currentTaskID, datatable, currentTaskFinishDate);

                adapter.Update(datatable); 
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
                List<TaskModel> leadingTasks = component.Tasks.FindAll(x => x.Predecessors == "" || x.Predecessors == null);

                foreach (var task in leadingTasks)
                {
                    task.StartDate = forwardDate;
                    task.FinishDate = AddBusinessDays(forwardDate, task.Duration);
                    ForwardDateTask(task.TaskID, component.Tasks, (DateTime)task.FinishDate);
                }

                tasksToUpdate.AddRange(component.Tasks);
            }

            UpdateTaskDates(tasksToUpdate);
        }
        private static void ClearSelectedComponentDates(List<ComponentModel> selectedComponents)
        {
            foreach (ComponentModel component in selectedComponents)
            {
                foreach (TaskModel task in component.Tasks)
                {
                    task.StartDate = null;
                    task.FinishDate = null;
                }
            }
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

        private static void ForwardDateTask(int predecessorID, List<TaskModel> componentTasks, DateTime predecessorFinishDate)
        {
            var result = from tasks in componentTasks
                         where tasks.HasMatchingPredecessor(predecessorID)
                         select tasks;

            foreach (TaskModel task in result)
            {
                if (task.StartDate == null || task.StartDate < predecessorFinishDate)
                {
                    task.StartDate = predecessorFinishDate;
                    task.FinishDate = AddBusinessDays((DateTime)task.StartDate, task.Duration);

                    ForwardDateTask(task.TaskID, componentTasks, (DateTime)task.FinishDate);
                }
            }
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
        private static void UpdateStartAndFinishDates(int id, DataTable dt, DateTime? fd)
        {
            if (fd == null)
            {
                return;
            }

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
        private static List<TaskModel> UpdateStartAndFinishDates(int id, List<TaskModel> componentTasks, DateTime? finishDate, bool fillBlanks = false, bool pullBackStartDates = false, bool promptToPushDatesForward = false)
        {
            var result = from tasks in componentTasks
                         where tasks.HasMatchingPredecessor(id)
                         select tasks;

            Console.WriteLine("Update Start and Finish Dates");

            foreach (TaskModel task in result)
            {
                if (task.StartDate == null)
                {
                    if (fillBlanks == true)
                    {
                        task.StartDate = finishDate;
                        task.FinishDate = AddBusinessDays(Convert.ToDateTime(task.StartDate), task.Duration.ToString());

                        Console.WriteLine(id + " " + task.TaskID + " " + task.TaskName + " " + task.StartDate + " " + task.FinishDate + " " + task.Predecessors); 
                    }
                }
                else if (Convert.ToDateTime(task.StartDate) < finishDate) // If start date of current task comes before finish date of predecessor.
                {
                    task.StartDate = finishDate;
                    task.FinishDate = AddBusinessDays(Convert.ToDateTime(task.StartDate), task.Duration.ToString());
                    Console.WriteLine(id + " " + task.TaskID + " " + task.TaskName + " " + Convert.ToDateTime(task.StartDate).ToShortDateString() + " " + Convert.ToDateTime(task.FinishDate).ToShortDateString() + " " + task.Predecessors);
                    //Console.WriteLine(currentTaskID + " " + currentTaskFinishDate + " " + nrow2["TaskID"] + " " + predecessorArr[i2].ToString() + " " + nrow2["Predecessors"]);
                }
                else if (Convert.ToDateTime(task.StartDate) > finishDate && pullBackStartDates == true) // If start date of current task comes after the finish date of predecessor.
                {
                    task.StartDate = finishDate;
                    task.FinishDate = AddBusinessDays(Convert.ToDateTime(task.StartDate), task.Duration.ToString());
                    Console.WriteLine(id + " " + task.TaskID + " " + task.TaskName + " " + Convert.ToDateTime(task.StartDate).ToShortDateString() + " " + Convert.ToDateTime(task.FinishDate).ToShortDateString() + " " + task.FinishDate);
                }

                if (task.FinishDate != null)
                    UpdateStartAndFinishDates(task.TaskID, componentTasks, Convert.ToDateTime(task.FinishDate), fillBlanks, pullBackStartDates);
            }

            return componentTasks;
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
        public static void BackDateTasks(List<ComponentModel> selectedComponents, DateTime backDate)
        {
            int highestTaskID;
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
                highestTaskID = component.Tasks.Max(x => x.TaskID);

                BackDateTask(highestTaskID, component.Tasks, backDate);

                tasksToUpdate.AddRange(component.Tasks);
            }

            UpdateTaskDates(tasksToUpdate);
        }
        public static void BackDatePredecessorTasks(int projectNumber, string component, int taskID, DateTime? backDate)
        {
            DataTable dt = new DataTable();
            string queryString = "SELECT * FROM Tasks WHERE ProjectNumber = @ProjectNumber AND Component = @Component ORDER BY TaskID";

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

                adapter.SelectCommand.Parameters.Add("@ProjectNumber", OleDbType.Integer, 12).Value = projectNumber;
                adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = component;

                OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

                adapter.Fill(dt);

                BackDateTask(taskID, component, false, backDate, dt);

                adapter.UpdateCommand = builder.GetUpdateCommand();

                adapter.Update(dt);
            }
        }

        public static void BackDatePredecessorTasks(int taskID, List<TaskModel> componentTasks, bool skipDatedTasks, DateTime? descendantStartDate)
        {

        }
        private static void BackDateTask(int taskID, List<TaskModel> componentTasks, bool skipDatedTasks, DateTime? descendantStartDate)
        {
            string[] predecessors;
            TaskModel task = componentTasks.Find(x => x.TaskID == taskID);

            task.FinishDate = descendantStartDate;
            task.StartDate = SubtractBusinessDays((DateTime)task.FinishDate, task.Duration);

            if (task.Predecessors.Contains(','))
            {
                predecessors = task.Predecessors.Split(',');

                foreach (string id in predecessors)
                {
                    BackDateTask(Convert.ToInt32(id), componentTasks, skipDatedTasks, task.StartDate);
                }
            }
            // If a task has one predecessor.
            // Backdate the one predecessor.
            else if (task.Predecessors != "")
            {
                BackDateTask(Convert.ToInt32(task.Predecessors), componentTasks, skipDatedTasks, task.StartDate);
            }
            // If a task has no predecessors.
            // Exit method.
            else if (task.Predecessors == "")
            {
                return;
            }
        }
        private static void BackDateTask(int taskID, string component, bool skipDatedTasks, DateTime? descendantStartDate, DataTable projectTaskTable)
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
        private static void BackDateTask(int id, List<TaskModel> componentTasks, DateTime descendentStartDate)
        {
            string[] predecessors;

            TaskModel task = componentTasks.Find(x => x.TaskID == id);

            task.FinishDate = descendentStartDate;
            task.StartDate = SubtractBusinessDays((DateTime)task.FinishDate, task.Duration); 

            if (task.Predecessors.Contains(','))
            {
                predecessors = task.Predecessors.Split(',');

                foreach (string predecessor in predecessors)
                {
                    BackDateTask(Convert.ToInt32(predecessor), componentTasks, (DateTime)task.StartDate);
                }
            }
            else if (task.Predecessors != "")
            {
                BackDateTask(Convert.ToInt32(task.Predecessors), componentTasks, (DateTime)task.StartDate);
            }
        }
        private static void BackDateTask2(int id, List<TaskModel> componentTasks, DateTime descendentStartDate, bool movedTask = false)
        {
            string[] predecessors;

            TaskModel task = componentTasks.Find(x => x.TaskID == id);

            if (movedTask == false)
            {
                if (task.FinishDate > descendentStartDate)
                {
                    task.FinishDate = descendentStartDate;
                    task.StartDate = SubtractBusinessDays((DateTime)task.FinishDate, task.Duration); 
                }
            }
            else
            {
                task.FinishDate = AddBusinessDays((DateTime)task.StartDate, task.Duration);
                task.StartDate = descendentStartDate; 
            }

            if (task.Predecessors.Contains(','))
            {
                predecessors = task.Predecessors.Split(',');

                foreach (string predecessor in predecessors)
                {
                    BackDateTask2(Convert.ToInt32(predecessor), componentTasks, (DateTime)task.StartDate);
                }
            }
            else if (task.Predecessors != "")
            {
                BackDateTask2(Convert.ToInt32(task.Predecessors), componentTasks, (DateTime)task.StartDate);
            }
        }
        public static DateTime SubtractBusinessDays(DateTime finishDate, string durationSt)
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

        public static void UpdateTasksTable(object s, CellValueChangedEventArgs ev, string resources = "")
        {
            try
            {
                var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

                //queryString = "UPDATE Tasks SET JobNumber = @jobNumber, Component = @component, TaskID = @taskID, TaskName = @taskName, " +
                //    "Duration = @duration, StartDate = @startDate, FinishDate = @finishDate, Predecessor = @predecessor, Machines = @machines, " +
                //    "Machine = @machine, Person = @person, Priority = @priority WHERE ID = @tID";

                using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    string queryString;
                    var p = new DynamicParameters();

                    if (ev.Column.FieldName == "TaskName")
                    {
                        queryString = "UPDATE Tasks SET TaskName = @TaskName WHERE (ID = @ID)";

                        p.Add("@TaskName", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Notes")
                    {
                        queryString = "UPDATE Tasks SET Notes = @Notes WHERE (ID = @ID)";

                        p.Add("@Notes", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Hours")
                    {
                        queryString = "UPDATE Tasks SET Hours = @Hours WHERE (ID = @ID)";

                        if (ev.Value.ToString() != "")
                        {
                            p.Add("@Hours", ev.Value.ToString());
                        }
                        else
                        {
                            p.Add("@Hours", 0);
                        }
                    }
                    else if (ev.Column.FieldName == "Duration")
                    {
                        queryString = "UPDATE Tasks SET Duration = @Duration WHERE (ID = @ID)";

                        p.Add("@Duration", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Predecessors")
                    {
                        queryString = "UPDATE Tasks SET Predecessors = @Predecessors WHERE (ID = @ID)";

                        p.Add("@Predecessors", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Machine")
                    {
                        queryString = "UPDATE Tasks SET Machine = @Machine, Resources = @Resources WHERE (ID = @ID)";

                        p.Add("@Machine", ev.Value.ToString());
                        p.Add("@Resources", resources);
                    }
                    else if (ev.Column.FieldName == "Resource")
                    {
                        queryString = "UPDATE Tasks SET Resource = @Resource, Resources = @Resources WHERE (ID = @ID)";

                        p.Add("@Resource", ev.Value.ToString());
                        p.Add("@Resources", resources);
                    }
                    else
                    {
                        MessageBox.Show(ev.Column.ToString() + " column is not editable.");
                        return;
                    }

                    p.Add("@ID", (grid.GetRowCellValue(ev.RowHandle, grid.Columns["ID"])));

                    connection.Execute(queryString, p);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static void UpdateTaskProjectInfo(ProjectModel project)
        {
            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                string queryString;
                var p = new DynamicParameters(project);

                queryString = "UPDATE Tasks SET JobNumber = @JobNumber, ProjectNumber = @ProjectNumber WHERE (ProjectNumber = @OldProjectNumber)";

                connection.Execute(queryString, p);
            }
        }

        #endregion

        #region Delete

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

        #endregion

        private DataTable CreateDataTableFromTaskList(ProjectModel project, List<TaskModel> taskList)
        {
            DataTable dt = new DataTable();
            string component = "";
            int i = 1;

            // These three lines add the necessary columns to the datatable without adding data.

            string queryString = "SELECT * FROM Tasks WHERE ID = 0";

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

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
                row["Resources"] = task.Resources;
                row["Machine"] = task.Machine;
                row["Priority"] = task.Priority;
                row["DateAdded"] = task.DateAdded;
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
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    OleDbCommand cmd = new OleDbCommand("INSERT INTO Resources (ResourceName, ResourceType) VALUES (@resourceName, @resourceType)", connection);

                    cmd.Parameters.AddWithValue("resourceID", resourceName);
                    cmd.Parameters.AddWithValue("resourceType", resourceType);

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
            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                DataTable dt = new DataTable();

                string queryString = "SELECT * From Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID ORDER BY Resources.ResourceName ASC";

                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

                adapter.Fill(dt);

                return dt;
            }
        }

        public static List<string> GetResourceList()
        {
            DataTable dt = new DataTable();
            List<string> ResourceList = new List<string>();

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                string queryString = "SELECT * From Resources ORDER BY ResourceName ASC";

                OleDbCommand cmd = new OleDbCommand(queryString, connection);

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
            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                string queryString = "SELECT ID FROM Resources WHERE ResourceName = @resourceName";

                OleDbCommand sqlCommand = new OleDbCommand(queryString, connection);

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

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                string queryString = "SELECT Resources.ResourceName, Roles.Role FROM Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID WHERE Roles.Role = @role ORDER BY Resources.ResourceName ASC";

                OleDbCommand cmd = new OleDbCommand(queryString, connection);

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
            
            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                string queryString = "SELECT ResourceName, ResourceType, Resources.ID, Role, Departments.Department From (Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID) LEFT OUTER JOIN Departments ON Roles.DepartmentID = Departments.ID ORDER BY ResourceName ASC";

                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

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

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                OleDbCommand cmd = new OleDbCommand(queryString, connection);

                cmd.Parameters.AddWithValue("@resourceName", resource);

                connection.Open();

                return cmd.ExecuteScalar().ToString();
            }
        }

        public static List<string> GetAllResourcesOfType(string resourceType)
        {
            string queryString = "SELECT ResourceName FROM Resources WHERE ResourceType = @resourceType ORDER BY ResourceName";

            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
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

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                OleDbCommand cmd = new OleDbCommand(queryString, connection);

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
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    OleDbCommand cmd1 = new OleDbCommand("DELETE FROM Roles WHERE ResourceID = @resourceID ", connection);
                    OleDbCommand cmd2 = new OleDbCommand("DELETE FROM Resources WHERE ID = @resourceID ", connection);

                    cmd1.Parameters.AddWithValue("resourceID", GetResourceID(resourceName));
                    cmd2.Parameters.AddWithValue("ID", GetResourceID(resourceName));

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
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    OleDbCommand cmd = new OleDbCommand("INSERT INTO Roles (ResourceID, Role, DepartmentID) VALUES (@resourceID, @role, @departmentID)", connection);

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

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                string queryString = "SELECT COUNT(*) AS RoleCount, Role FROM Roles GROUP BY Role";

                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

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
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    OleDbCommand cmd = new OleDbCommand("DELETE FROM Roles WHERE ResourceID = @resourceID AND Role = @role", connection);

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

            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
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
                string queryString = "INSERT INTO Workload (ToolNumber, MWONumber, ProjectNumber, Stage, Customer, PartName, DeliveryInWeeks, StartDate, FinishDate, AdjustedDeliveryDate, MoldCost, Engineer, Designer, ToolMaker, RoughProgrammer, FinishProgrammer, ElectrodeProgrammer, Apprentice, Manifold, MoldBase, GeneralNotes) VALUES " +
                                                          "(@toolNumber, @mwoNumber, @projectNumber, @stage, @customer, @partName, @deliveryInWeeks, @startDate, @finishDate, @adjustedDeliveryDate, @moldCost, @engineer, @designer, @toolMaker, @roughProgrammer, @finishProgrammer, @electrodeProgrammer, @apprentice, @manifold, @moldBase, @generalNotes)";

                connection.Execute(queryString, workload);
            }
        }

        #endregion

        #region Read

        public static List<WorkLoadModel> GetWorkloads()
        {
            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                List<WorkLoadModel> workLoads = connection.Query<WorkLoadModel>("SELECT * FROM Workload").ToList();

                return workLoads;
            }
        }

        #endregion

        #region Update

        public static bool UpdateWorkloadTable(object s, CellValueChangedEventArgs ev)
        {
            var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

            if (grid.GetFocusedRowCellValue("ID").ToString() == "")
            {
                // I'm tricking the system here.
                return true;
            }

            WorkLoadModel wli = grid.GetFocusedRow() as WorkLoadModel;

            //queryString = "UPDATE Tasks SET JobNumber = @jobNumber, Component = @component, TaskID = @taskID, TaskName = @taskName, " +
            //    "Duration = @duration, StartDate = @startDate, FinishDate = @finishDate, Predecessor = @predecessor, Machines = @machines, " +
            //    "Machine = @machine, Person = @person, Priority = @priority WHERE ID = @tID";

            using (SqlConnection connection = new SqlConnection(Helper.CnnValue(SQLClientConnectionName)))
            {
                var cmd = new SqlCommand();
                cmd.CommandType = CommandType.Text;
                cmd.Connection = connection;

                if (ev.Column.FieldName == "ToolNumber")
                {
                    cmd.CommandText = "UPDATE WorkLoad SET ToolNumber = @toolNumber WHERE (ID = @tID)";

                    cmd.Parameters.AddWithValue("@toolNumber", wli.ToolNumber);
                }
                else if (ev.Column.FieldName == "MWONumber")
                {
                    cmd.CommandText = "UPDATE WorkLoad SET MWONumber = @mwoNumber WHERE (ID = @tID)";

                    cmd.Parameters.AddWithValue("@mwoNumber", wli.MWONumber);

                    //if (int.TryParse(ev.Value.ToString(), out int mwoNumber))
                    //{
                    //    cmd.Parameters.AddWithValue("@mwoNumber", mwoNumber);
                    //}
                    //else if(ev.Value.ToString() == "")
                    //{
                    //    cmd.Parameters.AddWithValue("@mwoNumber", DBNull.Value);
                    //}
                }
                else if (ev.Column.FieldName == "ProjectNumber")
                {
                    cmd.CommandText = "UPDATE WorkLoad SET ProjectNumber = @projectNumber WHERE (ID = @tID)";

                    cmd.Parameters.AddWithValue("@projectNumber", wli.ProjectNumber);

                    //if (int.TryParse(ev.Value.ToString(), out int projectNumber))
                    //{
                    //    cmd.Parameters.AddWithValue("@projectNumber", projectNumber);
                    //}
                    //else if(ev.ToString() == "")
                    //{
                    //    cmd.Parameters.AddWithValue("@projectNumber", DBNull.Value);
                    //}
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
                else if (ev.Column.FieldName == "Apprentice")
                {
                    cmd.CommandText = "UPDATE WorkLoad SET Apprentice = @apprentice WHERE (ID = @tID)";

                    if (ev.Value.ToString() != "")
                    {
                        cmd.Parameters.AddWithValue("@apprentice", ev.Value.ToString());
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@apprentice", "");
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
                connection.Open();
                cmd.ExecuteNonQuery();
                return true;
            }
        }

        public static bool UpdateWorkloadTable2(object s, CellValueChangedEventArgs ev)
        {
            var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

            if (grid.GetFocusedRowCellValue("ID").ToString() == "")
            {
                // I'm tricking the system here.
                return true;
            }

            using (IDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                string queryString;
                var p = new DynamicParameters();

                if (ev.Column.FieldName == "ToolNumber")
                {
                    queryString = "UPDATE WorkLoad SET ToolNumber = @toolNumber WHERE (ID = @tID)";

                    p.Add("@toolNumber", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "MWONumber")
                {
                    queryString = "UPDATE WorkLoad SET MWONumber = @mwoNumber WHERE (ID = @tID)";

                    p.Add("@mwoNumber", ev.Value);
                }
                else if (ev.Column.FieldName == "ProjectNumber")
                {
                    queryString = "UPDATE WorkLoad SET ProjectNumber = @projectNumber WHERE (ID = @tID)";

                    if (int.TryParse(ev.Value.ToString(), out int projectNumber))
                    {
                        p.Add("@projectNumber", projectNumber);
                    }
                    else if (ev.ToString() == "")
                    {
                        p.Add("@projectNumber", DBNull.Value);
                    }
                }
                else if (ev.Column.FieldName == "Stage")
                {
                    queryString = "UPDATE WorkLoad SET Stage = @stage WHERE (ID = @tID)";

                    p.Add("@stage", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "Customer")
                {
                    queryString = "UPDATE WorkLoad SET Customer = @customer WHERE (ID = @tID)";

                    p.Add("@customer", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "PartName")
                {
                    queryString = "UPDATE WorkLoad SET PartName = @partName WHERE (ID = @tID)";

                    p.Add("@partName", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "Engineer")
                {
                    queryString = "UPDATE WorkLoad SET Engineer = @engineer WHERE (ID = @tID)";

                    p.Add("@engineer", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "DeliveryInWeeks")
                {
                    queryString = "UPDATE WorkLoad SET DeliveryInWeeks = @deliveryInWeeks WHERE (ID = @tID)";

                    p.Add("@deliveryInWeeks", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "StartDate")
                {
                    queryString = "UPDATE WorkLoad SET StartDate = @startDate WHERE (ID = @tID)";

                    p.Add("@startDate", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "FinishDate")
                {
                    queryString = "UPDATE WorkLoad SET FinishDate = @finishDate WHERE (ID = @tID)";

                    p.Add("@finishDate", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "AdjustedDeliveryDate")
                {
                    queryString = "UPDATE WorkLoad SET AdjustedDeliveryDate = @adjustedDeliveryDate WHERE (ID = @tID)";

                    p.Add("@adjustedDeliveryDate", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "MoldCost")
                {
                    queryString = "UPDATE WorkLoad SET MoldCost = @moldCost WHERE (ID = @tID)";

                    p.Add("@moldCost", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "Designer")
                {
                    queryString = "UPDATE WorkLoad SET Designer = @designer WHERE (ID = @tID)";

                    p.Add("@designer", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "Designer")
                {
                    queryString = "UPDATE WorkLoad SET Designer = @designer WHERE (ID = @tID)";

                    p.Add("@designer", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "ToolMaker")
                {
                    queryString = "UPDATE WorkLoad SET ToolMaker = @toolMaker WHERE (ID = @tID)";

                    p.Add("@toolMaker", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "RoughProgrammer")
                {
                    queryString = "UPDATE WorkLoad SET RoughProgrammer = @roughProgrammer WHERE (ID = @tID)";

                    p.Add("@roughProgrammer", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "FinishProgrammer")
                {
                    queryString = "UPDATE WorkLoad SET FinishProgrammer = @finishProgrammer WHERE (ID = @tID)";

                    p.Add("@finishProgrammer", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "ElectrodeProgrammer")
                {
                    queryString = "UPDATE WorkLoad SET ElectrodeProgrammer = @electrodeProgrammer WHERE (ID = @tID)";

                    p.Add("@electrodeProgrammer", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "Apprentice")
                {
                    queryString = "UPDATE WorkLoad SET Apprentice = @apprentice WHERE (ID = @tID)";

                    p.Add("@apprentice", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "Manifold")
                {
                    queryString = "UPDATE WorkLoad SET Manifold = @manifold WHERE (ID = @tID)";

                    p.Add("@manifold", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "MoldBase")
                {
                    queryString = "UPDATE WorkLoad SET MoldBase = @moldBase WHERE (ID = @tID)";

                    p.Add("@moldBase", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "GeneralNotes")
                {
                    queryString = "UPDATE WorkLoad SET GeneralNotes = @generalNotes WHERE (ID = @tID)";

                    p.Add("@generalNotes", ev.Value.ToString());
                }
                else if (ev.Column.FieldName == "GeneralNotesRTF")
                {
                    queryString = "UPDATE WorkLoad SET GeneralNotesRTF = @generalNotesRTF WHERE (ID = @tID)";

                    p.Add("@generalNotesRTF", ev.Value.ToString());
                }
                else
                {
                    MessageBox.Show(ev.Column.ToString() + " column is not editable.");
                    return false;
                }

                p.Add("@tID", (grid.GetRowCellValue(ev.RowHandle, grid.Columns["ID"])));

                connection.Execute(queryString, p);

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
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    OleDbCommand cmd = new OleDbCommand("INSERT INTO WorkLoadColors (ProjectID, ColumnFieldName, ARGBColor) VALUES (@projectID, @columnFieldName, @aRGBColor)", connection);

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
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    OleDbCommand cmd = new OleDbCommand("SELECT * FROM WorkLoadColors", connection);

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
                using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
                {
                    OleDbCommand cmd = new OleDbCommand();
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

        public DataTable GetDailyDepartmentCapacities()
        {
            string queryString = "SELECT * FROM Departments";

            DataTable dt = new DataTable();

            using (OleDbConnection connection = new OleDbConnection(Helper.CnnValue(OLEDBConnectionName)))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

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
