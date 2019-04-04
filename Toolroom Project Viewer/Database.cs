using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using DevExpress.Spreadsheet;
using System.Globalization;
using System.Runtime.InteropServices;
using ClassLibrary;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.Xml;

namespace Toolroom_Project_Viewer
{
    class Database
    {
        static readonly string ConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=X:\TOOLROOM\Workload Tracking System\Database\Workload Tracking System DB.accdb";
        OleDbConnection Connection = new OleDbConnection(ConnString);

        DataTable TaskIDKey = new DataTable();

        #region Projects Table Operations

        #region Create

        public bool LoadProjectToDB(ProjectInfo project)
        {

            //if(result == DialogResult.Yes)
            //{
            //    //int baseIDNumber = getHighestProjectTaskID(project.JobNumber, project.ProjectNumber);
            //    //updateProjectData(pi);
            //    //foreach (Component component in project.ComponentList)
            //    //{
            //    //    foreach (TaskInfo task in component.TaskList)
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

        private bool AddProjectDataToDatabase(ProjectInfo project)
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

        #endregion

        #region Read

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
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Connection.Close();
                Connection.Dispose();
            }

            return pi;
        }

        public List<ProjectInfo> GetProjectInfoList()
        {
            string queryString = "SELECT * FROM Projects";
            OleDbCommand cmd;
            ProjectInfo pi;
            List<ProjectInfo> piList = new List<ProjectInfo>();

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
                            pi = new ProjectInfo
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

        public ProjectInfo GetProject(int projectNumber)
        {
            ProjectInfo project = GetProjectInfo(projectNumber);

            AddComponents(project);

            AddTasks(project);

            return project;
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

        public ProjectInfo GetProjectInfo(int projectNumber)
        {
            OleDbCommand cmd;
            ProjectInfo pi = null;
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
                            pi = new ProjectInfo
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

        #endregion

        #region Update

        public bool EditProjectInDB(ProjectInfo project)
        {
            try
            {
                ProjectInfo databaseProject = GetProject(project.OldProjectNumber);
                List<Component> newComponentList = new List<Component>();
                //List<Component> updatedComponentList = new List<Component>();
                List<TaskInfo> newTaskList = new List<TaskInfo>();
                List<Component> deletedComponentList = new List<Component>();

                if (project.ProjectNumberChanged && ProjectExists(project.ProjectNumber))
                {
                    MessageBox.Show("There is another project with that same project number. Enter a different project number.");
                    return false;
                }

                UpdateProjectData(project);

                if (ProjectHasComponents(project.ProjectNumber))
                {
                    // Check modified project for added components.
                    foreach (Component component in project.ComponentList)
                    {
                        component.SetPosition(project.ComponentList.IndexOf(component));

                        bool componentExists = databaseProject.ComponentList.Exists(x => x.Name == component.Name);

                        if (componentExists)
                        {
                            UpdateComponentData(project, component);

                            foreach (TaskInfo task in component.TaskList)
                            {
                                Console.WriteLine(task.ID);
                            }

                            UpdateTasks(project.JobNumber, project.ProjectNumber, component.Name, component.TaskList);

                            //if (component.ReloadTaskList)
                            //{
                            //    RemoveTasks(project, component);
                            //    newTaskList.AddRange(component.TaskList);
                            //}
                            //else
                            //{
                            //    UpdateTasks(project.JobNumber, project.ProjectNumber, component.Name, component.TaskList);
                            //}
                        }
                        else
                        {
                            newComponentList.Add(component);
                            newTaskList.AddRange(component.TaskList);
                        }
                    }

                    // Check modified project for deleted components.
                    foreach (Component component in databaseProject.ComponentList)
                    {
                        bool componentExists = project.ComponentList.Exists(x => x.Name == component.Name);

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
                        foreach (Component component in deletedComponentList)
                        {
                            RemoveComponent(project, component);
                        }
                    }
                }
                else
                {
                    AddComponentDataTableToDatabase(CreateDataTableFromComponentList(project));
                }

                MessageBox.Show("Project Updated!");

                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\n" + e.StackTrace);
                return false;
            }
        }

        private void UpdateProjectData(ProjectInfo project)
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

        public bool UpdateProjectsTable(object s, CellValueChangedEventArgs ev)
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

                using (OleDbConnection connection = new OleDbConnection(ConnString))
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = connection;

                    if (ev.Column.FieldName == "JobNumber")
                    {
                        cmd.CommandText = "UPDATE Projects SET JobNumber = @jobNumber " + whereClause;

                        cmd.Parameters.AddWithValue("@jobNumber", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "ProjectNumber")
                    {
                        if (!ProjectExists((int)ev.Value))
                        {
                            cmd.CommandText = "UPDATE Projects SET ProjectNumber = @projectNumber " + whereClause;

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
                    else if (ev.Column.FieldName == "KanBanWorkbookPath")
                    {
                        cmd.CommandText = "UPDATE Projects SET KanBanWorkbookPath = @kanBanWorkbookPath " + whereClause;

                        cmd.Parameters.AddWithValue("@kanBanWorkbookPath", ev.Value.ToString());
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
                throw e;
            }
        }

        public void SetKanBanWorkbookPath(string path, string jobNumber, int projectNumber)
        {
            using (OleDbConnection connection = new OleDbConnection(ConnString))
            {
                string queryString = "UPDATE Projects SET KanBanWorkbookPath = @path " +
                                     "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber";

                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

                adapter.UpdateCommand = new OleDbCommand(queryString, connection);

                adapter.UpdateCommand.Parameters.AddWithValue("@path", path);
                adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
                adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

                connection.Open();
                adapter.UpdateCommand.ExecuteNonQuery();
            }

        }

        #endregion

        #region Delete

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

        #endregion

        #endregion // Project Operations

        #region Components Table Operations

        #region Create

        public void AddComponents(ProjectInfo project)
        {
            OleDbCommand cmd;
            Component component;

            string queryString;

            queryString = "SELECT * FROM Components WHERE ProjectNumber = @projectNumber";

            cmd = new OleDbCommand(queryString, Connection);
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

        #endregion

        #region Read

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

        public List<Component> GetComponentListFromTasksTable(int projectNumber)
        {
            OleDbCommand cmd;
            List<Component> componentList = new List<Component>();

            string queryString;

            queryString = "SELECT DISTINCT Component FROM Tasks WHERE ProjectNumber = @projectNumber";
            OleDbConnection Connection = new OleDbConnection(ConnString);
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

        #endregion

        #region Update

        private void UpdateComponentData(ProjectInfo project, Component component)
        {
            try
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                string queryString;

                queryString = "UPDATE Components " +
                              "SET Component = @name, Notes = @notes, Priority = @priority, [Position] = @position, Quantity = @quantity, Spares = @spares, Pictures = @picture, Material = @material, Finish = @finish, TaskIDCount = @taskIDCount " +
                              "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @oldName";

                adapter.UpdateCommand = new OleDbCommand(queryString, Connection);

                adapter.UpdateCommand.Parameters.AddWithValue("@name", component.Name);
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

                Console.WriteLine($"{component.Name} Updated.");
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

        #endregion

        #region Delete

        private void RemoveComponent(ProjectInfo project, Component component)
        {
            var adapter = new OleDbDataAdapter();

            adapter.DeleteCommand = new OleDbCommand("DELETE FROM Components WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component", Connection);

            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = project.JobNumber;
            adapter.DeleteCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = project.ProjectNumber;
            adapter.DeleteCommand.Parameters.Add("@component", OleDbType.VarChar, 35).Value = component.Name;

            Connection.Open();
            adapter.DeleteCommand.ExecuteNonQuery();
            Connection.Close();
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

        #endregion

        private DataTable CreateDataTableFromComponentList(ProjectInfo project)
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

            foreach (Component component in project.ComponentList)
            {
                DataRow row = dt.NewRow();

                row["JobNumber"] = project.JobNumber;
                row["ProjectNumber"] = project.ProjectNumber;
                row["Component"] = component.Name;
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

        private DataTable CreateDataTableFromComponentList(ProjectInfo project, List<Component> componentList)
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

            foreach (Component component in componentList)
            {
                DataRow row = dt.NewRow();

                row["JobNumber"] = project.JobNumber;
                row["ProjectNumber"] = project.ProjectNumber;
                row["Component"] = component.Name;
                row["Notes"] = component.Notes;
                row["Priority"] = component.Priority;
                row["Position"] = component.Position;
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
                //Console.WriteLine(component.Name);

                foreach (TaskInfo task in component.TaskList)
                {
                    //Console.WriteLine($"   {task.ID} {task.TaskName}");
                    task.HasInfo = true;
                }
            }
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

        #endregion

        #region Read

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

        public List<TaskInfo> GetProjectTaskList(string jobNumber, int projectNumber)
        {
            OleDbCommand cmd;
            List<TaskInfo> taskList = new List<TaskInfo>();

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
                            taskList.Add(new TaskInfo
                            (
                                    taskName: rdr["TaskName"],
                                          id: rdr["TaskID"],
                                  databaseId: rdr["ID"],
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

        private DataTable CreateDataTableFromTaskList(ProjectInfo project)
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

            foreach (Component component in project.ComponentList)
            {
                i = 1;

                foreach (TaskInfo task in component.TaskList)
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
            using (OleDbConnection connection = new OleDbConnection(ConnString))
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

        private string SetWeeklyHoursQueryString(string weekStart, string weekEnd)
        {
            string department = "All";
            string queryString = null;
            string selectStatment = "JobNumber, ProjectNumber, TaskName, Duration, StartDate, FinishDate, Hours";
            //string fromStatement = "Tasks";
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
                        if (rdr["TaskName"].ToString() == "Program Rough")
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

        public List<Week> GetWeekHours(string weekStart, string weekEnd)
        {
            List<Week> weekList = new List<Week>();
            List<Week> deptWeekList = new List<Week>();
            Week weekTemp;
            DateTime wsDate = Convert.ToDateTime(weekStart);
            int weekNum;
            Stopwatch stopwatch = new Stopwatch();

            string queryString = SetWeeklyHoursQueryString(weekStart, weekEnd);
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbCommand cmd = new OleDbCommand(queryString, Connection);

            weekList = InitializeDeptWeeksList(wsDate);

            //Console.WriteLine("\nLoad");

            Connection.Open();

            stopwatch.Start();

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

                        deptWeekList = week.ToList();

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
                                        if (weekTemp.Department == "CNC Finish")
                                            Console.WriteLine($"{weekTemp.Department} {weekTemp.WeekStart.ToShortDateString()} {date.DayOfWeek} Daily AVG. {dailyAVG} Hrs {hours} Days {days}");
                                        days -= 1;
                                    }


                                    date = date.AddDays(1);
                                }
                            }
                            else
                            {
                                weekTemp.AddHoursToDay((int)date.AddDays(days).DayOfWeek, dailyAVG);
                                if (weekTemp.Department == "CNC Finish")
                                    Console.WriteLine($"{weekTemp.Department} {weekTemp.WeekStart.ToShortDateString()} {date.AddDays(days).DayOfWeek} {dailyAVG} {days}");
                            }
                        }
                    }
                }
                else
                {

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

            Connection.Close();
            Connection.Dispose();

            //Console.WriteLine("\nReview:");

            //foreach (Week week in weekList)
            //{
            //    Console.WriteLine($"{week.Department} {week.GetWeekHours()} {week.WeekStart.ToShortDateString()} - {week.WeekEnd.ToShortDateString()}");
            //}

            return weekList;
        }

        private string SetQueryString(string department)
        {
            string queryString = null;
            string selectStatment = "ID, JobNumber & ' #' & ProjectNumber & ' ' & Component As Subject, TaskName & ' (' & Hours & ' Hours)' As Location, JobNumber, ProjectNumber, TaskID, TaskName, Component, Hours, StartDate, FinishDate, Machine, Resources, Resource, Status, Notes";
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

        public DataTable GetAppointmentData(string department, bool grouped = false)
        {
            DataTable dt = new DataTable();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            //string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resources FROM Tasks WHERE TaskName LIKE 'CNC Finish'";
            string queryString = SetQueryString(department);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);
            //int i = 1;

            //adapter.SelectCommand.Parameters.AddWithValue("@department", setQueryString(department);

            try
            {
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

        public DataTable GetAppointmentData()
        {
            DataTable dt = new DataTable();

            try
            {
                using (OleDbConnection connection = new OleDbConnection(ConnString))
                {
                    //string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resources FROM Tasks WHERE TaskName LIKE 'CNC Finish'";
                    string queryString = "SELECT JobNumber & ' ' & Component & ' ' & TaskName As Subject, StartDate, FinishDate, Machine, Resource, ToolMaker, Notes FROM Tasks WHERE TaskName = 'CNC Rough'";

                    OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

                    adapter.Fill(dt);
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

        #endregion

        #region Update
        // This method is for capturing changes made to tasks in the department schedule view.
        public bool UpdateTask(string jobNumber, int projectNumber, string component, int taskID, DateTime startDate, DateTime finishDate, string machine, string resource, string resources)
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(ConnString))
                {
                    string queryString;

                    queryString = "UPDATE Tasks SET StartDate = @startDate, FinishDate = @finishDate, Machine = @machine, Resource = @resource, Resources = @resources " +
                                    "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";

                    OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

                    adapter.UpdateCommand = new OleDbCommand(queryString, connection);

                    string predecessors = GetTaskPredecessors(jobNumber, projectNumber, component, taskID);

                    if (predecessors != "" && GetLatestPredecessorFinishDate(jobNumber, projectNumber, component, predecessors) > startDate)
                    {
                        MessageBox.Show("You cannot put a task start date before its predecessor's finish date.");
                        return false;
                    }

                    adapter.UpdateCommand.Parameters.AddWithValue("@startDate", startDate);
                    adapter.UpdateCommand.Parameters.AddWithValue("@finishDate", finishDate);
                    adapter.UpdateCommand.Parameters.AddWithValue("@machine", machine);
                    adapter.UpdateCommand.Parameters.AddWithValue("@resource", resource);
                    adapter.UpdateCommand.Parameters.AddWithValue("@resources", resources);
                    adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
                    adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
                    adapter.UpdateCommand.Parameters.AddWithValue("@component", component);
                    adapter.UpdateCommand.Parameters.AddWithValue("@taskID", taskID);

                    connection.Open();
                    adapter.UpdateCommand.ExecuteNonQuery();
                }

                MoveDescendents(jobNumber, projectNumber, component, finishDate, taskID);
                return true;

            }
            catch (Exception er)
            {
                return false;
            } 
            
        }

        // string machine, string resource, 
        public bool UpdateTask(string jobNumber, int projectNumber, string component, int taskID, DateTime startDate, DateTime finishDate)
        {
            try
            {
                string queryString;

                queryString = "UPDATE Tasks SET StartDate = @startDate, FinishDate = @finishDate " +
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

        public bool UpdateTaskResource(string jobNumber, int projectNumber, string component, int taskID, string resource)
        {
            try
            {
                string queryString = "UPDATE Tasks SET Resource = @resource " +
                                     "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";
                OleDbConnection Connection = new OleDbConnection(ConnString);
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
            catch(Exception ex)
            {
                return false;
            }
            finally
            {
                Connection.Close();
                Connection.Dispose();
            }
        }

        private void UpdateTasks(string jobNumber, int projectNumber, string component, List<TaskInfo> taskList)
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

        private List<TaskInfo> RenumberTaskIDs(List<TaskInfo> taskList)
        {
            for (int i = 0; i < taskList.Count; i++)
            {
                taskList[i].SetTaskID(i + 1);
                Console.WriteLine(taskList[i].ID);
            }

            return taskList;
        }

        private DataTable UpdateTaskDataTable(string jobNumber, int projectNumber, string component, List<TaskInfo> taskList, DataTable taskDataTable)
        {
            UpdateTasksInDataTable(taskList, taskDataTable);
            DeleteTasksFromDatatable(taskList, taskDataTable);
            AddTasksToDataTable(jobNumber, projectNumber, component, taskList, taskDataTable);
            
            return taskDataTable;
        }

        // DO NOT CALL THIS METHOD AFTER ADDING TASKS TO DATATABLE.  NEW TASKS DO NOT HAVE IDS ASSIGNED TO THEM AND THEREFORE CANNOT BE READ.
        private DataTable DeleteTasksFromDatatable(List<TaskInfo> taskList, DataTable taskDataTable)
        {
            int id;
            foreach (DataRow row in taskDataTable.Rows)
            {
                id = (int)row["ID"];
                if (!taskList.Exists(x => x.DatabaseID == (int)row["ID"]))
                {
                    row.Delete();
                }
            }

            return taskDataTable;
        }

        private DataTable AddTasksToDataTable(string jobNumber, int projectNumber, string component, List<TaskInfo> taskList, DataTable taskDataTable)
        {
            var tasksToAdd = from t in taskList
                             where t.DatabaseID == 0
                             select t;

            foreach (TaskInfo task in tasksToAdd.ToList())
            {
                DataRow row = taskDataTable.NewRow();

                row["TaskID"] = task.ID;
                row["JobNumber"] = jobNumber;
                row["ProjectNumber"] = projectNumber;
                row["Component"] = component;
                row["TaskName"] = task.TaskName;
                row["Hours"] = task.Hours;
                row["Duration"] = task.Duration;
                row["Machine"] = task.Machine;
                row["Resource"] = task.Personnel;
                row["Predecessors"] = task.Predecessors;
                row["Priority"] = task.Priority;
                row["Notes"] = task.Notes;

                taskDataTable.Rows.Add(row);
            }

            return taskDataTable;
        }

        // DO NOT CALL THIS METHOD AFTER DELETING TASKS FROM DATATABLE. FOR SOME REASON IT THROWS AN ERROR WHEN SELECTING ROWS.
        private DataTable UpdateTasksInDataTable(List<TaskInfo> taskList, DataTable taskDataTable)
        {
            var tasksToUpdate = from t in taskList
                                where t.DatabaseID != 0
                                select t;

            foreach (TaskInfo task in tasksToUpdate.ToList())
            {
                var rows = from DataRow myRow in taskDataTable.Rows
                           where (int)myRow["ID"] == task.DatabaseID
                           select myRow;

                DataRow row = rows.FirstOrDefault();

                row["TaskID"] = task.ID;
                row["TaskName"] = task.TaskName;
                row["Hours"] = task.Hours;
                row["Duration"] = task.Duration;
                row["Machine"] = task.Machine;
                row["Resource"] = task.Personnel;
                row["Predecessors"] = task.Predecessors;
                row["Priority"] = task.Priority;
                row["Notes"] = task.Notes;
            }

            return taskDataTable;
        }

        // This means both machines and personnel.
        public void SetTaskResources(object s, CellValueChangedEventArgs ev, SchedulerStorage schedulerStorage)
        {
            using (OleDbConnection connection = new OleDbConnection(ConnString))
            {
                string taskName;

                var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

                DataTable dt = new DataTable();

                //string queryString = "UPDATE Tasks SET Resource = @resource " +
                //                     "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND TaskName = @taskName";

                string queryString = "SELECT * FROM Tasks WHERE ProjectNumber = @projectNumber AND TaskName = @taskName";

                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, connection);

                if (grid.Columns["MWONumber"] != null)
                {
                    if (grid.GetRowCellValue(ev.RowHandle, grid.Columns["MWONumber"]).ToString() != "")
                    {
                        adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = grid.GetRowCellValue(ev.RowHandle, grid.Columns["MWONumber"]);
                    }
                    else
                    {
                        adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = grid.GetRowCellValue(ev.RowHandle, grid.Columns["ProjectNumber"]);
                    }
                }
                else
                {
                    adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = grid.GetRowCellValue(ev.RowHandle, grid.Columns["ProjectNumber"]);
                }

                

                taskName = "Program " + ev.Column.FieldName.Remove(ev.Column.FieldName.Length - 10, 10);

                if (taskName.Contains("Electrode"))
                {
                    taskName = taskName + "s";
                }

                adapter.SelectCommand.Parameters.Add("@taskName", OleDbType.VarChar, 20).Value = taskName;

                OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

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

        private string GenerateResourceIDsString(string machine, string resource, SchedulerStorage schedulerStorage)
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

        public void ChangeTaskStartDate(string jobNumber, int projectNumber, string component, DateTime currentTaskStartDate, string duration, int taskID)
        {
            try
            {
                DateTime currentTaskFinishDate = AddBusinessDays(currentTaskStartDate, duration);

                using (OleDbConnection connection = new OleDbConnection(ConnString))
                {
                    OleDbDataAdapter adapter = new OleDbDataAdapter();

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
                }

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
        }

        public void ChangeTaskFinishDate(string jobNumber, int projectNumber, string component, DateTime currentTaskFinishDate, int taskID)
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(ConnString))
                {
                    OleDbDataAdapter adapter = new OleDbDataAdapter();

                    string queryString;

                    queryString = "UPDATE Tasks " +
                                  "SET FinishDate = @finishDate " +
                                  "WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID";

                    adapter.UpdateCommand = new OleDbCommand(queryString, connection);

                    adapter.UpdateCommand.Parameters.AddWithValue("@finishDate", currentTaskFinishDate);
                    adapter.UpdateCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
                    adapter.UpdateCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
                    adapter.UpdateCommand.Parameters.AddWithValue("@component", component);
                    adapter.UpdateCommand.Parameters.AddWithValue("@taskID", taskID);

                    connection.Open();

                    adapter.UpdateCommand.ExecuteNonQuery();
                }

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

        }

        public void MoveDescendents(string jobNumber, int projectNumber, string component, DateTime currentTaskFinishDate, int currentTaskID)
        {
            using (OleDbConnection connection = new OleDbConnection(ConnString))
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

        public void UpdateTasksTable(object s, CellValueChangedEventArgs ev, string resources = "")
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
                    if (ev.Column.FieldName == "TaskName")
                    {
                        cmd.CommandText = "UPDATE Tasks SET TaskName = @taskName WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@taskName", ev.Value.ToString());
                    }
                    else if (ev.Column.FieldName == "Notes")
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
                        //cmd.Parameters.AddWithValue("@resources", resources);
                    }
                    else if (ev.Column.FieldName == "Resource")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Resource = @resource WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@resource", ev.Value.ToString());
                        //cmd.Parameters.AddWithValue("@resources", resources);
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

        #endregion

        #region Delete

        public void RemoveTasks(ProjectInfo project, Component component)
        {
            var adapter = new OleDbDataAdapter();

            adapter.DeleteCommand = new OleDbCommand("DELETE FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component", Connection);

            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = project.JobNumber;
            adapter.DeleteCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = project.ProjectNumber;
            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 35).Value = component.Name;

            Connection.Open();
            adapter.DeleteCommand.ExecuteNonQuery();
            Connection.Close();
        }

        #endregion

        private DataTable CreateDataTableFromTaskList(ProjectInfo project, List<TaskInfo> taskList)
        {
            DataTable dt = new DataTable();
            string component = "";
            int i = 1;

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
            //dt.Columns.Add("Machine", typeof(string));
            //dt.Columns.Add("Resource", typeof(string));
            //dt.Columns.Add("Hours", typeof(int));
            //dt.Columns.Add("Priority", typeof(string));
            //dt.Columns.Add("Status", typeof(string));
            //dt.Columns.Add("DateAdded", typeof(DateTime));
            //dt.Columns.Add("Notes", typeof(string));
            //dt.Columns.Add("Initials", typeof(string));
            //dt.Columns.Add("DateCompleted", typeof(string));
            //dt.Columns.Add("ToolMaker", typeof(string));

            foreach (TaskInfo task in taskList)
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

        private int GetPercentComplete(string status)
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

        #endregion // Task Operations

        #region Resources Table Operations

        #region Create

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

        #endregion

        #region Read

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
                Console.WriteLine($"{nrow["ResourceName"]}");
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

        // This method is for populating the department schedule view with resources.
        public DataTable GetResourceData()
        {
            DataTable dt = new DataTable();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            string queryString = "SELECT ResourceName, ResourceType, Resources.ID, Role, Departments.Department From (Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID) LEFT OUTER JOIN Departments ON Roles.DepartmentID = Departments.ID ORDER BY ResourceName ASC";
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            DataRow row = dt.NewRow();

            row["ResourceName"] = "None";
            row["Role"] = "None";
            row["Department"] = "None";

            dt.Rows.Add(row);

            //foreach (DataRow nrow in dt.Rows)
            //{
            //    Console.WriteLine($"{nrow["ID"]} {nrow["ResourceName"]} {nrow["Role"]} {nrow["Department"]}");
            //}

            return dt;
        }

        #endregion

        #region Update

        #endregion

        #region Delete

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

        #endregion

        #endregion // Resources

        #region Roles Table Operations

        #region Create

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

        #endregion

        #region Read

        // Get ResourceList joins the resources and roles tables under the Read region under the Resources Table region.

        public DataTable GetRoleCounts()
        {
            string queryString = "SELECT COUNT(*) AS RoleCount, Role FROM Roles GROUP BY Role";

            DataTable dt = new DataTable();
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            return dt;
        }

        #endregion

        #region Update

        #endregion

        #region Delete

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

        #endregion

        #endregion

        #region Workload Table Operations

        #region Create

        public void AddWorkLoadEntry(WorkLoadInfo wli)
        {
            try
            {
                OleDbCommand cmd = new OleDbCommand("INSERT INTO WorkLoad (ToolNumber, MWONumber, ProjectNumber, Stage, Customer, PartName, DeliveryInWeeks, StartDate, FinishDate, AdjustedDeliveryDate, MoldCost, Engineer, Designer, ToolMaker, RoughProgrammer, FinishProgrammer, ElectrodeProgrammer, Manifold, MoldBase, GeneralNotes) VALUES " +
                                                                       "(@toolNumber, @mwoNumber, @projectNumber, @stage, @customer, @partName, @deliveryInWeeks, @startDate, @finishDate, @adjustedDeliveryDate, @moldCost, @engineer, @designer, @toolMaker, @roughProgrammer, @finishProgrammer, @electrodeProgrammer, @manifold, @moldBase, @generalNotes)", Connection);

                cmd.Parameters.AddWithValue("@toolNumber", wli.ToolNumber);
                cmd.Parameters.AddWithValue("@mwoNumber", wli.MWONumber);
                cmd.Parameters.AddWithValue("@projectNumber", wli.ProjectNumber);
                cmd.Parameters.AddWithValue("@stage", wli.Stage);
                cmd.Parameters.AddWithValue("@customer", wli.Customer);
                cmd.Parameters.AddWithValue("@partName", wli.PartName);
                cmd.Parameters.AddWithValue("@deliveryInWeeks", wli.DeliveryInWeeks);

                if (wli.StartDate != null)
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
                cmd.Parameters.AddWithValue("@finishProgrammer", wli.FinisherProgrammer);
                cmd.Parameters.AddWithValue("@electrodeProgrammer", wli.ElectrodeProgrammer);
                cmd.Parameters.AddWithValue("@manifold", wli.Manifold);
                cmd.Parameters.AddWithValue("@moldBase", wli.MoldBase);
                cmd.Parameters.AddWithValue("@generalNotes", wli.GeneralNotes);

                Connection.Open();
                cmd.ExecuteNonQuery();
            }
            finally
            {
                Connection.Close();
            }
        }

        #endregion

        #region Read

        #endregion

        #region Update

        public void UpdateWorkloadTable(object s, CellValueChangedEventArgs ev)
        {
            var grid = (s as DevExpress.XtraGrid.Views.Grid.GridView);

            //queryString = "UPDATE Tasks SET JobNumber = @jobNumber, Component = @component, TaskID = @taskID, TaskName = @taskName, " +
            //    "Duration = @duration, StartDate = @startDate, FinishDate = @finishDate, Predecessor = @predecessor, Machines = @machines, " +
            //    "Machine = @machine, Person = @person, Priority = @priority WHERE ID = @tID";

            using (OleDbConnection connection = new OleDbConnection(ConnString))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandType = CommandType.Text;
                cmd.Connection = connection;

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
                connection.Open();
                cmd.ExecuteNonQuery();
            }
        }

        #endregion

        #region Delete

        public bool DeleteWorkLoadEntry(int id)
        {
            OleDbCommand cmd = new OleDbCommand("DELETE FROM WorkLoad WHERE ID = @id", Connection);

            cmd.Parameters.AddWithValue("@id", id);

            Connection.Open();
            cmd.ExecuteNonQuery();
            Connection.Close();

            return true;
        }

        #endregion

        #endregion // WorkLoad

        #region WorkloadColors Table Operations

        #region Create

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

        #endregion

        #region Read

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

        #endregion

        #region Update

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

        #endregion

        #region Delete

        public void DeleteColorEntries(int projectID)
        {
            OleDbCommand cmd = new OleDbCommand("DELETE FROM WorkLoadColors WHERE ProjectID = @projectID", Connection);

            cmd.Parameters.AddWithValue("@projectID", projectID);

            Connection.Open();
            cmd.ExecuteNonQuery();
            Connection.Close();
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

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.AddWithValue("@machineType", "%" + machineType + "%");

            adapter.Fill(dt);

            foreach (DataRow nrow in dt.Rows)
            {
                machineList.Add($"{nrow["MachineName"]}");
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
            OleDbConnection Connection = new OleDbConnection(ConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            return dt;
        }

        #endregion

        #region Update

        #endregion

        #region Delete

        // TODO: Create method for deleting departments from the database.

        #endregion

        #endregion // Departments Table Operations
        
        public double GetBusinessDays(DateTime startD, DateTime endD)
        {
            double calcBusinessDays =
                1 + ((endD - startD).TotalDays * 5 -
                (startD.DayOfWeek - endD.DayOfWeek) * 2) / 7;

            if (endD.DayOfWeek == DayOfWeek.Saturday) calcBusinessDays--;
            if (startD.DayOfWeek == DayOfWeek.Sunday) calcBusinessDays--;

            return calcBusinessDays;
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
    }
}
