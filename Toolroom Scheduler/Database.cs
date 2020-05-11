using ClassLibrary;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Toolroom_Scheduler
{
    class Database
    {
        static string ConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=X:\TOOLROOM\Workload Tracking System\Database\Workload Tracking System DB.accdb";
        OleDbConnection Connection = new OleDbConnection(ConnString);
		private string component, toolMaker, dateTime;
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Stopwatch stopWatch, stopWatch2;

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
                date = date.AddDays(1);
                days -= 1;
            }
            else if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = date.AddDays(2);
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

        public bool LoadProjectToDB(ProjectModel project)
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
                if(AddProjectDataToDatabase(project) && 
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

        public bool EditProjectInDB(ProjectModel project)
        {
            try
            {
                ProjectModel databaseProject = GetProject(project.ProjectNumber);
                List<ComponentModel> newComponentList = new List<ComponentModel>();
                //List<Component> updatedComponentList = new List<Component>();
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

        public void LoadProjectToDatabase(ProjectModel project)
        {
            if (ProjectExists(project.ProjectNumber))
            {
                UpdateProjectData(project);
            }
            else
            {
                AddProjectDataToDatabase(project);
            }
        }


        //public bool KanBanExists(string jobNumber, int projectNumber)
        //{
        //    getKanBanWorkbookPath(jobNumber, projectNumber);
        //}

		public List<string> GetResourceList(string role)
		{
            List<string> ResourceList = new List<string>();
            DataTable dt = new DataTable();

            string queryString3 = "SELECT DISTINCT Resources.ResourceName From Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID WHERE Role = @role OR Role LIKE @role ORDER BY Resources.ResourceName ASC";

            Stopwatch sw = new Stopwatch();
            sw.Start();

            OleDbDataAdapter adapter3 = new OleDbDataAdapter(queryString3, Connection);
			
			adapter3.SelectCommand.Parameters.AddWithValue("@role", "%" + role + "%");
			
			adapter3.Fill(dt);

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

            string queryString3 = "SELECT * From Resources ORDER BY ResourceName ASC";

            OleDbDataAdapter adapter3 = new OleDbDataAdapter(queryString3, Connection);
            
            adapter3.Fill(dt);

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
            string queryString3 = "SELECT ID FROM Resources WHERE ResourceName = @resourceName";

            OleDbCommand sqlCommand = new OleDbCommand(queryString3, Connection);

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

            string queryString3 = "SELECT Resources.ResourceName, Roles.Role FROM Resources INNER JOIN Roles ON Resources.ID = Roles.ResourceID WHERE Roles.Role = @role ORDER BY Resources.ResourceName ASC";

            OleDbDataAdapter adapter3 = new OleDbDataAdapter(queryString3, Connection);

            adapter3.SelectCommand.Parameters.AddWithValue("@role", role);

            adapter3.Fill(dt);

            foreach (DataRow nrow in dt.Rows)
            {
                RoleList.Add(nrow["ResourceName"].ToString());
                //Console.WriteLine($"Added: {nrow["FirstName"]} {nrow["LastName"]} {nrow["Role"]}");
            }

            return RoleList;
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

        public ProjectModel GetProject(string jobNumber, int projectNumber)
        {
            ProjectModel project = GetProjectInfo(jobNumber, projectNumber);

            AddComponents(project);

            AddTasks(project);

            return project;
        }

        public ProjectModel GetProject(int projectNumber)
        {
            ProjectModel project = GetProjectInfo(projectNumber);

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
                                       apprentice: Convert.ToString(rdr["Apprentice"]),
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
                                       apprentice: Convert.ToString(rdr["Apprentice"]),
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
                        project.AddComponentList(GetComponentListFromTasksTable(project.JobNumber, project.ProjectNumber));
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
        }

        public List<string> GetComponentList(string jobNumber, int projectNumber)
        {
            List<string> componentList = new List<string>();

            string queryString = "SELECT * FROM Components WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber";

            OleDbCommand cmd = new OleDbCommand(queryString, Connection);

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
                            componentList.Add(rdr["Component"].ToString());
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
            finally
            {
                Connection.Close();
            }

            return componentList;
        }

        public List<ComponentModel> GetComponentListFromTasksTable(string jobNumber, int projectNumber)
        {
            OleDbCommand cmd;
            List<ComponentModel> componentList = new List<ComponentModel>();

            string queryString;

            queryString = "SELECT DISTINCT Component FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber";

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
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Connection.Close();
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
                            orderby t.TaskID ascending
                            select t;

                component.AddTaskList(tasks.ToList());
            }

            // This assumes that all tasks in the database have task info associated with them.
            // This can be assumed because all tasks before the project is created are checked
            // to see if they have task info.  If they do not project creation is disallowed.

            foreach (ComponentModel component in project.Components)
            {
                foreach (TaskModel task in component.Tasks)
                {
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
            catch (Exception)
            {
                throw;
            }
            finally
            {
                Connection.Close();
            }

            return taskList;
        }

        public bool ProjectHasDates(DataTable dt)
        {
            foreach (DataRow nrow in dt.Rows)
            {
                if(nrow["StartDate"] != DBNull.Value || nrow["FinishDate"] != DBNull.Value)
                {
                    return true;
                }
            }

            return false;
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

            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber ORDER BY ID DESC";

            adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
            adapter.SelectCommand.Parameters.Add("@projectNumber", OleDbType.Integer, 12).Value = projectNumber;

            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter); // This is needed in order for update command to work for some reason.

            adapter.Fill(dt);

            if(componentList == null)
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

                if(componentList == null)
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
                if(nrow["StartDate"] == DBNull.Value || Convert.ToDateTime(nrow["StartDate"]) < predecessorFinishDate)
                {
                    if(skipDatedTasks == true && (nrow["StartDate"] != DBNull.Value || nrow["FinishDate"] != DBNull.Value))
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

        public void BackDateProjectTasks(string jobNumber, int projectNumber, List<string> componentList, DateTime backDate)
        {
            OleDbDataAdapter adapter;
            DataTable dt = new DataTable();
            string queryString;
            bool skipDatedTasks = false;

            if(backDate == new DateTime(2000, 1, 1))
            {
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
                if(componentList == null)
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
                if(skipDatedTasks == true && (nrow["FinishDate"] != DBNull.Value || nrow["StartDate"] != DBNull.Value))
                {
                    goto Skip;
                }

                nrow["FinishDate"] = descendantStartDate;
                //MessageBox.Show(nrow["TaskName"].ToString());
                nrow["StartDate"] = SubtractBusinessDays(Convert.ToDateTime(nrow["FinishDate"]), nrow["Duration"].ToString());

                Skip:;

                // If a task has more than one predecessor.
                // Backdate each predecessor.
                if(nrow["Predecessors"].ToString().Contains(','))
                {
                    predecessors = nrow["Predecessors"].ToString().Split(',');

                    foreach(string id in predecessors)
                    {
                        BackDateTask(Convert.ToInt32(id), component, skipDatedTasks, Convert.ToDateTime(nrow["StartDate"]), projectTaskTable);
                    }
                }
                // If a task has one predecessor.
                // Backdate the one predecessor.
                else if(nrow["Predecessors"].ToString() != "")
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

		private Boolean SheetNExists(int sheetn)
		{
			foreach (Excel.Worksheet sheet in wb.Sheets)
			{
				if (sheet.Index == sheetn)
				{
					return true;
				}
			}

			return false;
		}

		private Boolean SheetNExists(string sheetname)
		{
			foreach (Excel.Worksheet sheet in wb.Sheets)
			{
				if (sheet.Name == sheetname)
				{
					return true;
				}
			}

			return false;
		}

        public void UpdateDatabase(object s, DataGridViewCellEventArgs ev)
        {
            try
            {
                var grid = (s as DataGridView);

                //queryString = "UPDATE Tasks SET JobNumber = @jobNumber, Component = @component, TaskID = @taskID, TaskName = @taskName, " +
                //    "Duration = @duration, StartDate = @startDate, FinishDate = @finishDate, Predecessor = @predecessor, Machines = @machines, " +
                //    "Machine = @machine, Person = @person, Priority = @priority WHERE ID = @tID";

                using (Connection)
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = Connection;

                    if (grid.Columns[ev.ColumnIndex].Name == "TaskName")
                    {
                        cmd.CommandText = "UPDATE Tasks SET TaskName = @taskName WHERE (ID = @tID)";

                        if ((grid.Rows[ev.RowIndex]).Cells[ev.ColumnIndex].Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@taskName", (grid.Rows[ev.RowIndex]).Cells["TaskName"].Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@taskName", " ");
                        }
                    }
                    else if (grid.Columns[ev.ColumnIndex].Name == "Hours")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Hours = @hours WHERE (ID = @tID)";

                        if ((grid.Rows[ev.RowIndex]).Cells["Hours"].Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@hours", (grid.Rows[ev.RowIndex]).Cells["Hours"].Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@hours", DBNull.Value);
                        }
                    }
                    else if (grid.Columns[ev.ColumnIndex].Name == "Duration")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Duration = @duration WHERE (ID = @tID)";

                        if ((grid.Rows[ev.RowIndex]).Cells["Duration"].Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@duration", (grid.Rows[ev.RowIndex]).Cells["Duration"].Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@duration", "");
                        }
                    }
                    else if (grid.Columns[ev.ColumnIndex].Name == "StartDate")
                    {
                        cmd.CommandText = "UPDATE Tasks SET StartDate = @startDate WHERE (ID = @tID)";

                        if ((grid.Rows[ev.RowIndex]).Cells["StartDate"].Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@startDate", (grid.Rows[ev.RowIndex]).Cells["StartDate"].Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@startDate", DBNull.Value);
                        }
                    }
                    else if (grid.Columns[ev.ColumnIndex].Name == "FinishDate")
                    {
                        cmd.CommandText = "UPDATE Tasks SET FinishDate = @finishDate WHERE (ID = @tID)";

                        if ((grid.Rows[ev.RowIndex]).Cells["FinishDate"].Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@finishDate", (grid.Rows[ev.RowIndex]).Cells["FinishDate"].Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@finishDate", DBNull.Value);
                        }
                    }
                    else if (grid.Columns[ev.ColumnIndex].Name == "Predecessors")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Predecessors = @predecessors WHERE (ID = @tID)";

                        cmd.Parameters.AddWithValue("@predecessors", (grid.Rows[ev.RowIndex]).Cells["Predecessors"].Value.ToString());
                    }
                    else if (grid.Columns[ev.ColumnIndex].Name == "Resource")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Resource = @resource WHERE (ID = @tID)";

                        if ((grid.Rows[ev.RowIndex]).Cells["Resource"].Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@resource", (grid.Rows[ev.RowIndex]).Cells["Resource"].Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@resource", "");
                        }
                    }
                    else if (grid.Columns[ev.ColumnIndex].Name == "Machine")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Machine = @machine WHERE (ID = @tID)";

                        if ((grid.Rows[ev.RowIndex]).Cells["Machine"].Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@machine", (grid.Rows[ev.RowIndex]).Cells["Machine"].Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@machine", "");
                        }
                    }
                    else if (grid.Columns[ev.ColumnIndex].Name == "Priority")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Priority = @priority WHERE (ID = @tID)";

                        if ((grid.Rows[ev.RowIndex]).Cells["Priority"].Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@priority", (grid.Rows[ev.RowIndex]).Cells["Priority"].Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@priority", "");
                        }
                    }
                    else if (grid.Columns[ev.ColumnIndex].Name == "Status")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Status = @status WHERE (ID = @tID)";

                        if ((grid.Rows[ev.RowIndex]).Cells["Status"].Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@status", (grid.Rows[ev.RowIndex]).Cells["Status"].Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@status", "");
                        }
                    }
                    else if (grid.Columns[ev.ColumnIndex].Name == "Notes")
                    {
                        cmd.CommandText = "UPDATE Tasks SET Notes = @notes WHERE (ID = @tID)";

                        if ((grid.Rows[ev.RowIndex]).Cells["Notes"].Value.ToString() != "")
                        {
                            cmd.Parameters.AddWithValue("@notes", (grid.Rows[ev.RowIndex]).Cells["Notes"].Value.ToString());
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@notes", "");
                        }
                    }

                    cmd.Parameters.AddWithValue("@tID", (grid.Rows[ev.RowIndex]).Cells["ID"].Value.ToString());

                    //Console.WriteLine(connectionString);
                    //Console.WriteLine(queryString);
                    //Console.WriteLine((grid.Rows[ev.RowIndex]).Cells[0].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[1].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[2].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[3].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[4].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[5].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[6].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[7].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[8].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[9].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[10].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[11].Value.ToString() + " " + (grid.Rows[ev.RowIndex]).Cells[12].Value.ToString() + " ");
                    Connection.Open();

                    cmd.ExecuteNonQuery();
                }
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

        public void BulkAssignRoles(string jobNumber, string roughProgrammer, string finishProgrammer, string electrodeProgrammer)
        {
            var adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;

            if (jobNumber == "All")
            {
                MessageBox.Show("Just select a single job for now.");
                return;
            }

            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber ORDER BY TaskID ASC";
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 12).Value = jobNumber;
            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
            adapter.Fill(datatable);

            foreach (DataRow nrow in datatable.Rows)
            {

                if(nrow["TaskName"].ToString() == "Program Rough")
                {
                    nrow["Resource"] = roughProgrammer;
                }
                else if (nrow["TaskName"].ToString() == "Program Finish")
                {
                    nrow["Resource"] = finishProgrammer;
                }
                else if (nrow["TaskName"].ToString() == "Program Electrodes")
                {
                    nrow["Resource"] = electrodeProgrammer;
                }

            }
            
            adapter.UpdateCommand = builder.GetUpdateCommand();
            adapter.Update(datatable);
            MessageBox.Show("Done!");
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

        // Only need to delete the project from projects since the Database is set to cascade delete related records.
        public void RemoveProject(string jobNumber, int projectNumber)
        {
            var adapter = new OleDbDataAdapter();

            adapter.DeleteCommand = new OleDbCommand("DELETE FROM Projects WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", Connection);
            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
            adapter.DeleteCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = projectNumber;

            Connection.Open();
            adapter.DeleteCommand.ExecuteNonQuery();
            Connection.Close();
        }

        public void RemoveComponents(string jobNumber, int projectNumber)
        {
            var adapter = new OleDbDataAdapter();

            adapter.DeleteCommand = new OleDbCommand("DELETE FROM Components WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", Connection);
            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
            adapter.DeleteCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = projectNumber;

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

        public void RemoveTasks(string jobNumber, int projectNumber)
        {
            var adapter = new OleDbDataAdapter();

            adapter.DeleteCommand = new OleDbCommand("DELETE FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", Connection);

            adapter.DeleteCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 25).Value = jobNumber;
            adapter.DeleteCommand.Parameters.Add("@projectNumber", OleDbType.VarChar, 12).Value = projectNumber;

            Connection.Open();
            adapter.DeleteCommand.ExecuteNonQuery();
            Connection.Close();
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

        public void CycleThroughProjects()
        {
            //try
            //{
                stopWatch = new Stopwatch();
                string queryString = "SELECT * FROM Projects ORDER BY ID";
                OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                //DataView view = new DataView(dt);
                //DataTable distinctJobNumbers = view.ToTable(true, "JobNumber", "ProjectNumber");

                stopWatch.Start();

                foreach (DataRow nrow in dt.Rows) // Need to check if project table projects has all projects before switching to using it for this.  Also need option for if the Kan Ban filepath is in the database.
                {
                    Console.WriteLine($"{nrow["JobNumber"]} {nrow["ProjectNumber"]} {stopWatch.Elapsed.ToString("mm\\:ss\\.ff")}");

                    if(nrow["KanBanWorkbookPath"].ToString() == "")
                    {
                        FindKanBanSheet(nrow["JobNumber"].ToString(), Convert.ToInt32(nrow["ProjectNumber"]));
                    }
                    else
                    {
                        LoadProjectStatusesToDB(nrow["JobNumber"].ToString(), Convert.ToInt32(nrow["ProjectNumber"]), OpenAndReadKanBanSheet(nrow["KanBanWorkbookPath"].ToString(), nrow["JobNumber"].ToString(), Convert.ToInt32(nrow["ProjectNumber"])));
                    }
                    
                }

                MessageBox.Show("Done! \r\n\r\nTime Elapsed: " + stopWatch.Elapsed.ToString("mm\\:ss\\.ff"));
                //stopWatch2.Stop();
                stopWatch.Stop();
            //}
            //catch(Exception e)
            //{
            //    MessageBox.Show(e.Message);
            //}

        }

        public void FindKanBanSheet(string jn, int pn)
        {
            string toolYearFolderDirectory = "";
            string[] rootFolderEntries = Directory.GetDirectories(@"X:\TOOLROOM");
            //stopWatch2 = new Stopwatch();
            string jnFull = jn;

            if (jn.Length >= 6)
            {
                if (!int.TryParse(jn.Substring(0, 6), out int n))
                {
                    //MessageBox.Show("MWO with no job number found.");
                    goto MyEnd;
                    //FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                    //folderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer;
                    //folderBrowserDialog.SelectedPath = @"X:\TOOLROOM\";
                    //if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                    //{
                    //    //textBox2.Text = folderBrowserDialog.SelectedPath;
                    //}
                }

                jn = jn.Substring(0, 6);
            }
            else
            {
                if (!int.TryParse(jn.Substring(0, jn.Length), out int n))
                {
                    goto MyEnd;
                }
            }

            //stopWatch2.Start();

            foreach (string subdirectory in rootFolderEntries)
            {
                //Console.WriteLine(subdirectory); // The subdirectory includes the filepath.
                if (subdirectory.Substring(12, 2) == jn.Substring(0, 2))
                {
                    Console.WriteLine(subdirectory);
                    toolYearFolderDirectory = subdirectory;
                    break;
                }
            }

            string[] toolYearFolderEntries = Directory.GetDirectories(toolYearFolderDirectory);

            foreach (string subdirectory in toolYearFolderEntries)
            {
                //Console.WriteLine(subdirectory); // The subdirectory includes the filepath.
                if (subdirectory.Contains(jn))
                {
                    string[] fileEntries = Directory.GetFiles(subdirectory);

                    foreach (string fileEntry in fileEntries)
                    {
                        //Console.WriteLine(subdirectory);
                        if (fileEntry.Contains(".xlsx") && !fileEntry.Contains("~") && fileEntry.Contains("#" + pn))
                        {
                            Console.WriteLine(fileEntry);
                            LoadProjectStatusesToDB(jnFull,pn,OpenAndReadKanBanSheet(fileEntry,jn,pn));
                            goto MyEnd;
                        }
                    }

                    string[] toolFolderEntries = Directory.GetDirectories(subdirectory);

                    foreach (string subdirectory2 in toolFolderEntries)
                    {
                        string[] fileEntries2 = Directory.GetFiles(subdirectory2);

                        foreach (string fileEntry in fileEntries2)
                        {
                            //Console.WriteLine(subdirectory2);
                            if (fileEntry.Contains(".xlsx") && !fileEntry.Contains("~") && fileEntry.Contains("#" + pn))
                            {
                                Console.WriteLine(fileEntry);
                                LoadProjectStatusesToDB(jnFull, pn, OpenAndReadKanBanSheet(fileEntry, jn, pn));
                                goto MyEnd;
                            }
                        }
                    }
                }
            }
            MessageBox.Show($"Kan Ban with JobNumber: {jnFull} Project #: {pn} not found.");
            MyEnd:;
        }

        public void IterateThroughKanBanSheets(ProjectModel pi)
        {
            FileInfo file;
            string[] fileNameArr;
            string fileName;
            stopWatch = new Stopwatch();
            stopWatch2 = new Stopwatch();

            stopWatch.Start();

            //string blankSnapshotPath;
            //blankSnapshot = excelApp.Workbooks.Open(blankSnapshotPath);
            string[] fileEntries = Directory.GetFiles("Fix me later");
            foreach (string fileEntry in fileEntries)
            {
                if (fileEntry.Contains(".xlsx") && !fileEntry.Contains("~") && fileEntry.Contains("#" + pi.ProjectNumber))
                {
                    file = new FileInfo(fileEntry);
                    fileNameArr = fileEntry.Split('\\');
                    fileName = fileNameArr[fileNameArr.Length - 1].ToString();

                    stopWatch2.Start();
                     //test(fileEntry);

                    Console.WriteLine($"            File: {fileName}");
                    Console.WriteLine($"         Project: ");
                    Console.WriteLine($"Interaction Time: {stopWatch2.Elapsed.ToString("mm\\:ss\\.ff")}");
                    Console.WriteLine($"    Overall Time: {stopWatch.Elapsed.ToString("mm\\:ss\\.ff")}");
                    Console.WriteLine($" ");
                    stopWatch2.Reset();

                }

                //blankSnapshot.SaveCopyAs(targetSnapshotFolder + fileName);
            }
            MessageBox.Show("Finished!");
            stopWatch.Stop();
            stopWatch2.Stop();
        }

        public List<TaskModel> OpenAndReadKanBanSheet(string filepath, string jn, int pn)
        {
            int r;
            List<TaskModel> taskInfoList = new List<TaskModel>();
            Excel.Workbook wb;
            //Excel.Worksheet ws;
            
            Excel.Application excelApp = new Excel.Application();
            
            wb = excelApp.Workbooks.Open($"{filepath}", ReadOnly:true);

            foreach (Excel.Worksheet ws in wb.Sheets)
            {
                
                if(ws.Index > 1 && ws.Cells[2,2].value != null && ws.Name != "Mold") // Sheet 2 is a mold task sheet with just the design and create drawings task for the entire mold.
                {
                    r = 2;

                    //Console.WriteLine(" ");
                    //Console.WriteLine($"{ws.Cells[2, 2].Value.ToString()} {stopWatch2.Elapsed.ToString("mm\\:ss\\.ff")}");

                    do
                    {
                        if(ws.Cells[r, 9].Value != null)
                        {
                            //Console.WriteLine($"   {Convert.ToInt16(ws.Cells[r, 3].Value)} {ws.Cells[r, 4].Value.ToString().Trim(' ')} Completed"); // Shows what completed tasks were found.
                            taskInfoList.Add(new TaskModel(jn, pn, ws.Cells[2, 2].Value.ToString().Trim(' '), ws.Cells[r, 4].Value.ToString().Trim(' '), Convert.ToInt16(ws.Cells[r, 3].Value), "Completed" ));
                        }

                        r++;
                    } while (ws.Cells[r, 2].Value != null);
                }                                   
            }

            wb.Close(false);
            Console.WriteLine("");
            Console.WriteLine($"   Workbook Closed. {stopWatch.Elapsed.ToString("mm\\:ss\\.ff")}");
            excelApp.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(excelApp);
            Marshal.ReleaseComObject(wb);
            //Marshal.ReleaseComObject(ws);
            return taskInfoList;
        }

        public void OpenKanBanWorkbook(string filepath)
        {
            if(filepath != null && filepath != "")
            {
                FileInfo fi = new FileInfo(filepath);

                if (fi.Exists)
                {
                    //var attributes = File.GetAttributes(fi.FullName);    

                    var res = Process.Start("EXCEL.EXE", "/r \"" + fi.FullName + "\"");
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

        private void LoadProjectStatusesToDB(string jn, int pn, List<TaskModel> tia)
        {
            var adapter = new OleDbDataAdapter();
            DataTable datatable = new DataTable();
            string queryString;
            int i = 0;

            if(tia == null)
            {
                return;
            }

            //List<int> taskIDs = tia.Select(t => t.TaskID).ToList();
            List<TaskModel> tasksTemp = tia.ToList();
            //Console.WriteLine($"   Loading: {jn} {pn} {tia.Count}");
            //Console.WriteLine($"   ");
            queryString = "SELECT * FROM Tasks WHERE JobNumber = @jobNumber AND PartNumber = @partNumber ORDER BY TaskID ASC";
            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            adapter.SelectCommand.Parameters.Add("@jobNumber", OleDbType.VarChar, 12).Value = jn;
            adapter.SelectCommand.Parameters.Add("@partNumber", OleDbType.VarChar, 12).Value = pn;
            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
            adapter.Fill(datatable);

            foreach (DataRow nrow in datatable.Rows)
            {
                foreach(TaskModel ti in tia)
                {
                    if (Convert.ToInt16(nrow["TaskID"]) == ti.TaskID && nrow["Component"].ToString() == ti.Component)
                    {
                        i++;
                        //Console.WriteLine($"   Loaded: {i} {ti.Component} {ti.TaskName}");
                        nrow["Status"] = ti.Status;
                        foreach (var task in tia)
                        {
                            if (task.TaskID == ti.TaskID && task.Component == ti.Component)
                                tasksTemp.Remove(task);
                        }
                        break;
                    }
                     else
                    {
                        //Console.WriteLine($"   Failed: {ti.TaskID} {ti.Component} {ti.TaskName} {ti.Status}");
                    }
                }
                
            }

            adapter.UpdateCommand = builder.GetUpdateCommand();
            adapter.Update(datatable);
            if(i < tia.Count)
            {
                Console.WriteLine(" ");
                foreach (var task in tasksTemp)
                {
                    Console.WriteLine($"   Missed: {task.TaskID} {task.Component} {task.TaskName}");
                }
                
            }
            Console.WriteLine($" ");
            Console.WriteLine($"   Loaded: {i}/{tia.Count} {jn} {pn}");
            Console.WriteLine($" ");
            //MessageBox.Show("Done!");
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
                if(component.GetPictureByteArray() != null)
                {
                    row["Pictures"] = component.GetPictureByteArray();
                }
                else
                {
                    row["Pictures"] = DBNull.Value;
                }
                row["Notes"] = component.Notes;
                row["Position"] = position++;
                row["TaskIDCount"] = component.TaskIDCount;

                dt.Rows.Add(row);
            }

            foreach(DataRow nrow in dt.Rows)
            {
                Console.WriteLine($"{nrow["JobNumber"]} {nrow["ProjectNumber"]} {nrow["Component"]} {nrow["Position"]}");
            }

            Console.WriteLine("Component DataTable Created.");

            return dt;
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

                if(component != task.Component)
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
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Connection.Close();
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

        private bool ProjectExists(int projectNumber)
        {
            OleDbCommand sqlCommand = new OleDbCommand("SELECT COUNT(*) from Projects WHERE ProjectNumber = @projectNumber", Connection);

            sqlCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

            Connection.Open();
            int projectCount = (int)sqlCommand.ExecuteScalar();
            Connection.Close();

            if(projectCount > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
            
        }

        public bool ProjectExists(string jobNumber, int projectNumber)
        {
            OleDbCommand sqlCommand = new OleDbCommand("SELECT COUNT(*) from Projects WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", Connection);

            sqlCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
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

        public bool ComponentExists(string jobNumber, int projectNumber, string component)
        {
            OleDbCommand sqlCommand = new OleDbCommand("SELECT COUNT(*) from Components WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component", Connection);

            sqlCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            sqlCommand.Parameters.AddWithValue("@projectNumber", projectNumber);
            sqlCommand.Parameters.AddWithValue("@component", component);

            Connection.Open();
            int componentCount = (int)sqlCommand.ExecuteScalar();
            Connection.Close();

            if (componentCount > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool ProjectTasksExist(string jn, string pn)
        {
            OleDbCommand sqlCommand = new OleDbCommand("SELECT COUNT(*) from Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", Connection);

            Connection.Open();
            sqlCommand.Parameters.AddWithValue("@jobNumber", jn);
            sqlCommand.Parameters.AddWithValue("@projectNumber", pn);
            int userCount = (int)sqlCommand.ExecuteScalar();
            Connection.Close();

            if (userCount > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public int GetHighestProjectTaskID(string jn, int pn)
        {
            OleDbCommand sqlCommand = new OleDbCommand("SELECT MAX(TaskID) from Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", Connection);

            Connection.Open();
            sqlCommand.Parameters.AddWithValue("@jobNumber", jn);
            sqlCommand.Parameters.AddWithValue("@projectNumber", pn);
            int highestTaskID = (int)sqlCommand.ExecuteScalar();
            Connection.Close();

            return highestTaskID;
        }

        public DateTime GetFinishDate(string jn, int pn, string component, int tID)
        {
            DateTime FinishDate = DateTime.Today;

            OleDbCommand sqlCommand = new OleDbCommand("SELECT FinishDate from Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber AND Component = @component AND TaskID = @taskID", Connection);

            sqlCommand.Parameters.AddWithValue("@jobNumber", jn);
            sqlCommand.Parameters.AddWithValue("@projectNumber", pn);
            sqlCommand.Parameters.AddWithValue("@component", component);
            sqlCommand.Parameters.AddWithValue("@taskID", tID);

            try
            {
                Connection.Open();

                FinishDate = (DateTime)sqlCommand.ExecuteScalar();
            }
            catch
            {
                MessageBox.Show("A predecessor has no finish date.");
            }

            Connection.Close();

            return FinishDate;
        }

        public List<string> GetJobNumberComboList()
        {
            string queryString = "SELECT DISTINCT JobNumber, ProjectNumber FROM Tasks";
            DataTable dt = new DataTable();
            List<string> jobNumberList = new List<string>();

            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, Connection);

            adapter.Fill(dt);

            jobNumberList.Add("All");

            foreach (DataRow nrow in dt.Rows)
            {
                //Console.WriteLine(nrow["JobNumber"]);
                //
                jobNumberList.Add($"{nrow["JobNumber"].ToString()} - #{nrow["ProjectNumber"].ToString()}");
            }

            return jobNumberList;
        }



        public DataTable GetProjectTasksTable(string jobNumber, int projectNumber)
        {
            string queryString;
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable dt = new DataTable();

            queryString = "SELECT ProjectNumber, JobNumber, Component, TaskID, TaskName, Duration, StartDate, FinishDate, Predecessors, ToolMaker, Status, Initials, DateCompleted " +
              "FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber ORDER BY TaskID ASC";

            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            adapter.SelectCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            adapter.SelectCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

            adapter.Fill(dt);

            return dt;
        }

        public DataTable GetComponentsTable(string jobNumber, int projectNumber)
        {
            string queryString;
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable dt = new DataTable();

            queryString = "SELECT DISTINCT Component FROM Tasks WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber ORDER BY Component ASC";

            adapter.SelectCommand = new OleDbCommand(queryString, Connection);
            adapter.SelectCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            adapter.SelectCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

            adapter.Fill(dt);

            return dt;
        }

        public string GetKanBanWorkbookPath(string jobNumber, int projectNumber)
        {
            string kanBanWorkbookPath = "";

            OleDbCommand sqlCommand = new OleDbCommand("SELECT KanBanWorkbookPath from Projects WHERE JobNumber = @jobNumber AND ProjectNumber = @projectNumber", Connection);

            sqlCommand.Parameters.AddWithValue("@jobNumber", jobNumber);
            sqlCommand.Parameters.AddWithValue("@projectNumber", projectNumber);

            try
            {
                Connection.Open();
                
                kanBanWorkbookPath = ConvertObjectToString(sqlCommand.ExecuteScalar());
            }
            catch(Exception ex)
            {
                Connection.Close();
                MessageBox.Show(ex.Message);
            }

            Connection.Close();

            return kanBanWorkbookPath;
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

                if(component.GetPictureByteArray() != null)
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

        private string ConvertObjectToString(object obj)
        {
            if (obj != null)
            {
                return obj.ToString();
            }
            else
            {
                return "";
            }
        }
    }
}
