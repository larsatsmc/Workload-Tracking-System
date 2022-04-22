using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Access = Microsoft.Office.Interop.Access;

namespace DatabaseBackupApp
{
    public class Database
    {
        static readonly string ConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=X:\TOOLROOM\Workload Tracking System\Database\Workload Tracking System DB.accdb";

        static readonly string ConnString2 = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\s-fs1-smdrv\mydocs$\Joshua.Meservey\Microsoft Access\Tool Room Scheduler Database\Workload Tracking System DB.accdb";

        static readonly string ConnBase = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";

        // Does not set primary keys.
        public static void CopyDatabase(string destinationPath, string sourcePath = "")
        {
            string queryString1a = $"SELECT * INTO Projects IN '{destinationPath}' FROM Projects";
            string queryString1b = $"SELECT * INTO Components IN '{destinationPath}' FROM Components";
            string queryString1c = $"SELECT * INTO Tasks IN '{destinationPath}' FROM Tasks";
            string queryString1d = $"SELECT * INTO Resources IN '{destinationPath}' FROM Resources";
            string queryString1e = $"SELECT * INTO Roles IN '{destinationPath}' FROM Roles";
            string queryString1f = $"SELECT * INTO Departments IN '{destinationPath}' FROM Departments";
            string queryString1g = $"SELECT * INTO WorkLoad IN '{destinationPath}' FROM WorkLoad";
            string queryString1h = $"SELECT * INTO WorkLoadColors IN '{destinationPath}' FROM WorkLoadColors";

            using (OleDbConnection connection1 = new OleDbConnection(ConnString))
            {
                OleDbCommand cmd1 = new OleDbCommand(queryString1a, connection1);
                OleDbCommand cmd2 = new OleDbCommand(queryString1b, connection1);
                OleDbCommand cmd3 = new OleDbCommand(queryString1c, connection1);
                OleDbCommand cmd4 = new OleDbCommand(queryString1d, connection1);
                OleDbCommand cmd5 = new OleDbCommand(queryString1e, connection1);
                OleDbCommand cmd6 = new OleDbCommand(queryString1f, connection1);
                OleDbCommand cmd7 = new OleDbCommand(queryString1g, connection1);
                OleDbCommand cmd8 = new OleDbCommand(queryString1h, connection1);

                connection1.Open();

                cmd1.ExecuteNonQuery();
                cmd2.ExecuteNonQuery();
                cmd3.ExecuteNonQuery();
                cmd4.ExecuteNonQuery();
                cmd5.ExecuteNonQuery();
                cmd6.ExecuteNonQuery();
                cmd7.ExecuteNonQuery();
                cmd8.ExecuteNonQuery();
            }

            string queryString2a = $"ALTER TABLE Projects ADD PRIMARY KEY(JobNumber, ProjectNumber)";
            string queryString2b = $"ALTER TABLE Components ADD PRIMARY KEY(JobNumber, ProjectNumber, Component)";
            string queryString2c = $"ALTER TABLE Tasks ADD PRIMARY KEY(JobNumber, ProjectNumber, Component, TaskID)";
            string queryString2d = $"ALTER TABLE Resources ADD PRIMARY KEY(ID)";
            string queryString2e = $"ALTER TABLE Roles ADD PRIMARY KEY(ID)";
            string queryString2f = $"ALTER TABLE Departments ADD PRIMARY KEY(ID)";
            string queryString2g = $"ALTER TABLE WorkLoad ADD PRIMARY KEY(ID)";
            string queryString2h = $"ALTER TABLE WorkLoadColors ADD PRIMARY KEY(ID)";
            string queryString2i = $"ALTER TABLE WorkLoadColors ADD FOREIGN KEY(ProjectID) REFERENCES WorkLoad(ID)";
            string queryString2j = $"ALTER TABLE Components ADD FOREIGN KEY(JobNumber, ProjectNumber) REFERENCES Projects(JobNumber, ProjectNumber)";
            string queryString2k = $"ALTER TABLE Tasks ADD FOREIGN KEY(JobNumber, ProjectNumber, Component) REFERENCES Components(JobNumber, ProjectNumber, Component)";


            using (OleDbConnection connection2 = new OleDbConnection($"{ConnBase}{destinationPath}"))
            {
                OleDbCommand cmd1 = new OleDbCommand(queryString2a, connection2);
                OleDbCommand cmd2 = new OleDbCommand(queryString2b, connection2);
                OleDbCommand cmd3 = new OleDbCommand(queryString2c, connection2);
                OleDbCommand cmd4 = new OleDbCommand(queryString2d, connection2);
                OleDbCommand cmd5 = new OleDbCommand(queryString2e, connection2);
                OleDbCommand cmd6 = new OleDbCommand(queryString2f, connection2);
                OleDbCommand cmd7 = new OleDbCommand(queryString2g, connection2);
                OleDbCommand cmd8 = new OleDbCommand(queryString2h, connection2);
                OleDbCommand cmd9 = new OleDbCommand(queryString2i, connection2);
                OleDbCommand cmd10 = new OleDbCommand(queryString2j, connection2);
                OleDbCommand cmd11 = new OleDbCommand(queryString2k, connection2);

                connection2.Open();

                cmd1.ExecuteNonQuery();
                cmd2.ExecuteNonQuery();
                cmd3.ExecuteNonQuery();
                cmd4.ExecuteNonQuery();
                cmd5.ExecuteNonQuery();
                cmd6.ExecuteNonQuery();
                cmd7.ExecuteNonQuery();
                cmd8.ExecuteNonQuery();
                cmd9.ExecuteNonQuery();
                cmd10.ExecuteNonQuery();
                cmd11.ExecuteNonQuery();
            }
        }

        public static void CopyDatabase2(string destinationPath, string source = "")
        {
            //Access.Dao.
        }
    }
}
