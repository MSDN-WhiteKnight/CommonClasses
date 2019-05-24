/* Copyright: MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight)
 * CommonClasses: database parameters
 */
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Text;


namespace CommonClasses
{
    /// <summary>
    /// Represents a set of parameters used to connect to a database. Base class for specific database type bindings.
    /// </summary>
    public abstract class DatabaseParams
    {
        public string ConnectionString { get; set; }

        /// <summary>
        /// When overridden in derived class, constructs a DbConnection object for specific database
        /// </summary>
        public abstract DbConnection BuildConnection();

        /// <summary>
        /// When overridden in derived class, constructs DbCommand object for specific database
        /// </summary>
        public abstract DbCommand BuildCommand();

        /// <summary>
        /// When overridden in derived class, constuct a parameter object for parametrized queries to specific database
        /// </summary>
        public abstract DbParameter BuildParameter(string name,object value);
    }

    /// <summary>
    /// Represents parameters used for connection to MS SQL Server Database
    /// </summary>
    public class SqlDatabaseParams : DatabaseParams
    {
        /// <summary>
        /// Construct SqlDatabaseParams object based on connection string
        /// </summary>
        public SqlDatabaseParams(string constr)
        {
            this.ConnectionString = constr;
        }

        /// <summary>
        /// Construct new SqlDatabaseParams object
        /// </summary>
        /// <param name="server">Server instance name</param>
        /// <param name="db">Initial database</param>
        /// <param name="user">User name</param>
        /// <param name="pass">Password</param>
        public SqlDatabaseParams(string server, string db,string user = "", string pass = "")
        {
            SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder();
            sb.DataSource = server;
            sb.InitialCatalog = db;
            sb.UserID = user;
            sb.Password = pass;
            this.ConnectionString = sb.ConnectionString;
        }           

        public override DbConnection BuildConnection()
        {
            var conn = new SqlConnection(this.ConnectionString);
            return conn;
        }

        public override DbCommand BuildCommand()
        {
            return new System.Data.SqlClient.SqlCommand();
        }

        public override DbParameter BuildParameter(string name, object value)
        {
            return new SqlParameter(name, value);
        }
    }
}
