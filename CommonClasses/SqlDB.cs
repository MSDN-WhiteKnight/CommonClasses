/* Copyright: MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight)
 * CommonClasses: MSSQL data access layer
 */
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace CommonClasses
{
    /// <summary>
    /// Provides MSSQL-specific methods for SQL execution and data access.
    /// </summary>
    public class SqlDB : DB
    {        

        /// <summary>
        /// Gets or sets database access parameters object
        /// </summary>
        public new DatabaseParams Params
        {
            get { return _Params; }

            set
            {
                if (!(value is SqlDatabaseParams))
                {
                    throw new InvalidOperationException(
                        "Params property must be an instance of SqlDatabaseParams class"
                        );
                }
                _Params = value;
            }
        }

        public SqlDB (SqlDatabaseParams pars)
        {
            this.Params = pars;
        }

        /// <summary>
        /// Executes specified stored procedure and returns results as a DataTable
        /// </summary>
        /// <param name="proc">Procedure name</param>
        /// <param name="args">Variable-length array of procedure parameters</param>
        public DataTable QueryTableSP(string proc, params DbParameter[] args)
        {
            DbConnection con = this.Params.BuildConnection();

            using (con)
            {
                con.Open();
                System.Data.SqlClient.SqlCommand cmd = (System.Data.SqlClient.SqlCommand)this.Params.BuildCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = proc;
                cmd.Connection = (System.Data.SqlClient.SqlConnection)con;
                for (int i = 0; i < args.Length; i++)
                {
                    cmd.Parameters.Add(args[i]);
                }

                DbDataReader rd = cmd.ExecuteReader();
                using (rd)
                {
                    DataTable dt = new DataTable();
                    dt.Load(rd);

                    return dt;
                }
            }

        }

        /// <summary>
        /// Executes specified stored procedure and returns results as a collection of objects.
        /// Sets object's public properties into corresponding query column values.
        /// </summary>
        /// <param name="proc">Procedure name</param>
        /// <param name="args">Variable-length array of procedure parameters</param>
        public IEnumerable<T> QueryCollectionSP<T>(string proc, params DbParameter[] args)
        {
            DbConnection con = this.Params.BuildConnection();

            using (con)
            {
                con.Open();
                System.Data.SqlClient.SqlCommand cmd = (System.Data.SqlClient.SqlCommand)this.Params.BuildCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = proc;
                cmd.Connection = (System.Data.SqlClient.SqlConnection)con;
                for (int i = 0; i < args.Length; i++)
                {
                    cmd.Parameters.Add(args[i]);
                }

                DbDataReader rd = cmd.ExecuteReader();
                using (rd)
                {
                    var res = ExtractCollection<T>(rd);
                    foreach (var item in res) yield return item;
                }
            }

        }

        /// <summary>
        /// Executes specified stored procedure and returns results as a collection of strings
        /// </summary>
        /// <param name="proc">Procedure name</param>
        /// <param name="args">Variable-length array of procedure parameters</param>
        public IEnumerable<string> QueryStringsSP(string proc, params DbParameter[] args)
        {
            DbConnection con = this.Params.BuildConnection();

            using (con)
            {
                con.Open();
                System.Data.SqlClient.SqlCommand cmd = (System.Data.SqlClient.SqlCommand)this.Params.BuildCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = proc;
                cmd.Connection = (System.Data.SqlClient.SqlConnection)con;
                for (int i = 0; i < args.Length; i++)
                {
                    cmd.Parameters.Add(args[i]);
                }


                DbDataReader rd = cmd.ExecuteReader();                

                using (rd)
                {
                    while (true)
                    {
                        if (rd.Read() == false) break;
                        string val = rd[0].ToString();
                        yield return val;
                    }
                }                
            }
        }


        /// <summary>
        /// Executes specified stored procedure and returns result object as a scalar value
        /// </summary>
        /// <param name="proc">Procedure name</param>
        /// <param name="args">Variable-length array of procedure parameters</param>
        public object QueryValueSP(string proc, params DbParameter[] args)
        {
            DbConnection con = this.Params.BuildConnection();

            using (con)
            {
                con.Open();
                System.Data.SqlClient.SqlCommand cmd = (System.Data.SqlClient.SqlCommand)this.Params.BuildCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = proc;
                cmd.Connection = (System.Data.SqlClient.SqlConnection)con;
                for (int i = 0; i < args.Length; i++)
                {
                    cmd.Parameters.Add(args[i]);
                }

                object val = cmd.ExecuteScalar();
                return val;
            }

        }

        /// <summary>
        /// Executes specified stored procedure
        /// </summary>
        /// <param name="proc">Procedure name</param>
        /// <param name="args">Variable-length array of procedure parameters</param>
        /// <returns>Number of rows affected</returns>
        public int ExecuteSP(string proc, params DbParameter[] args)
        {
            DbConnection con = this.Params.BuildConnection();

            using (con)
            {
                con.Open();
                System.Data.SqlClient.SqlCommand cmd = (System.Data.SqlClient.SqlCommand)this.Params.BuildCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = proc;
                cmd.Connection = (System.Data.SqlClient.SqlConnection)con;
                for (int i = 0; i < args.Length; i++)
                {
                    cmd.Parameters.Add(args[i]);
                }

                int res = cmd.ExecuteNonQuery();
                return res;
            }

        }
    }
}
