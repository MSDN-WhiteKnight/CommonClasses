/* Copyright: MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight)
 * CommonClasses: generic data access layer
 */
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Reflection;
using System.Linq;
using System.Text;

namespace CommonClasses
{
    /// <summary>
    /// Provides methods for SQL execution and querying data from databases. 
    /// Classes that implement functionality specific to certain DB type derive from this class.
    /// </summary>
    public class DB
    {
        /// <summary>
        /// Object that we use to construct connections and other DB-specific objects
        /// </summary>
        protected DatabaseParams _Params;

        public static DB Create(DatabaseParams pars)
        {
            DB db = new DB();
            db.Params = pars;
            return db;
        }

        public static DB FromFile(string path)
        {
            string ext = System.IO.Path.GetExtension(path);

            if (ext == ".mdb" || ext == ".accdb")
                return DB.Create(new AccessDatabaseParams(path, OleDbProviders.Auto));
            else if (ext == ".xls" || ext == ".xlsx")
                return DB.Create(new ExcelDatabaseParams(path, true));
            else
                throw new InvalidOperationException("File type '"+ext+"' is not supported");
        }

        public static DataTable GetTable(string filepath, string table)
        {
            string ext = System.IO.Path.GetExtension(filepath);
            DB db = DB.FromFile(filepath);

            string sql;

            if (ext == ".xls" || ext == ".xlsx") sql = "SELECT * FROM [" + table + "$]";
            else sql = "SELECT * FROM [" + table + "]";

            var res = db.QueryTable(sql);
            return res;
        }

        /// <summary>
        /// Gets or sets database access parameters object
        /// </summary>
        public DatabaseParams Params
        {
            get { return _Params; }

            set { _Params = value; }
        }

        /// <summary>
        /// Creates parameter instance for parametized queries used with this class
        /// </summary>                
        public DbParameter CreateParameter(string name, object value)
        {
            return this._Params.BuildParameter(name, value);
        }

        /// <summary>
        /// Executes SQL and returns result rows as DataTable object
        /// </summary>
        /// <param name="sql">Query text</param>
        /// <param name="args">Variable-length array of query parameters</param>        
        public DataTable QueryTable(string sql, params DbParameter[] args)
        {
            DbConnection con = this.Params.BuildConnection();            

            using (con)
            {
                con.Open();
                DbCommand cmd = this.Params.BuildCommand();              
                cmd.CommandText = sql;
                cmd.Connection = con;
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
        /// Convert database query results into a colelction of object. 
        /// This method will set public properties into corresponding query result column values. 
        /// </summary>
        public ICollection<T> ExtractCollection<T>(DbDataReader rd)
        {
            List<T> list = new List<T>(100);
            var props = typeof(T).GetProperties();

            while (true)
            {
                if (rd.Read() == false) break;
                T item = Activator.CreateInstance<T>();

                foreach (var prop in props)
                {
                    object val;

                    try
                    {
                        val = rd[prop.Name];
                    }
                    catch (System.IndexOutOfRangeException)
                    {
                        val = null;
                    }

                    if (val == null) continue;
                    if (val == DBNull.Value) continue;

                    object converted;

                    try
                    {
                        if (prop.PropertyType.Equals(typeof(string)))
                            converted = val.ToString();
                        else if (prop.PropertyType.Equals(val.GetType()))
                            converted = val;
                        else
                            converted = Convert.ChangeType(val, prop.PropertyType);

                        prop.SetValue(item, converted, null);
                    }
                    catch (InvalidCastException ex)
                    {
                        System.Diagnostics.Debug.WriteLine(ex.ToString());
                    }
                }

                list.Add(item);
            }

            return list;
        }

        /// <summary>
        /// Executes SQL and returns results as a collection of objects.
        /// This method will set public properties into corresponding query result column values. 
        /// </summary>
        /// <param name="sql">Query text</param>
        /// <param name="args">Variable-length array of query parameters</param>   
        public ICollection<T> QueryCollection<T>(string sql, params DbParameter[] args)
        {        
            DbConnection con = this.Params.BuildConnection();

            using (con)
            {
                con.Open();
                DbCommand cmd = this.Params.BuildCommand();
                cmd.CommandText = sql;
                cmd.Connection = con;
                for (int i = 0; i < args.Length; i++)
                {
                    cmd.Parameters.Add(args[i]);
                }

                DbDataReader rd = cmd.ExecuteReader();
                using (rd)
                {
                    return ExtractCollection<T>(rd);
                }
            }
            
        }

        /// <summary>
        /// Exceutes SQL and returns results as a string collection
        /// </summary>
        /// <param name="sql">Query text</param>
        /// <param name="args">Variable-length array of query parameters</param>   
        public ICollection<string> QueryStrings(string sql, params DbParameter[] args)
        {
            DbConnection con = this.Params.BuildConnection();

            using (con)
            {
                con.Open();
                DbCommand cmd = this.Params.BuildCommand();
                cmd.CommandText = sql;
                cmd.Connection = con;
                for (int i = 0; i < args.Length; i++)
                {
                    cmd.Parameters.Add(args[i]);
                }


                DbDataReader rd = cmd.ExecuteReader();
                List<string> list = new List<string>(100);

                using (rd)
                {
                    while (true)
                    {
                        if(rd.Read() == false) break;
                        string val = rd[0].ToString();
                        list.Add(val);
                    }
                }

                return list;
            }

        }

        /// <summary>
        /// Executes sql and returns result object as a scalar value
        /// </summary>
        /// <param name="sql">Query text</param>
        /// <param name="args">Variable-length array of query parameters</param>   
        public object QueryValue(string sql, params DbParameter[] args)
        {
            DbConnection con = this.Params.BuildConnection();

            using (con)
            {
                con.Open();
                DbCommand cmd = this.Params.BuildCommand();
                cmd.CommandText = sql;
                cmd.Connection = con;
                for (int i = 0; i < args.Length; i++)
                {
                    cmd.Parameters.Add(args[i]);
                }

                object val = cmd.ExecuteScalar();
                return val;
            }

        }

        /// <summary>
        /// Executes passed SQL command on the database
        /// </summary>
        /// <param name="sql">Query text</param>
        /// <param name="args">Variable-length array of query parameters</param>   
        /// <returns>Number of rows affected</returns>
        public int ExecuteSQL(string sql, params DbParameter[] args)
        {
            DbConnection con = this.Params.BuildConnection();

            using (con)
            {
                con.Open();
                DbCommand cmd = this.Params.BuildCommand();
                cmd.CommandText = sql;
                cmd.Connection = con;
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
