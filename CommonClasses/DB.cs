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

        static bool CheckSqlInput(string s)
        {
            return ! (s.Contains("'") || s.Contains("\"") || s.Contains("]") || s.Contains(",") || s.Contains(";"));
        }

        static string PrepareSqlInput(string s)
        {
            return s.Replace("'", "").Replace("\"", "").Replace("]", "").Replace(",", "").Replace(";", "").Replace(" ", "");
        }

        public static DataTable GetTable(string filepath, string table)
        {
            if (CheckSqlInput(table) == false) throw new ArgumentException("Table name contains SQL special characters");

            string ext = System.IO.Path.GetExtension(filepath);
            DB db = DB.FromFile(filepath);
            string sql;

            if ((ext == ".xls" || ext == ".xlsx") && !table.EndsWith("$")) table = table + "$";  
            sql = "SELECT * FROM [" + table + "]";

            var res = db.QueryTable(sql);
            return res;
        }

        public static int WriteTable(string filepath, string table, DataTable dt)
        {
            if (CheckSqlInput(table) == false) throw new ArgumentException("Table name contains SQL special characters");

            string ext = System.IO.Path.GetExtension(filepath);
            DB db = DB.FromFile(filepath);
            int res = 0;

            string sql;
            StringBuilder sb;
            DataColumn col;

            sb = new StringBuilder();
            sb.Append("INSERT INTO [" + table + "](");

            col = dt.Columns[0];
            sb.Append("["+PrepareSqlInput(col.ColumnName)+"]");
            for (int i = 1; i < dt.Columns.Count; i++)
            {
                col = dt.Columns[i];
                sb.Append(", [" + PrepareSqlInput(col.ColumnName) + "]");
            }
            sb.Append(") VALUES (");

            col = dt.Columns[0];
            sb.Append("@" + PrepareSqlInput(col.ColumnName));

            for (int i = 1; i < dt.Columns.Count; i++)
            {
                col = dt.Columns[i];
                sb.Append(",@" + PrepareSqlInput(col.ColumnName));
            }
            sb.Append(")");
            sql = sb.ToString();

            DbParameter[] pars = new DbParameter[dt.Columns.Count];

            if (!db.QueryTableNames().Contains(table))
            {
                sb = new StringBuilder();
                sb.Append("CREATE TABLE [" + table + "] (");

                col = dt.Columns[0];
                sb.Append("[" + PrepareSqlInput(col.ColumnName) + "] TEXT");
                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    col = dt.Columns[i];
                    sb.Append(", [" + PrepareSqlInput(col.ColumnName) + "] TEXT");
                }
                sb.Append(")");
                
                res += db.ExecuteSQL(sb.ToString(), new DbParameter[0]);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow row = dt.Rows[i];
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        pars[j] = db.CreateParameter(PrepareSqlInput(dt.Columns[j].ColumnName), row[j].ToString());
                    }
                    res += db.ExecuteSQL(sql, pars);
                }
            }
            else
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow row = dt.Rows[i];
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        pars[j] = db.CreateParameter(PrepareSqlInput(dt.Columns[j].ColumnName), row[j]);
                    }
                    res += db.ExecuteSQL(sql, pars);
                }
            }
            return res;
        }

        public static DataTable GetTable(string filepath, uint index)
        {            
            string ext = System.IO.Path.GetExtension(filepath);
            DB db = DB.FromFile(filepath);

            string table=null;

            uint i = 0;
            foreach (string s in db.QueryTableNames())
            {
                if (i == index) { table = s; break; }
                i++;
            }

            if (table == null) throw new ArgumentOutOfRangeException("index","Table with index "+index.ToString()+" does not exist");

            string sql;            
            sql = "SELECT * FROM [" + table + "]";

            var res = db.QueryTable(sql);
            return res;
        }

        public static IEnumerable<T> GetCollection<T>(string filepath, string table)
        {
            if (CheckSqlInput(table) == false) throw new ArgumentException("Table name contains SQL special characters");

            string ext = System.IO.Path.GetExtension(filepath);
            DB db = DB.FromFile(filepath);
            string sql;

            if ((ext == ".xls" || ext == ".xlsx") && !table.EndsWith("$")) table = table + "$"; 
            sql = "SELECT * FROM [" + table + "]";

            return db.QueryCollection<T>(sql, new DbParameter[0]);            
        }

        public static IEnumerable<T> GetCollection<T>(string filepath, uint index)
        {
            string ext = System.IO.Path.GetExtension(filepath);
            DB db = DB.FromFile(filepath);

            string table = null;

            uint i = 0;
            foreach (string s in db.QueryTableNames())
            {
                if (i == index) { table = s; break; }
                i++;
            }

            if (table == null) 
                throw new ArgumentOutOfRangeException("index","Table with index " + index.ToString() + " does not exist");

            string sql;
            sql = "SELECT * FROM [" + table + "]";

            return db.QueryCollection<T>(sql, new DbParameter[0]); 
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
        /// Convert database query results into a collection of objects. 
        /// This method will set public properties into corresponding query result column values. 
        /// </summary>
        public IEnumerable<T> ExtractCollection<T>(DbDataReader rd)
        {            
            var props = typeof(T).GetProperties();

            while (true)
            {
                if (rd.Read() == false) break;
                T item = Activator.CreateInstance<T>();

                foreach (var prop in props)
                {
                    string fname = prop.Name;

                    object[] attrs = prop.GetCustomAttributes(typeof(DbFieldAttribute),true);
                    if (attrs.Length > 0)
                    {
                        fname = (attrs[0] as DbFieldAttribute).FieldName;
                    }

                    object val;

                    try
                    {
                        val = rd[fname];
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

                yield return item;                
            }            
        }

        /// <summary>
        /// Executes SQL and returns results as a collection of objects.
        /// This method will set public properties into corresponding query result column values. 
        /// </summary>
        /// <param name="sql">Query text</param>
        /// <param name="args">Variable-length array of query parameters</param>   
        public IEnumerable<T> QueryCollection<T>(string sql, params DbParameter[] args)
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
                    var res = ExtractCollection<T>(rd);
                    foreach (var item in res) yield return item;
                }
            }            
        }

        /// <summary>
        /// Exceutes SQL and returns results as a string collection
        /// </summary>
        /// <param name="sql">Query text</param>
        /// <param name="args">Variable-length array of query parameters</param>   
        public IEnumerable<string> QueryStrings(string sql, params DbParameter[] args)
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
                    while (true)
                    {
                        if(rd.Read() == false) break;
                        string val = rd[0].ToString();
                        yield return val;
                    }
                }                
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

        public IEnumerable<string> QueryTableNames()
        {
            DbConnection con = this.Params.BuildConnection();
            
            using (con)
            {
                con.Open();
                DataTable dtSchema = con.GetSchema("Tables");
                foreach (DataRow row in dtSchema.Rows)
                {
                    yield return row.Field<string>("TABLE_NAME");
                }                
            }            
        }

    }

    public class DbFieldAttribute : System.Attribute
    {
        string fname;

        public DbFieldAttribute(string FieldName)
        {
            this.fname = FieldName;
        }

        public string FieldName { get { return fname; } }
    }
}
