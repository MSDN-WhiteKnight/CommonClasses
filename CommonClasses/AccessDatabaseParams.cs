using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace CommonClasses
{
    public class AccessDatabaseParams : DatabaseParams
    {
        /// <summary>
        /// Construct AccessDatabaseParams object based on connection string
        /// </summary>
        public AccessDatabaseParams(string constr)
        {
            this.ConnectionString = constr;
        }

        public AccessDatabaseParams(
            string file,
            OleDbProviders provider)
        {
            string p = "";
            string ext = System.IO.Path.GetExtension(file);

            switch (provider)
            {
                case OleDbProviders.ACE: p = "Microsoft.ACE.OLEDB.12.0"; break;

                case OleDbProviders.Jet: p = "Microsoft.Jet.OLEDB.4.0"; break;

                case OleDbProviders.Auto:
                    if (ext == ".accdb" || IntPtr.Size >= 8)
                    {
                        p = "Microsoft.ACE.OLEDB.12.0";
                    }
                    else
                    {
                        p = "Microsoft.Jet.OLEDB.4.0";
                    }
                    break;
            }

            OleDbConnectionStringBuilder b = new OleDbConnectionStringBuilder();
            b.Provider = p;
            b.DataSource = file;

            this.ConnectionString = b.ConnectionString;
        }

        public override DbCommand BuildCommand()
        {
            return new OleDbCommand();
        }

        public override DbConnection BuildConnection()
        {
            return new OleDbConnection(this.ConnectionString);
        }

        public override DbParameter BuildParameter(string name, object value)
        {
            return new OleDbParameter(name, value);
        }
    }
}
