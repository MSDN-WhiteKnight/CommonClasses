using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace CommonClasses
{
    public enum ExcelFormats
    {
        Auto =0,
        Xls,
        Xlsx
    }

    public enum OleDbProviders
    {
        Auto = 0,
        Jet,
        ACE
    }

    public class ExcelDatabaseParams : DatabaseParams
    {
        /// <summary>
        /// Construct ExcelDatabaseParams object based on connection string
        /// </summary>
        public ExcelDatabaseParams(string constr)
        {
            this.ConnectionString = constr;
        }       

        public ExcelDatabaseParams(
            string file, 
            bool headers, 
            ExcelFormats format = ExcelFormats.Auto,
            OleDbProviders provider = OleDbProviders.Auto)
        {
            string p = "";
            string extended_properties = "";
            string ext = System.IO.Path.GetExtension(file);

            switch (provider)
            {
                case OleDbProviders.ACE: p = "Microsoft.ACE.OLEDB.12.0";break;

                case OleDbProviders.Jet: p = "Microsoft.Jet.OLEDB.4.0"; break;

                case OleDbProviders.Auto:
                    if (format == ExcelFormats.Xlsx || ext == ".xlsx" || IntPtr.Size >= 8)
                    {
                        p = "Microsoft.ACE.OLEDB.12.0";
                    }
                    else
                    {
                        p = "Microsoft.Jet.OLEDB.4.0";
                    }
                    break;
            }

            /*Формирование расширенных свойств подключения*/
            switch (format)
            {
                case ExcelFormats.Xlsx: extended_properties = "Excel 12.0 Xml;"; break;

                case ExcelFormats.Xls: extended_properties = "Excel 8.0;"; break;

                case ExcelFormats.Auto:
                    
                    if(ext == ".xls") extended_properties = "Excel 8.0;";
                    else extended_properties = "Excel 12.0 Xml;";
                    break;
            }

            if (headers) extended_properties += "HDR=YES;";
            else extended_properties += "HDR=NO;";

            
            OleDbConnectionStringBuilder b = new OleDbConnectionStringBuilder();
            b.Provider = p;
            b.DataSource = file;
            b["Extended Properties"] = extended_properties;            

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

/*Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyExcel.xls;
Extended Properties="Excel 8.0;HDR=Yes;IMEX=1";*/
