using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

namespace CommonClasses.Console
{
    public class ErrorHandler : ErrorHandlerBase
    {
        public static ErrorHandler Default { get; set; }

        static ErrorHandler()
        {
            Default = new ErrorHandler();
        }

        public ErrorHandler()
        {
            InitDefaultLogPath();
        }

        public ErrorHandler(string logpath)
        {
            this._logfile = logpath;
        }

        public override void DisplayError(Exception ex, string action = "", ErrorLevel level = ErrorLevel.Undefined)
        {
            if (action == null) action = "";
            if (ex == null) ex = new ApplicationException("Unknown error");

            var wr = global::System.Console.Error;            

            wr.WriteLine("*** " + DateTime.Now.ToString() + " ***");
            if (action.Length > 0)
            {
                wr.WriteLine("Action: *" + action + "*");
            }
            wr.WriteLine(ex.ToString());
            wr.WriteLine("Module: " + ex.Source);
            if (ex.HelpLink != null)
            {
                wr.WriteLine("Help link: " + ex.HelpLink);
            }

            if (ex.Data != null && ex.Data.Count > 0)
            {
                wr.WriteLine("-- Additional data --");
                foreach (object val in ex.Data.Keys)
                {
                    wr.WriteLine(val + ": " + ex.Data[val]);
                }
                wr.WriteLine("----------------------");
            }

            wr.WriteLine("***********************************");
            wr.WriteLine();

        }

        public override void DisplayError(string message, ErrorLevel level = ErrorLevel.Undefined)
        {
            if (message == null) message = "";                        

            var wr = global::System.Console.Error;
            if (level == ErrorLevel.Info) wr = global::System.Console.Out;

            wr.WriteLine("*** " + DateTime.Now.ToString() + " ***");

            wr.WriteLine(message);

            wr.WriteLine("***********************************");
            wr.WriteLine();

        }
    }
}
