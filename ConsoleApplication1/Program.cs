using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using CommonClasses;
using CommonClasses.Console;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var res = DB.GetTable("c:\\Test\\Database1.mdb", "Table1");

                Utils.PrintTable(res, Console.Out);                
                
            }
            catch (Exception ex)
            {
                ErrorHandler.Default.OnError(ex);
            }
            Console.ReadKey();
        }
    }
}
