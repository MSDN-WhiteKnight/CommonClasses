using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace CommonClasses
{
    public enum ErrorLevel
    {
        Undefined = 0,
        Info,
        Warning,
        Error
    }

    public abstract class ErrorHandlerBase
    {
        /// <summary>
        /// Путь к файлу для этого экземпляра обработчика
        /// </summary>
        protected string _logfile;

        protected void InitDefaultLogPath()
        {
            var pr = System.Diagnostics.Process.GetCurrentProcess();
            using (pr)
            {
                string path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                    "Logs");
                if (Directory.Exists(path) == false) Directory.CreateDirectory(path);
                this._logfile = Path.Combine(path, pr.ProcessName + ".log");
            }
        }

        public abstract void DisplayError(Exception ex, string action = "", ErrorLevel level = ErrorLevel.Undefined);

        public abstract void DisplayError(string message, ErrorLevel level = ErrorLevel.Undefined);

        /// <summary>
        /// Обработка исключения с выводом сообщения об ошибке, без указания действия
        /// </summary>
        /// <param name="ex">Данные исключения</param>
        public void OnError(Exception ex)
        {
            OnError(ex, "", false);
        }

        /// <summary>
        /// Обработка исключения с выводом сообщения об ошибке и указанным действием
        /// </summary>
        /// <param name="ex">Данные исключения</param>
        /// <param name="action">Действие, при котором произошла ошибка</param>
        public void OnError(Exception ex, string action)
        {
            OnError(ex, action, false);
        }

        /// <summary>
        /// Обработка исключения указанным признаком вывода собщения и действием
        /// </summary>
        /// <param name="ex">Данные исключения</param>
        /// <param name="action">Действие, при котором произошла ошибка</param>
        /// <param name="silent">Отключает вывод сообщения об ошибке</param>
        public void OnError(Exception ex, string action, bool silent, ErrorLevel level= ErrorLevel.Error)
        {
            if (action == null) action = "";
            if (ex == null) ex = new ApplicationException("Unknown error");
                       

            /*Запись ошибки в файл отчета*/
            StreamWriter wr = new StreamWriter(_logfile, true);
            using (wr)
            {

                wr.WriteLine("**** " + DateTime.Now.ToString() + " ****");
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

                wr.WriteLine("*******************************");
                wr.WriteLine();
            }

            if (!silent)
            {
                DisplayError(ex, action, level);
            }

        }

        public void OnError(string mes, bool silent = false, ErrorLevel level = ErrorLevel.Undefined)
        {
            if (mes == null) mes = "";           


            /*Запись ошибки в файл отчета*/
            StreamWriter wr = new StreamWriter(_logfile, true);
            using (wr)
            {

                wr.WriteLine("**** " + DateTime.Now.ToString() + " ****");
                
                wr.WriteLine(mes);                

                wr.WriteLine("*******************************");
                wr.WriteLine();
            }

            if (!silent)
            {
                DisplayError(mes,level);
            }

        }


    }
}
