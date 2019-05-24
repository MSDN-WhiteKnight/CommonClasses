using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Text;

namespace CommonClasses
{
    public static class Utils
    {
        //находит максимальную длину элемента для всех столбцов DataTable
        public static int[] GetTableMaxColumnWidths(DataTable dt)
        {
            int[] res = new int[dt.Columns.Count];

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                res[i] = dt.Columns[i].ColumnName.Length;
                foreach (DataRow row in dt.Rows)
                {
                    string s = row[i].ToString();
                    if (s.Length > res[i]) res[i] = s.Length;
                }
            }
            return res;
        }

        //выводит таблицу в TextWriter
        public static void PrintTable(DataTable dt, TextWriter wr)
        {
            int[] len = GetTableMaxColumnWidths(dt);
            string s;

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                DataColumn col = dt.Columns[i];
                s = col.ColumnName.PadLeft(len[i]);
                wr.Write(s + " | ");                
            }
            wr.WriteLine();

            foreach (DataRow row in dt.Rows)
            {
                for (int i = 0; i < row.ItemArray.Length; i++)
                {
                    s = row[i].ToString().PadLeft(len[i]);
                    wr.Write(s + " | ");
                }
                wr.WriteLine();
            }
        }
    }
}
