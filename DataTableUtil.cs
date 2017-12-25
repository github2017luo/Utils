using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PEGA.SI.One.Common
{
    public class DataTableUtil
    {
        /// <summary>
        /// 將string[,]二維數組轉換為DataTable，其中第一行為列表名
        /// </summary>
        /// <param name="arry"></param>
        /// <returns></returns>
        public static DataTable GetDataTableFromArray(string[,] arry)
        {
            DataTable dt = new DataTable();
            dt.TableName = "result";
            int rowCount = arry.GetLength(0);
            int columnCount = arry.GetLength(1);
            for (int i = 0; i < rowCount; i++)
            {
                if (i == 0)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        dt.Columns.Add(arry[i, j], typeof(string));
                    }
                }
                else
                {
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = arry[i, j];
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }
        public static DataTable SortDataTableByExpression(DataTable dt, string expression)
        {
            int rowCount = dt.Rows.Count;
            int columnCount = dt.Columns.Count;
            DataRow[] dataRows = dt.Select("1=1", expression);
            DataTable dtOrdered = new DataTable();
            dtOrdered.TableName = "result";
            for (int i = 0; i < columnCount; i++)
            {
                dtOrdered.Columns.Add(dt.Columns[i].ColumnName, typeof(string));
            }
            for (int i = 0; i < rowCount; i++)
            {
                DataRow dr = dtOrdered.NewRow();
                for (int j = 0; j < columnCount; j++)
                {
                    dr[j] = dataRows[i][j].ToString();
                }
                dtOrdered.Rows.Add(dr);

            }
            return dtOrdered;
        }

        /// <summary>
        /// 將DataTable轉換為string[,]二維數組，其中string[,]第一行為列名
        /// </summary>
        /// <param name="arry"></param>
        /// <returns></returns>
        public static string[,] GetStringArrayFromDataTable(DataTable dt)
        {
            int rowCount = dt.Rows.Count;
            int columnCount = dt.Columns.Count;
            string[,] array = new string[rowCount + 1, columnCount];
            for (int i = 0; i < columnCount; i++)
            {
                array[0, i] = dt.Columns[i].ColumnName;
            }
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    array[i + 1, j] = dt.Rows[i][j].ToString();
                }
            }
            return array;
        }
    }
}
