using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace PEGA.SI.One.Common
{
    public class NPOIUtils
    {
        public static MemoryStream ExportToExcel<T>(List<T> objects)
        {
            IWorkbook book = new XSSFWorkbook();
            ISheet sheet = book.CreateSheet("Data");
            IRow firstRow = sheet.CreateRow(0);

            Type tType = typeof(T);
            PropertyInfo[] pis = tType.GetProperties();

            for (int i = 0; i < pis.Length; i++)
            {
                ICell cell = firstRow.CreateCell(i);  //在第一行中创建单元格
                cell.SetCellValue(pis[i].Name);//循环往第一行的单元格中添加数据
            }

            for (int i = 0; i < objects.Count; i++)
            {
                IRow otherRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < pis.Length; j++)
                {
                    ICell cell = otherRow.CreateCell(j);
                    cell.SetCellValue(Convert.ToString( pis[j].GetValue(objects[i])));
                }
            }
            //using (FileStream fs = new FileStream("d:/abc.xlsx", FileMode.OpenOrCreate, FileAccess.Write))
            //{
            //    book.Write(fs);   //向打开的这个xls文件中写入mySheet表并保存。

            //}
            MemoryStream ms = new MemoryStream();
            book.Write(ms);
            return ms;
        }


        public static MemoryStream ExportToExcel(DataTable dt, string header)
        {
            XSSFWorkbook book = new XSSFWorkbook();
            ISheet sheet1 = book.CreateSheet("資料信息");
            IRow row1 = sheet1.CreateRow(0);
            ICellStyle dateStyle = book.CreateCellStyle();
            dateStyle.DataFormat = book.CreateDataFormat().GetFormat("yyyy/m/d h:mm:ss");

            ICellStyle numberStyle = book.CreateCellStyle();
            numberStyle.DataFormat = book.CreateDataFormat().GetFormat("0.00000");

            ICellStyle textStyle = book.CreateCellStyle();
            textStyle.DataFormat = book.CreateDataFormat().GetFormat("@");
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                DataColumn col = dt.Columns[i];
                row1.CreateCell(i).SetCellValue(col.ColumnName);
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow rowtemp = sheet1.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    rowtemp.CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                }
            }
            MemoryStream ms = new MemoryStream();
            book.Write(ms);
            //ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }


        public static DataTable GetDataTableFromFile(string fileName)
        {

            XSSFWorkbook hssfworkbook;
            #region//初始化信息
            try
            {
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    hssfworkbook = new XSSFWorkbook(file);
                }
            }
            catch (Exception e1)
            {
                throw e1;
            }
            #endregion
            ISheet sheet = hssfworkbook.GetSheetAt(0);

            DataTable table = new DataTable();
            IRow headerRow = sheet.GetRow(1);//第二行为标题行
            int cellCount = headerRow.LastCellNum;//LastCellNum = PhysicalNumberOfCells
            int rowCount = sheet.LastRowNum;//LastRowNum = PhysicalNumberOfRows ‐ 1
            //handling header.
            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }
            for (int i = (sheet.FirstRowNum + 2); i <= rowCount; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = table.NewRow();
                if (row != null)
                {
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            dataRow[j] = GetCellValue(row.GetCell(j)).Trim();
                        }
                    }
                }
                table.Rows.Add(dataRow);
            }
            return table;

        }

        private static string GetCellValue(ICell cell)
        {
            string res = null;
            if (cell == null)
                res = string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    res = string.Empty;
                    break;
                case CellType.Boolean:
                    res = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Error:
                    res = cell.ErrorCellValue.ToString();
                    break;
                case CellType.Numeric:
                    res = cell.NumericCellValue.ToString();
                    break;
                case CellType.Unknown:
                default:
                    res = cell.ToString();
                    break;
                case CellType.String:
                    res = cell.StringCellValue;
                    break;
                case CellType.Formula:
                    try
                    {
                        HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        res = cell.ToString();
                    }
                    catch
                    {
                        res = cell.NumericCellValue.ToString();
                    }
                    break;
            }
            return res;
        }
    }
}
