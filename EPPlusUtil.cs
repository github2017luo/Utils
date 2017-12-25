using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Data;
using System.Reflection;
using System.ComponentModel;

namespace PEGA.SI.One.Common
{
    public static class EPPlusUtil
    {
        public static ExcelWorksheet GetList(Stream stream)
        {
            ExcelPackage excelPackage = new ExcelPackage(stream);
            ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets[1];
            return workSheet;
        }

        /// <summary>
        /// 创建一个Excel 后缀xlsx
        /// </summary>
        /// <param name="filePath">文件的绝对路径 (eg.D:\em.xlsx)</param>
        public static void CreateExcel(string file, int sheetCount = 3)
        {
            try
            {
                if (File.Exists(file)) File.Delete(file);
            }
            catch (Exception)
            {
                throw new Exception("创建Excel文件失败,该文件已存在并拒绝访问。");
            }
            try
            {
                using (var excel = new ExcelPackage(new FileInfo(file)))
                {
                    for (int i = 0; i < sheetCount; i++)
                    {
                        excel.Workbook.Worksheets.Add("Sheet" + (i + 1).ToString());
                    }
                    excel.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }

        /// <summary>
        /// 创建一个Excel
        /// </summary>
        /// <param name="file">文件的绝对路径 (eg.D:\em.xlsx)</param>
        /// <param name="list">根据List创建Sheet</param>
        public static void CreateExcel(string file, List<string> list)
        {
            try
            {
                if (File.Exists(file)) File.Delete(file);
            }
            catch (Exception)
            {
                throw new Exception("创建Excel文件失败,该文件已存在并拒绝访问。");
            }
            try
            {
                using (var excel = new ExcelPackage(new FileInfo(file)))
                {
                    foreach (var item in list)
                    {
                        excel.Workbook.Worksheets.Add(item);
                    }
                    excel.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }
        public static void SetValue(string filePath, int sheetIndex, int Row, int column, string content = "")
        {
            if (!File.Exists(filePath))
            {
                CreateExcel(filePath);
            }
            using (ExcelPackage pack = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = pack.Workbook.Worksheets[sheetIndex];
                sheet.SetValue(Row, column, content);
                pack.Save();
            }
        }
        public static object GetValue(string filePath, int sheetIndex, int Row, int column)
        {
            using (ExcelPackage pack = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = pack.Workbook.Worksheets[sheetIndex];
                return sheet.GetValue(Row, column);
            }
        }
        /// <summary>
        /// DataTable写到Excel
        /// </summary>
        /// <param name="filePath">xlsx路径,如果不存在则尝试创建</param>
        /// <param name="dt">Datatable</param>
        /// <param name="sheetIndex">写到Excel的Sheet索引,缺省情况下写第一个Sheet</param>
        /// <param name="containColums">是否包含列标题,缺省情况下包含列标题</param>
        public static void DataTable2Excel(string filePath, DataTable dt, int sheetIndex = 1, bool containColums = true)
        {
            if (!File.Exists(filePath))
            {
                CreateExcel(filePath);
            }
            try
            {
                using (ExcelPackage pack = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet sheet = pack.Workbook.Worksheets[sheetIndex];
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (containColums)//包含列标题
                            {
                                if (i == 0)
                                {
                                    sheet.SetValue(1, j + 1,Convert.ToString( dt.Columns[j]));
                                }
                                sheet.SetValue(i + 2, j + 1, Convert.ToString(dt.Rows[i][j]));//第一列是标题，所以第二列给DT的数据
                            }
                            else
                            {
                                sheet.SetValue(i + 1, j + 1, Convert.ToString(dt.Rows[i][j]));
                            }
                        }
                    }
                    pack.Save();
                }
            }
            catch (Exception)
            {
                throw new Exception("Excel写Sheet发生异常:未能找到要写入的Sheet");
            }

        }

       


        private static DataTable AsDataTable<T>(this IEnumerable<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            var table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }


        public static byte[] ExportExcel<T>(IEnumerable<T> data, bool containColums = true)
        {
            DataTable dt = AsDataTable(data);
            MemoryStream ms = new MemoryStream();
            ExcelPackage pack = new ExcelPackage();
            ExcelWorksheet sheet = pack.Workbook.Worksheets.Add("sheet1"); //创建sheet  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (containColums)//包含列标题
                    {
                        if (i == 0)
                        {
                            sheet.SetValue(1, j + 1, Convert.ToString(dt.Columns[j]));
                        }
                        sheet.SetValue(i + 2, j + 1, Convert.ToString(dt.Rows[i][j]));//第一列是标题，所以第二列给DT的数据
                    }
                    else
                    {
                        sheet.SetValue(i + 1, j + 1, Convert.ToString(dt.Rows[i][j]));
                    }
                }
            }
            pack.SaveAs(ms);
            return ms.ToArray();
        }
    }
}
