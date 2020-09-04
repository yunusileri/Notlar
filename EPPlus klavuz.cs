//EPPlus version 4.5.3.3


using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace excelexport
{
    public class Program
    {

        static void Main(string[] args)
        {
            DataTableExport();
            ListExport();
            ExcelToDataTable();
            Console.ReadKey();
        }
        public static string filePath = @"C:\Users\yunusi\Desktop";
        static void DataTableExport()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("tarih");
            dt.Columns.Add("isim");

            dt.Rows.Add(DateTime.Now.ToString("yyyy-MM-dd HH:ss"), "yunus");
            dt.Rows.Add(DateTime.Now.AddDays(1).ToString("yyyy-MM-dd HH:ss"), "yunusileri");

            var excel = ExcelExport(dt);
            ExcelSave(excel, filePath, "DataTableExport");
        }

        static void ListExport()
        {
            List<test> list = new List<test>();
            for (int i = 0; i < 100; i++)
            {
                list.Add(new test() { isim = $"test{i + 1}", Tarih = DateTime.Now.ToString("yyyy-MM-dd HH:ss") });

            }
            var dt = CreateDataTable(list);
            var excel = ExcelExport(dt);
            ExcelSave(excel, filePath, "ListExport");
        }


        static void ExcelToDataTable()
        {
            var dt = GetDataTableFromExcel(Path.Combine(filePath, "ListExport.xlsx"));

            foreach (DataColumn dc in dt.Columns)
            {
                Console.Write($"{dc.ColumnName}\t\t");
            }
            Console.WriteLine("");
            foreach (DataRow dr in dt.Rows)
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    Console.Write($"{dr[dc.ColumnName]}\t\t");
                }
                Console.WriteLine("");
            }

        }
        //List<T> gönder DataTable al
        public static DataTable CreateDataTable<T>(IEnumerable<T> list)
        {
            Type type = typeof(T);
            var properties = type.GetProperties();

            DataTable dataTable = new DataTable();
            foreach (PropertyInfo info in properties)
            {
                dataTable.Columns.Add(new DataColumn(info.Name, Nullable.GetUnderlyingType(info.PropertyType) ?? info.PropertyType));
            }

            foreach (T entity in list)
            {
                object[] values = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    values[i] = properties[i].GetValue(entity);
                }

                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        public static ExcelPackage ExcelExport(DataTable dt, bool AutoFit = true)
        {
            string methodName = "ExcelExport()";


            ExcelPackage excel = new ExcelPackage();

            try
            {
                var workSheet = excel.Workbook.Worksheets.Add("Page1");
                List<string> header = new List<string>();
                int i = 0;
                foreach (DataColumn column in dt.Columns)
                {
                    header.Add(column.ColumnName);
                    workSheet.Cells[1, (i + 1)].Value = header[i];
                    if (AutoFit)
                    {
                        workSheet.Column(i + 1).AutoFit();
                    }
                    i++;
                }
                int y = 2; // sutun 
                int x = 0; // satır
                foreach (DataRow dr in dt.Rows)
                {
                    foreach (var item in header)
                    {
                        workSheet.Cells[y, x + 1].Value = dr[item].ToString();
                        x++;
                    }
                    x = 0;
                    y++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{methodName} {ex}");

            }
            return excel;
        }
        public static string ExcelSave(ExcelPackage excel, string filePath, string FileName)
        {
            string methodName = $"ExcelSave(excel, {filePath}, {FileName})";


            string FullFilePath = string.Empty;
            try
            {
                FileName += ".xlsx";
                FullFilePath = Path.Combine(filePath, FileName);

                FileStream objFileStrm = System.IO.File.Create(FullFilePath);
                objFileStrm.Close();
                System.IO.File.WriteAllBytes(FullFilePath, excel.GetAsByteArray());
                excel.Dispose();


            }
            catch (Exception ex)
            {
                Console.WriteLine($"{methodName} {ex}");

            }
            return FullFilePath;
        }

        //Excel to DataTable
        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }
    }
    class test
    {
        public string Tarih { get; set; }
        public string isim { get; set; }
    }
}
