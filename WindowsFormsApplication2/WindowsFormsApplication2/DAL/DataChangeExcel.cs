using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication2.DAL
{
    public class DataChangeExcel
    {
        public static void TableToExcel(DataTable dt, string file)
        {
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return; }
            ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Overview") : workbook.CreateSheet(dt.TableName);

            //表头  
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件  
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }


        /// <summary>
        /// 数据库转为excel表格
        /// </summary>
        /// <param name="dataTable">数据库数据</param>
        /// <param name="SaveFile">导出的excel文件</param>
        public static void DataSetToExcel(DataTable dataTable, string SaveFile)         {
             Microsoft.Office.Interop.Excel.Application excel;
             Microsoft.Office.Interop.Excel._Workbook workBook;
             Microsoft.Office.Interop.Excel._Worksheet workSheet;
             object misValue = System.Reflection.Missing.Value;
             excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
             workBook = excel.Workbooks.Add(misValue);
             workSheet = (Microsoft.Office.Interop.Excel._Worksheet)workBook.ActiveSheet;
             int rowIndex = 1;
             int colIndex = 0;
             //取得标题
             foreach (DataColumn col in dataTable.Columns)
             {
                 colIndex++;
                 excel.Cells[1, colIndex] = col.ColumnName;
             }
             //取得表格中的数据
             foreach (DataRow row in dataTable.Rows)
             {
                 rowIndex++;
                 colIndex = 0;
                 foreach (DataColumn col in dataTable.Columns)
                 {
                     colIndex++;
                     excel.Cells[rowIndex, colIndex] =
                          row[col.ColumnName].ToString().Trim();
                    //设置表格内容居中对齐
                    
                       //excel.Cells[rowIndex, colIndex].HorizontalAlignment =
                       //Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                 }
             }
             excel.Visible = false;
             workBook.SaveAs(SaveFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue,
                 misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                 misValue, misValue, misValue, misValue, misValue);
             dataTable = null;
             workBook.Close(true, misValue, misValue);
             excel.Quit();
             //PublicMethod.Kill(excel);//调用kill当前excel进程
             releaseObject(workSheet);
             releaseObject(workBook);
             releaseObject(excel);
         }
 
         private static void releaseObject(object obj)
         {
             try
             {
                 System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                 obj = null;
             }
             catch
             {
                 obj = null;
             }
            finally
             {
                 GC.Collect();
             }
         }
    }
}
