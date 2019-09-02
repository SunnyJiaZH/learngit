using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Data;
using System.IO;
using System.Collections.Generic;
using NPOI.SS.Formula.Eval;
using NPOI;

namespace Unilever.Common
{
    /// <summary>
    /// NPOI 操作 Office 类
    /// </summary>
    public static class NPOI_Office
    {
        /// <summary>
        /// Excel 操作类
        /// </summary>
        public class NPOI_Excel
        {
            #region 通用版本 导入Excel
            /// <summary>
            /// 导入Excel
            /// </summary>
            public class ImportExcel
            {
                /// <summary>
                /// is or not office 2007
                /// </summary>
                private static bool is2007 = true;
                /// <summary>
                /// office Excel Row
                /// </summary>
                private static IRow row;
                /// <summary>
                /// 2007 or 2003 Row chanage
                /// </summary>
                private static object Row
                {
                    set
                    {
                        if (is2007)
                        { row = (XSSFRow)value; }
                        else
                        { row = (HSSFRow)value; }
                    }
                    get
                    { return row; }
                }

                /// <summary>
                /// 导出 Excel 到 DataSet
                /// </summary>
                /// <param name="stream"> Excel 路径 </param>
                /// <param name="index"> sheet index </param>
                /// <param name="header"> 是否数据标题行 </param>
                /// <param name="cellindex"> 根据列空白排除 </param>
                /// <returns>DataSet</returns>
                public static DataSet ToDataSet(string excel, int index, bool header, int cellindex)
                {
                    DataSet ds = new DataSet();
                    ds.Tables.Add(ToDataTable(excel, index, header, cellindex));
                    return ds;
                }
                /// <summary>
                /// 导出 Excel 到 DataSet
                /// </summary>
                /// <param name="stream"> Excel数据流 </param>
                /// <param name="index"> sheet index </param>
                /// <param name="header"> 是否数据标题行 </param>
                /// <param name="cellindex"> 根据列空白排除 </param>
                /// <returns>DataSet</returns>
                public static DataSet ToDataSet(FileStream stream, int index, bool header, int cellindex)
                {
                    DataSet ds = new DataSet();
                    ds.Tables.Add(ToDataTable(stream, index, header, cellindex));
                    return ds;
                }
                /// <summary>
                /// 导出 Excel 到 DataSet
                /// </summary>
                /// <param name="stream"> Excel路径 </param>
                /// <returns>DataSet</returns>
                public static DataSet ToDataSet(string excelpath)
                {

                    DataSet ds = new DataSet();
                    // is2007 = true;

                    /// office work
                    IWorkbook workbook;
                    ISheet sheet;
                    FileStream stream = new FileStream(excelpath, FileMode.Open, FileAccess.Read);
                    // 选择 不同 office版本 不同实例
                    //try
                    //{
                    //    workbook = new XSSFWorkbook(stream);
                    //}
                    //catch (Exception ex)
                    //{
                    //    workbook = new HSSFWorkbook(stream);
                    //    is2007 = false;
                    //}
                    workbook = WorkbookFactory.Create(stream);
                    if (workbook == null) return null;
                    //获取excel的总页数
                    int SheetCount = workbook.NumberOfSheets;
                    for (int index = 0; index < SheetCount; index++)
                    {
                        sheet = workbook.GetSheetAt(index);

                        DataTable dt = new DataTable(Path.GetFileNameWithoutExtension(stream.Name)
                           + sheet.SheetName);

                        var rows = sheet.GetRowEnumerator();
                        rows.MoveNext();
                        Row = rows.Current;
                        int cellcount = row.Cells.Count;

                        for (int i = 0; i < cellcount; i++)
                        {
                            ICell cell = row.GetCell(i);
                            string columnName = true ? cell.StringCellValue : i.ToString();

                            dt.Columns.Add(columnName, typeof(string));
                        }

                        if (!true)
                        {
                            DataRow first = dt.NewRow();
                            for (int i = 0; i < row.LastCellNum; i++)
                            {
                                ICell cell = row.GetCell(i);
                                first[i] = cell.StringCellValue.Trim().Replace("\0", "");
                            }
                            dt.Rows.Add(first);
                        }

                        // 开始按行导出
                        while (rows.MoveNext())
                        {
                            Row = rows.Current;

                            //// 根据 cellindex 列排除
                            //if (
                            //    row.GetCell(cellindex) == null ||
                            //    string.IsNullOrEmpty(row.GetCell(cellindex).StringCellValue)
                            //    )
                            //    continue;

                            DataRow dataRow = dt.NewRow();
                            for (int i = 0; i < row.LastCellNum; i++)
                            {
                                ICell cell = row.GetCell(i);

                                try
                                {
                                    dataRow[i] = cell.StringCellValue.Trim().Replace("\0", "");
                                }
                                catch
                                {
                                    try
                                    {
                                        dataRow[i] = cell.NumericCellValue;
                                    }
                                    catch
                                    { }
                                }
                            }
                            dt.Rows.Add(dataRow);
                        }
                        ds.Tables.Add(dt);

                    }

                    return ds;
                }
                #region 从excel中将数据导出到datatable
                /// <summary>读取excel
                /// 默认第一行为标头
                /// </summary>
                /// <param name="strFileName">excel文档路径</param>
                /// <returns></returns>
                public static DataTable ImportExceltoDt(string strFileName)
                {
                    DataTable dt = new DataTable();
                    IWorkbook wb;
                    using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                    {
                        wb = WorkbookFactory.Create(file);
                    }
                    ISheet sheet = wb.GetSheetAt(0);
                    dt = ImportDt(sheet, 0, true);
                    return dt;
                }

                /// <summary>
                /// 读取excel
                /// </summary>
                /// <param name="strFileName">excel文件路径</param>
                /// <param name="sheet">需要导出的sheet</param>
                /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
                /// <returns></returns>
                public static DataTable ImportExceltoDt(string strFileName, string SheetName, int HeaderRowIndex)
                {
                    HSSFWorkbook workbook;
                    IWorkbook wb;
                    using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                    {
                        wb = WorkbookFactory.Create(file);
                    }
                    ISheet sheet = wb.GetSheet(SheetName);
                    DataTable table = new DataTable();
                    table = ImportDt(sheet, HeaderRowIndex, true);
                    //ExcelFileStream.Close();
                    workbook = null;
                    sheet = null;
                    return table;
                }

                /// <summary>
                /// 读取excel
                /// </summary>
                /// <param name="strFileName">excel文件路径</param>
                /// <param name="sheet">需要导出的sheet序号</param>
                /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
                /// <returns></returns>
                public static DataTable ImportExceltoDt(string strFileName, int SheetIndex, int HeaderRowIndex,out string exMsg)
                {
                    try
                    {
                        HSSFWorkbook workbook;
                        IWorkbook wb;
                        using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                        {
                            wb = WorkbookFactory.Create(file);
                        }
                        ISheet isheet = wb.GetSheetAt(SheetIndex);
                        DataTable table = new DataTable();
                        table = ImportDt(isheet, HeaderRowIndex, true);
                        //ExcelFileStream.Close();
                        workbook = null;
                        isheet = null;
                        exMsg = string.Empty;
                        return table;
                    }
                    catch (EncryptedDocumentException ex)
                    {
                        exMsg = "Sorry, the system does not support confidential documents, please delete the password and try again.";
                        return null;
                    }
                   
                }

                /// <summary>
                /// 读取excel
                /// </summary>
                /// <param name="strFileName">excel文件路径</param>
                /// <param name="sheet">需要导出的sheet</param>
                /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
                /// <returns></returns>
                public static DataTable ImportExceltoDt(string strFileName, string SheetName, int HeaderRowIndex, bool needHeader)
                {
                    HSSFWorkbook workbook;
                    IWorkbook wb;
                    using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                    {
                        wb = WorkbookFactory.Create(file);
                    }
                    ISheet sheet = wb.GetSheet(SheetName);
                    DataTable table = new DataTable();
                    table = ImportDt(sheet, HeaderRowIndex, needHeader);
                    //ExcelFileStream.Close();
                    workbook = null;
                    sheet = null;
                    return table;
                }

                /// <summary>
                /// 读取excel
                /// </summary>
                /// <param name="strFileName">excel文件路径</param>
                /// <param name="sheet">需要导出的sheet序号</param>
                /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
                /// <returns></returns>
                public static DataTable ImportExceltoDt(string strFileName, int SheetIndex, int HeaderRowIndex, bool needHeader)
                {
                    HSSFWorkbook workbook;
                    IWorkbook wb;
                    using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                    {
                        wb = WorkbookFactory.Create(file);
                    }
                    ISheet sheet = wb.GetSheetAt(SheetIndex);
                    DataTable table = new DataTable();
                    table = ImportDt(sheet, HeaderRowIndex, needHeader);
                    //ExcelFileStream.Close();
                    workbook = null;
                    sheet = null;
                    return table;
                }

                /// <summary>
                /// 将制定sheet中的数据导出到datatable中
                /// </summary>
                /// <param name="sheet">需要导出的sheet</param>
                /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
                /// <returns></returns>
                static DataTable ImportDt(ISheet sheet, int HeaderRowIndex, bool needHeader)
                {
                    DataTable table = new DataTable();
                    IRow headerRow;
                    int cellCount;
                    try
                    {
                        if (HeaderRowIndex < 0 || !needHeader)
                        {
                            headerRow = sheet.GetRow(0);
                            cellCount = headerRow.LastCellNum;

                            for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                            {
                                DataColumn column = new DataColumn(Convert.ToString(i));
                                table.Columns.Add(column);
                            }
                        }
                        else
                        {
                            headerRow = sheet.GetRow(HeaderRowIndex);
                            cellCount = headerRow.LastCellNum;

                            for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                            {
                                if (headerRow.GetCell(i) == null)
                                {
                                    if (table.Columns.IndexOf(Convert.ToString(i)) > 0)
                                    {
                                        DataColumn column = new DataColumn(Convert.ToString("RepeatColumn" + i));
                                        table.Columns.Add(column);
                                    }
                                    else
                                    {
                                        DataColumn column = new DataColumn(Convert.ToString(i));
                                        table.Columns.Add(column);
                                    }

                                }
                                else if (table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0)
                                {
                                    DataColumn column = new DataColumn(Convert.ToString("RepeatColumn" + i));
                                    table.Columns.Add(column);
                                }
                                else
                                {
                                    var cell = headerRow.GetCell(i);
                                    DataColumn column = null;
                                    if (cell.CellType == CellType.String)
                                    {
                                        column = new DataColumn(headerRow.GetCell(i).ToString());

                                    }
                                    else if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
                                    {
                                        column = new DataColumn(headerRow.GetCell(i).NumericCellValue.ToString());
                                    }
                                    else
                                    {
                                        column = new DataColumn("undefined" + i);
                                    }
                                    table.Columns.Add(column);
                                }
                            }
                        }
                        int rowCount = sheet.LastRowNum;
                        for (int i = (HeaderRowIndex + 1); i <= sheet.LastRowNum; i++)
                        {
                            try
                            {
                                IRow row;
                                if (sheet.GetRow(i) == null)
                                {
                                    row = sheet.CreateRow(i);
                                }
                                else
                                {
                                    row = sheet.GetRow(i);
                                }

                                DataRow dataRow = table.NewRow();

                                for (int j = row.FirstCellNum; j <= cellCount; j++)
                                {
                                    try
                                    {
                                        if (row.GetCell(j) != null)
                                        {
                                            switch (row.GetCell(j).CellType)
                                            {
                                                case CellType.String:
                                                    string str = row.GetCell(j).StringCellValue;
                                                    if (str != null && str.Length > 0)
                                                    {
                                                        dataRow[j] = str.ToString();
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = null;
                                                    }
                                                    break;
                                                case CellType.Numeric:
                                                    if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                                    {
                                                        dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                                    }
                                                    break;
                                                case CellType.Boolean:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                    break;
                                                case CellType.Error:
                                                    dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                    break;
                                                case CellType.Formula:
                                                    switch (row.GetCell(j).CachedFormulaResultType)
                                                    {
                                                        case CellType.String:
                                                            string strFORMULA = row.GetCell(j).StringCellValue;
                                                            if (strFORMULA != null && strFORMULA.Length > 0)
                                                            {
                                                                dataRow[j] = strFORMULA.ToString();
                                                            }
                                                            else
                                                            {
                                                                dataRow[j] = null;
                                                            }
                                                            break;
                                                        case CellType.Numeric:
                                                            dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                            break;
                                                        case CellType.Boolean:
                                                            dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                            break;
                                                        case CellType.Error:
                                                            dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                            break;
                                                        default:
                                                            dataRow[j] = "";
                                                            break;
                                                    }
                                                    break;
                                                default:
                                                    dataRow[j] = "";
                                                    break;
                                            }
                                        }
                                    }
                                    catch (Exception exception)
                                    {
                                        //wl.WriteLogs(exception.ToString());
                                    }
                                }
                                table.Rows.Add(dataRow);
                            }
                            catch (Exception exception)
                            {
                                //wl.WriteLogs(exception.ToString());
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        // wl.WriteLogs(exception.ToString());
                    }
                    return table;
                }

                #endregion 
                /// <summary>
                /// 导出 Excel 到 DataTable
                /// </summary>
                /// <param name="excel"> Excel 路径 </param>
                /// <param name="index"> sheet index </param>
                /// <param name="header"> 是否数据标题行 </param>
                /// <param name="cellindex"> 根据列空白排除 </param>
                /// <returns>DataTable</returns>
                public static DataTable ToDataTable(string excel, int index, bool header, int cellindex)
                {
                    using (FileStream file = new FileStream(excel, FileMode.Open, FileAccess.Read))
                    {
                        return ToDataTable(file, index, header, cellindex);
                    }
                }
                //<summary>
                //导出 Excel 到 DataTable
                //</summary>
                //<param name="stream"> Excel数据流 </param>
                //<param name="index"> sheet index </param>
                //<param name="header"> 是否数据标题行 </param>
                //<param name="cellindex"> 根据列空白排除 </param>
                //<returns>DataTable</returns>
                public static DataTable ToDataTable(FileStream stream, int index, bool header, int cellindex)
                {
                    is2007 = true;
                    DataTable dt;
                    /// office work
                    IWorkbook workbook;
                    ISheet sheet;

                    // 选择 不同 office版本 不同实例
                    try
                    {
                        workbook = new XSSFWorkbook(stream);
                    }
                    catch (Exception ex)
                    {
                        workbook = new HSSFWorkbook(stream);
                        is2007 = false;
                    }
                    if (workbook == null) return null;
                    int SheetCount = workbook.NumberOfSheets;
                    if (SheetCount - 1 < index)
                        throw new Exception("Excel表项有误！");
                    sheet = workbook.GetSheetAt(index);

                    dt = new DataTable(Path.GetFileNameWithoutExtension(stream.Name)
                        + sheet.SheetName);

                    var rows = sheet.GetRowEnumerator();
                    rows.MoveNext();
                    Row = rows.Current;
                    int cellcount = row.Cells.Count;

                    for (int i = 0; i < cellcount; i++)
                    {
                        ICell cell = row.GetCell(i);
                        string columnName = header ? i.ToString() : cell.StringCellValue;

                        dt.Columns.Add(columnName, typeof(string));
                    }

                    if (header)
                    {
                        DataRow first = dt.NewRow();
                        for (int i = 0; i < row.LastCellNum; i++)
                        {
                            ICell cell = row.GetCell(i);
                            first[i] = cell.StringCellValue.Trim().Replace("\0", "");
                        }
                        dt.Rows.Add(first);
                    }

                    // 开始按行导出
                    while (rows.MoveNext())
                    {
                        Row = rows.Current;

                        // 根据 cellindex 列排除
                        if (
                            row.GetCell(cellindex) == null ||
                            string.IsNullOrEmpty(row.GetCell(cellindex).StringCellValue)
                            )
                            continue;

                        DataRow dataRow = dt.NewRow();
                        for (int i = 0; i < row.LastCellNum; i++)
                        {
                            ICell cell = row.GetCell(i);

                            try
                            {
                                dataRow[i] = cell.StringCellValue.Trim().Replace("\0", "");
                            }
                            catch
                            {
                                try
                                {
                                    dataRow[i] = cell.NumericCellValue;
                                }
                                catch
                                { }
                            }
                        }
                        dt.Rows.Add(dataRow);
                    }

                    return dt;
                }
            }
            #endregion

            #region 特用版本 导出Excel
            /// <summary>
            /// 导出Excel
            /// </summary>
            public class ExportExcel
            {

                /// <summary>
                /// DataTable数据导出到Excel中
                /// </summary>
                /// <param name="dt">数据表</param>
                /// <param name="path">保存文件路径</param>
                /// <param name="listRemoveCol">所要删除的列名</param>
                public static void ToExcel(DataTable dt, string path, List<string> listRemoveCol)
                {
                    try
                    {
                        MemoryStream ms = new MemoryStream();
                        NPOI.SS.UserModel.IWorkbook npoi_workbook = new HSSFWorkbook();
                        NPOI.SS.UserModel.ISheet npoi_worksheet = npoi_workbook.CreateSheet("Sheet1");
                        NPOI.SS.UserModel.IRow npoi_workRow = npoi_worksheet.CreateRow(0);
                        NPOI.SS.UserModel.ICell npoi_workcell;

                        if (dt != null)
                        {
                            if (listRemoveCol != null)
                            {
                                foreach (string strColName in listRemoveCol)
                                {
                                    dt.Columns.Remove(strColName);
                                }
                            }
                            #region 表头
                            ICellStyle cellstyle = null;
                            //short iMaxLen = 4;
                            //short iTmpMax = 4;
                            int iColIndex = 0;
                            foreach (DataColumn column in dt.Columns)
                            {
                                //int iLen = System.Text.ASCIIEncoding.Default.GetByteCount(column.ColumnName);
                                //iTmpMax = (short)(iLen / 2);
                                //if (iTmpMax > iMaxLen)
                                //{
                                //    iMaxLen = iTmpMax;
                                //    iMaxLen = (short)(40 * iMaxLen / 4);
                                //}
                                //npoi_workRow.HeightInPoints = iMaxLen;

                                npoi_workcell = npoi_workRow.CreateCell(iColIndex);
                                npoi_workcell.CellStyle = cellstyle;
                                npoi_workcell.SetCellValue(column.ColumnName);

                                iColIndex++;
                            }
                            #endregion

                            #region 表数据
                            int iRowIndex = 1;
                            cellstyle = null;
                            foreach (DataRow dr in dt.Rows)
                            {
                                npoi_workRow = npoi_worksheet.CreateRow(iRowIndex++);
                                iColIndex = 0;
                                foreach (DataColumn column in dt.Columns)
                                {
                                    string strValue = dr[column.ColumnName].ToString();
                                    npoi_workcell = npoi_workRow.CreateCell(iColIndex++);
                                    npoi_workcell.CellStyle = cellstyle;
                                    npoi_workcell.SetCellValue("\t" + strValue + "\t");
                                }
                            }
                            #endregion
                        }
                        npoi_workbook.Write(ms);
                        File.WriteAllBytes(path, ms.ToArray());
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message, "错误提示");
                    }
                }

            }
            #endregion
        }

        /// <summary>
        /// Word 操作类
        /// </summary>
        public class NPOI_Word
        { }
    }


    /*
                try
                {
                    string strIput = "";
                    int IcountSta = 0;
                    for (int i = 0; i < dgv1.Rows.Count; i++)
                    {
                        strIput = dgv1.Rows[i].Cells["HB11"].Value.ToString();
                        strIput = strIput.Replace(",", "，");
                        string[] strSplit = strIput.Split('，');

                        if (strSplit.Count() > IcountSta)
                        {
                            IcountSta = strSplit.Count();
                        }
                    }

                    int IcountA = IcountSta + 4;
                    worksheet.Cells[1, 1] = "审批表";
                    worksheet.Cells[1, 2] = "公告编号";
                    worksheet.Cells[1, 3] = "披露日期";
                    worksheet.Cells[1, 4] = "公告名称及内容";
                    for (int i = 5; i <= IcountA; i++)
                    {
                        worksheet.Cells[1, i] = " ";
                    }
                    Microsoft.Office.Interop.Excel.Range range1 = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 4]];
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.Font.Size = 10;
                    range1.Font.Bold = true;
                    range1.Cells.Interior.Color = System.Drawing.Color.LightSkyBlue;
                    int icount = dgv1.Rows.Count;
                    Microsoft.Office.Interop.Excel.Range excelrange = worksheet.Range[worksheet.Cells[2, IcountA], worksheet.Cells[icount + 1, IcountA]];
                    Microsoft.Office.Interop.Excel.Range range;
                    Microsoft.Office.Interop.Excel.Range range2;
                    Microsoft.Office.Interop.Excel.Range range3;
                    range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[icount + 1, IcountA]];
                    //range.Cells.Borders.LineStyle = 1;
                    for (int i = 0; i < icount; i++)
                    {
                        if (dgv1.Rows[i].Cells["HB1"].Value.ToString() != "" && dgv1.Rows[i].Cells["HB1"].Value.ToString() != null)
                        {
                            range3 = worksheet.Range[worksheet.Cells[i + 2, 1], worksheet.Cells[i + 2, 4]];
                            range3.Interior.ColorIndex = 6;
                            range3.Font.Size = 10;
                        }
                        range2 = worksheet.Range[worksheet.Cells[i + 2, 1], worksheet.Cells[i + 2, IcountA]];
                        range2.Font.Size = 10;
                        range2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                        worksheet.Cells[i + 2, 1] = dgv1.Rows[i].Cells["HB1"].Value.ToString();
                        worksheet.Cells[i + 2, 2] = dgv1.Rows[i].Cells["HB2"].Value.ToString();
                        worksheet.Cells[i + 2, 3] = dgv1.Rows[i].Cells["HB3"].Value.ToString();
                        worksheet.Cells[i + 2, 4] = dgv1.Rows[i].Cells["HB4"].Value.ToString();
                        if (dgv1.Rows[i].Cells["HB11"].Value == null)
                        {
                            dgv1.Rows[i].Cells["HB11"].Value = "";
                        }

                        if (dgv1.Rows[i].Cells["HB11"].Value != null && dgv1.Rows[i].Cells["HB11"].Value.ToString().Trim() != "")
                        {
                            strIput = dgv1.Rows[i].Cells["HB11"].Value.ToString();
                            strIput = strIput.Replace(",", "，");
                            string[] strSplit = strIput.Split('，');

                            for (int j = 5; j <= IcountA; j++)
                            {
                                try
                                {
                                    worksheet.Cells[i + 2, j] = strSplit[j - 5];
                                }
                                catch
                                {
                                    worksheet.Cells[i + 2, j] = "";
                                }
                            }

                        }
                    }
                    worksheet.Columns.EntireColumn.AutoFit();
                    worksheet.Rows.EntireRow.AutoFit();
                    workbook.Saved = true;
                    workbook.SaveCopyAs(localFilePath);

     * */


    //Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
    //Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
    //Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
    //Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1 
    //try
    //{
    //    ////定义Range对象,此对象代表单元格区域 
    //    //Microsoft.Office.Interop.Excel.Range range;
    //    worksheet.Cells.WrapText = true;
    //    #region 表头
    //    worksheet.Columns.NumberFormatLocal = "@";
    //    worksheet.Cells[1, 1] = "议案";
    //    Microsoft.Office.Interop.Excel.Range range11 = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 1]];
    //    range11.ColumnWidth = 50;
    //    worksheet.Cells[1, 2] = "参加表决的有效票数";
    //    range11 = worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[1, 10]];
    //    range11.ColumnWidth = 20;
    //    worksheet.Cells[1, 3] = "同意";
    //    worksheet.Cells[1, 4] = "占比";
    //    worksheet.Cells[1, 5] = "反对";
    //    worksheet.Cells[1, 6] = "占比";
    //    worksheet.Cells[1, 7] = "弃权";
    //    worksheet.Cells[1, 8] = "占比";
    //    worksheet.Cells[1, 9] = "回避";
    //    worksheet.Cells[1, 10] = "占比";
    //    Microsoft.Office.Interop.Excel.Range range1 = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 10]];
    //    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
    //    range1.Font.Size = 14;
    //    range1.Font.Bold = true;
    //    #endregion
    //    SIS.BLL.FillBills bill = new SIS.BLL.FillBills();
    //    List<SIS.Model.FillBills> modelA = bill.GetModelList("D_ID='" + iID + "'", 1).Where(x => x.B_Type != 2 && x.B_Type != 6).ToList();
    //    int icountA = 1;

    //    Microsoft.Office.Interop.Excel.Range range;
    //    range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[modelA.Count() + 1, 10]];
    //    range.Cells.Borders.LineStyle = 2;

    //    modelAD = _GDDH.GetStockXCTP(iID).Where(x=>x.StockTa ==1).ToList();
    //    if (Common.HY.GetSX() == "深市")
    //    {
    //        #region 深市
    //        modelAB = _GDDH.GetStockSSWT(iID);
    //        foreach (SIS.Model.FillBills list in modelA.Where(x => x.B_Type == 1).OrderBy(s => s.B_No))
    //        {
    //            int icountAS = 1;
    //            List<SIS.BLL.FillBill> listA = modelAll.Where(x => x.B_BH == list.B_ID).ToList();
    //            if (listA.Count == 0)
    //            {
    //                Microsoft.Office.Interop.Excel.Range excelrange = worksheet.Range[worksheet.Cells[icountM, 1], worksheet.Cells[icountM, 10]];
    //                excelrange.Merge(excelrange.MergeCells);
    //                worksheet.Cells[icountM, 1] = "议案" + icountA.ToString() + ":" + list.B_Name;
    //                icountM++;
    //                foreach (SIS.Model.FillBills listA1 in modelA.Where(x => x.B_ParentID == list.B_ID).OrderBy(s => s.B_No))
    //                {
    //                    worksheet.Cells[icountM, 1] = "(" + icountAS.ToString() + ")" + listA1.B_Name;
    //                    AddExcel(worksheet, listA1, modelAll.Where(x => x.B_BH == listA1.B_ID).First());
    //                    icountAS++;
    //                    icountM++;
    //                }
    //            }
    //            else
    //            {
    //                worksheet.Cells[icountM, 1] = "议案" + icountA.ToString() + ":" + list.B_Name;
    //                AddExcel(worksheet, list, listA[0]);
    //                icountM++;
    //            }

    //            icountA++;
    //        }
    //        #endregion
    //    }
    //    else
    //    {
    //        #region 沪市
    //        modelAC = _GDDH.GetStockHSWT(iID);
    //        foreach (SIS.Model.FillBills list in modelA.Where(x => x.B_Type == 1).OrderBy(s => s.B_No))
    //        {
    //            List<SIS.BLL.FillBill> listA = modelAll.Where(x => x.B_BH == list.B_ID).ToList();
    //            if (listA.Count == 0)
    //            {
    //                Microsoft.Office.Interop.Excel.Range excelrange = worksheet.Range[worksheet.Cells[icountM, 1], worksheet.Cells[icountM, 10]];
    //                excelrange.Merge(excelrange.MergeCells);
    //                worksheet.Cells[icountM, 1] = "议案" + icountA.ToString() + ":" + list.B_Name;
    //                icountM++;
    //                int icountA1 = 1;
    //                foreach (SIS.Model.FillBills listA1 in modelA.Where(x => x.B_ParentID == list.B_ID).OrderBy(s => s.B_No))
    //                {
    //                    List<SIS.BLL.FillBill> listB = modelAll.Where(x => x.B_BH == listA1.B_ID).ToList();
    //                    if (listB.Count == 0)
    //                    {
    //                        Microsoft.Office.Interop.Excel.Range excelrange1 = worksheet.Range[worksheet.Cells[icountM, 1], worksheet.Cells[icountM, 10]];
    //                        excelrange1.Merge(excelrange1.MergeCells);
    //                        worksheet.Cells[icountM, 1] = icountA.ToString() + "." + icountA1.ToString() + listA1.B_Name;
    //                        icountM++;
    //                        SetExcel(modelA, listA1.B_ID, worksheet);
    //                    }
    //                    else
    //                    {
    //                        worksheet.Cells[icountM, 1] = icountA.ToString() + "." + icountA1.ToString() + listA1.B_Name;
    //                        AddExcels(worksheet, listA1, listB[0]);
    //                        icountM++;
    //                    }
    //                    icountA1++;
    //                }
    //            }
    //            else
    //            {
    //                worksheet.Cells[icountM, 1] = "议案" + icountA.ToString() + ":" + list.B_Name;
    //                AddExcels(worksheet, list, listA[0]);
    //                icountM++;
    //            }

    //            icountA++;
    //        }
    //        #endregion
    //    }
    //    worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。
    //    worksheet.Rows.EntireRow.AutoFit();//行宽自适应
    //    workbook.Saved = true;
    //    workbook.SaveCopyAs(localFilePath);
    //    //linkLabel1.Enabled = true;
    //}
    //catch
    //{
    //    MessageBox.Show("投票情况汇总生成失败。", "提示");
    //}
    //finally
    //{
    //    xlApp.Quit();
    //    GC.Collect();//强行销毁 
    //}


    //private void SetExcel(List<SIS.Model.FillBills> modelA, string strID, Microsoft.Office.Interop.Excel.Worksheet worksheet)
    //{
    //    foreach (SIS.Model.FillBills list in modelA.Where(x => x.B_ParentID == strID))
    //    {
    //        List<SIS.BLL.FillBill> listA = modelAll.Where(x => x.B_BH == list.B_ID).ToList();
    //        if (listA.Count == 0)
    //        {
    //            Microsoft.Office.Interop.Excel.Range excelrange = worksheet.Range[worksheet.Cells[icountM, 1], worksheet.Cells[icountM, 10]];
    //            excelrange.Merge(excelrange.MergeCells);
    //            worksheet.Cells[icountM, 1] = list.B_Name;
    //            icountM++;
    //            SetExcel(modelA, list.B_ID, worksheet);
    //        }
    //        else
    //        {
    //            worksheet.Cells[icountM, 1] = "(" + icountN.ToString() + ")" + list.B_Name;
    //            AddExcels(worksheet, list, listA[0]);
    //            icountM++;
    //            icountN++;
    //        }
    //    }
    //}
    ////添加沪市
    //private void AddExcels(Microsoft.Office.Interop.Excel.Worksheet worksheet, SIS.Model.FillBills list, SIS.BLL.FillBill listA)
    //{
    //    List<SIS.Model.StockXCTP> modeAD = modelAD.Where(x => x.B_ID == list.B_ID).ToList();
    //    List<SIS.Model.StockHSWT> modeAB = modelAC.Where(x => x.B_ID == listA.B_BM).ToList();
    //    if (modeAB.Count > 0)
    //    {
    //        if (list.B_Tag.Value == 3)
    //        {
    //            Microsoft.Office.Interop.Excel.Range excelrange = worksheet.Range[worksheet.Cells[icountM, 2], worksheet.Cells[icountM, 10]];
    //            excelrange.Merge(excelrange.MergeCells);
    //            decimal iZS = modeAD.Sum(x => x.StockNum.Value);
    //            worksheet.Cells[icountM, 2] = iZS.ToString("f0") + "票";
    //        }
    //        else
    //        {
    //            decimal iZS = modeAD.Sum(x => x.HaveNum).Value + modeAB[0].StockTYGQZH.Value + modeAB[0].StockFDGQZH.Value + modeAB[0].StockQQGQZH.Value;
    //            decimal iTY = modeAD.Where(x => x.StockNum == 0).Sum(x => x.HaveNum).Value + modeAB[0].StockTYGQZH.Value;
    //            decimal iFD = modeAD.Where(x => x.StockNum == 1).Sum(x => x.HaveNum).Value + modeAB[0].StockFDGQZH.Value;
    //            decimal iQQ = modeAD.Where(x => x.StockNum == 2).Sum(x => x.HaveNum).Value + modeAB[0].StockQQGQZH.Value;
    //            decimal iHB = modeAD.Where(x => x.StockNum == 3).Sum(x => x.HaveNum).Value;
    //            decimal iYSZS = iTY + iFD + iQQ;
    //            worksheet.Cells[icountM, 2] = iYSZS.ToString("f0");
    //            worksheet.Cells[icountM, 3] = iTY.ToString("f0");
    //            if (iYSZS > 0)
    //            {
    //                worksheet.Cells[icountM, 4] = (iTY / iYSZS * 100).ToString("f4") + "%";
    //                worksheet.Cells[icountM, 5] = iFD.ToString("f0");
    //                worksheet.Cells[icountM, 6] = (iFD / iYSZS * 100).ToString("f4") + "%";
    //                worksheet.Cells[icountM, 7] = iQQ.ToString("f0");
    //                worksheet.Cells[icountM, 8] = (iQQ / iYSZS * 100).ToString("f4") + "%";
    //                worksheet.Cells[icountM, 9] = iHB.ToString("f0");
    //            }
    //            else
    //            {
    //                worksheet.Cells[icountM, 4] = 0 + "%";
    //                worksheet.Cells[icountM, 5] = iFD.ToString("f0");
    //                worksheet.Cells[icountM, 6] = 0 + "%";
    //                worksheet.Cells[icountM, 7] = iQQ.ToString("f0");
    //                worksheet.Cells[icountM, 8] = 0 + "%";
    //                worksheet.Cells[icountM, 9] = iHB.ToString("f0");
    //            }
    //            if (iZS > 0)
    //            {
    //                worksheet.Cells[icountM, 10] = (iHB / iZS * 100).ToString("f4") + "%";
    //            }
    //            else
    //            {
    //                worksheet.Cells[icountM, 10] = 0 + "%";
    //            }
    //        }
    //    }
    //    else
    //    {
    //        if (list.B_Tag.Value == 3)
    //        {
    //            Microsoft.Office.Interop.Excel.Range excelrange = worksheet.Range[worksheet.Cells[icountM, 2], worksheet.Cells[icountM, 10]];
    //            excelrange.Merge(excelrange.MergeCells);
    //            decimal iZS = modeAD.Sum(x => x.StockNum.Value);
    //            worksheet.Cells[icountM, 2] = iZS.ToString("f0") + "票";
    //        }
    //        else
    //        {
    //            decimal iZS = modeAD.Sum(x => x.HaveNum).Value;
    //            decimal iTY = modeAD.Where(x => x.StockNum == 0).Sum(x => x.HaveNum).Value;
    //            decimal iFD = modeAD.Where(x => x.StockNum == 1).Sum(x => x.HaveNum).Value;
    //            decimal iQQ = modeAD.Where(x => x.StockNum == 2).Sum(x => x.HaveNum).Value;
    //            decimal iHB = modeAD.Where(x => x.StockNum == 3).Sum(x => x.HaveNum).Value;
    //            decimal iYSZS = iTY + iFD + iQQ;
    //            worksheet.Cells[icountM, 2] = iYSZS.ToString("f0");
    //            worksheet.Cells[icountM, 3] = iTY.ToString("f0");
    //            if (iYSZS > 0)
    //            {
    //                worksheet.Cells[icountM, 4] = (iTY / iYSZS * 100).ToString("f4") + "%";
    //                worksheet.Cells[icountM, 5] = iFD.ToString("f0");
    //                worksheet.Cells[icountM, 6] = (iFD / iYSZS * 100).ToString("f4") + "%";
    //                worksheet.Cells[icountM, 7] = iQQ.ToString("f0");
    //                worksheet.Cells[icountM, 8] = (iQQ / iYSZS * 100).ToString("f4") + "%";
    //                worksheet.Cells[icountM, 9] = iHB.ToString("f0");
    //            }
    //            else
    //            {
    //                worksheet.Cells[icountM, 4] = 0 + "%";
    //                worksheet.Cells[icountM, 5] = iFD.ToString("f0");
    //                worksheet.Cells[icountM, 6] = 0 + "%";
    //                worksheet.Cells[icountM, 7] = iQQ.ToString("f0");
    //                worksheet.Cells[icountM, 8] = 0 + "%";
    //                worksheet.Cells[icountM, 9] = iHB.ToString("f0");
    //            }
    //            if (iZS > 0)
    //            {
    //                worksheet.Cells[icountM, 10] = (iHB / iZS * 100).ToString("f4") + "%";
    //            }
    //            else
    //            {
    //                worksheet.Cells[icountM, 10] = 0 + "%";
    //            }
    //        }
    //    }
    //}
    ////添加深市
    //private void AddExcel(Microsoft.Office.Interop.Excel.Worksheet worksheet, SIS.Model.FillBills list, SIS.BLL.FillBill listA)
    //{
    //    List<SIS.Model.StockXCTP> modeAD = modelAD.Where(x => x.B_ID == list.B_ID).ToList();
    //    List<SIS.Model.StockSSWT> modeAB = modelAB.Where(x => x.B_ID == listA.B_BM).ToList();
    //    if (list.B_Tag.Value == 3)
    //    {
    //        Microsoft.Office.Interop.Excel.Range excelrange = worksheet.Range[worksheet.Cells[icountM, 2], worksheet.Cells[icountM, 10]];
    //        excelrange.Merge(excelrange.MergeCells);
    //        decimal iZS = modeAD.Sum(x => x.StockNum.Value);
    //        worksheet.Cells[icountM, 2] = iZS.ToString("f0") + "票";
    //    }
    //    else
    //    {
    //        decimal iZS = modeAD.Sum(x => x.HaveNum).Value ;
    //        decimal iTY = modeAD.Where(x => x.StockNum == 0).Sum(x => x.HaveNum).Value;
    //        decimal iFD = modeAD.Where(x => x.StockNum == 1).Sum(x => x.HaveNum).Value;
    //        decimal iQQ = modeAD.Where(x => x.StockNum == 2).Sum(x => x.HaveNum).Value;
    //        decimal iHB = modeAD.Where(x => x.StockNum == 3).Sum(x => x.HaveNum).Value;

    //        iZS += modeAB.Sum(x => x.HaveNum).Value;
    //        iTY += modeAB.Where(x => x.StockValues == "同意").Sum(x => x.HaveNum).Value;
    //        iFD += modeAB.Where(x => x.StockValues == "反对").Sum(x => x.HaveNum).Value;
    //        iQQ += modeAB.Where(x => x.StockValues == "弃权").Sum(x => x.HaveNum).Value;
    //        iHB += modeAB.Where(x => x.StockValues == "回避").Sum(x => x.HaveNum).Value;




    //        decimal iYSZS = iTY + iFD + iQQ;
    //        worksheet.Cells[icountM, 2] = iYSZS.ToString("f0");
    //        worksheet.Cells[icountM, 3] = iTY.ToString("f0");
    //        if (iYSZS > 0)
    //        {
    //            worksheet.Cells[icountM, 4] = (iTY / iYSZS * 100).ToString("f4") + "%";
    //            worksheet.Cells[icountM, 5] = iFD.ToString("f0");
    //            worksheet.Cells[icountM, 6] = (iFD / iYSZS * 100).ToString("f4") + "%";
    //            worksheet.Cells[icountM, 7] = iQQ.ToString("f0");
    //            worksheet.Cells[icountM, 8] = (iQQ / iYSZS * 100).ToString("f4") + "%";
    //            worksheet.Cells[icountM, 9] = iHB.ToString("f0");
    //        }
    //        else
    //        {
    //            worksheet.Cells[icountM, 4] = 0 + "%";
    //            worksheet.Cells[icountM, 5] = iFD.ToString("f0");
    //            worksheet.Cells[icountM, 6] = 0 + "%";
    //            worksheet.Cells[icountM, 7] = iQQ.ToString("f0");
    //            worksheet.Cells[icountM, 8] = 0 + "%";
    //            worksheet.Cells[icountM, 9] = iHB.ToString("f0");
    //        }
    //        if (iZS > 0)
    //        {
    //            worksheet.Cells[icountM, 10] = (iHB / iZS * 100).ToString("f4") + "%";
    //        }
    //        else
    //        {
    //            worksheet.Cells[icountM, 10] = 0 + "%";
    //        }
    //    }
    //}
}