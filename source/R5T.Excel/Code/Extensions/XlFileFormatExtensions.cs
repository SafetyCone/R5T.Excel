using System;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.Excel
{
    public static class XlFileFormatExtensions
    {
        internal static ExcelFileFormat ToExcelFileFormat(this Xl.XlFileFormat xlFileFormat)
        {
            switch(xlFileFormat)
            {
                case Xl.XlFileFormat.xlCSV:
                    return ExcelFileFormat.CSV;

                case Xl.XlFileFormat.xlExcel8:
                    return ExcelFileFormat.XLS;

                case Xl.XlFileFormat.xlOpenXMLWorkbookMacroEnabled:
                    return ExcelFileFormat.XLSM;

                case Xl.XlFileFormat.xlOpenXMLWorkbook:
                    return ExcelFileFormat.XLSX;

                default:
                    return ExcelFileFormat.Other;
            }
        }
    }
}
