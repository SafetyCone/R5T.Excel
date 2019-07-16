using System;

using R5T.NetStandard;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.Excel
{
    public static class ExcelFileFormatExtensions
    {
        internal static Xl.XlFileFormat ToXlFileFormat(this ExcelFileFormat excelFileFormat)
        {
            switch(excelFileFormat)
            {
                case ExcelFileFormat.CSV:
                    return Xl.XlFileFormat.xlCSV;

                case ExcelFileFormat.XLS:
                    return Xl.XlFileFormat.xlExcel8;

                case ExcelFileFormat.XLSM:
                    return Xl.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;

                case ExcelFileFormat.XLSX:
                default:
                    return Xl.XlFileFormat.xlOpenXMLWorkbook;
            }
        }
    }
}
