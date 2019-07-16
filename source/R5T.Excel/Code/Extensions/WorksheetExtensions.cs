using System;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.Excel
{
    public static class WorksheetExtensions
    {
        public static void Show(this Worksheet worksheet)
        {
            worksheet.XlWorksheet.Visible = Xl.XlSheetVisibility.xlSheetVisible;
        }

        public static void Hide(this Worksheet worksheet)
        {
            worksheet.XlWorksheet.Visible = Xl.XlSheetVisibility.xlSheetHidden;
        }

        public static void HideVeryHidden(this Worksheet worksheet)
        {
            worksheet.XlWorksheet.Visible = Xl.XlSheetVisibility.xlSheetVeryHidden;
        }

        public static void SetColumnWidths(this Worksheet worksheet, params double[] columnWidths)
        {
            var range = worksheet.GetA1Range();
            foreach (var columnWidth in columnWidths)
            {
                range.ColumnWidth = columnWidth;

                range = range.GetOffset(0, 1);
            }
        }
    }
}
