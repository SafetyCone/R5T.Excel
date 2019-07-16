using System;


namespace R5T.Excel
{
    public static class WorksheetExtensions
    {
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
