using System;



namespace R5T.Excel
{
    public static class WorkbookExtensions
    {
        public static Worksheet NewWorksheet(this Workbook workbook, string name)
        {
            var worksheet = workbook.NewWorksheet();

            worksheet.Name = name;

            return worksheet;
        }

        public static void DeleteWorksheet(this Workbook workbook, string name)
        {
            var worksheet = workbook.GetWorksheet(name);

            worksheet.Delete();
        }
    }
}
