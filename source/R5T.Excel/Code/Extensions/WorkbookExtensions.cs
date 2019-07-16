using System;

using Xl = Microsoft.Office.Interop.Excel;


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

        public static bool HasWorksheet(this Workbook workbook, string name)
        {
            var output = false;
            foreach (Xl.Worksheet worksheet in workbook.XlWorkbook.Worksheets)
            {
                if (name == worksheet.Name)
                {
                    output = true;
                    break;
                }
            }

            return output;
        }

        public static void DeleteWorksheet(this Workbook workbook, string name)
        {
            var worksheet = workbook.GetWorksheet(name);

            worksheet.Delete();
        }

        public static void AddNamedRange(this Workbook workbook, Range range, string name)
        {
            workbook.XlWorkbook.Names.Add(name, range.XlRange);
        }

        public static bool HasNamedRange(this Workbook workbook, string name)
        {
            foreach (Xl.Name xlName in workbook.XlWorkbook.Names)
            {
                if(name == xlName.Name)
                {
                    return true;
                }
            }

            return false;
        }

        public static Range GetNamedRange(this Workbook workbook, string name)
        {

        }
    }
}
