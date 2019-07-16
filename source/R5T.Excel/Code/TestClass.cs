using System;

using XL = Microsoft.Office.Interop.Excel;


namespace R5T.Excel
{
    public class TestClass2
    {
        public static string GetTestValue()
        {
            var testWorkbookFilePath = @"C:\Temp\Book2.xlsx";
            var rangeAddress = "B4";

            var value = TestClass2.GetTestValue(testWorkbookFilePath, rangeAddress);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return value;
        }

        private static string GetTestValue(string testWorkbookFilePath, string rangeAddress)
        {
            var excelApp = new XL.Application
            {
                Visible = true
            };

            var workbook = excelApp.Workbooks.Open(testWorkbookFilePath);

            var worksheet = workbook.Worksheets["Sheet1"] as XL.Worksheet;

            var range = worksheet.Range[rangeAddress];

            var value = range.Value.ToString();

            workbook.Close();

            excelApp.Quit();

            return value;
        }

        public static void CreateTestWorkbook(string message = "Hello world (default)!")
        {
            TestClass2.ActuallyCreateTestWorkbook(message);

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void ActuallyCreateTestWorkbook(string message = "Hello world (default)!")
        {
            var excelApp = new XL.Application
            {
                Visible = true
            };

            var workbooks = excelApp.Workbooks;
            var workbook = workbooks.Add();

            var worksheets = workbook.Worksheets;
            var worksheet = worksheets.Add() as XL.Worksheet;

            //var rangeAsSomeType = worksheet.Cells[1, "A"];
            //var range = worksheet.Cells[1, "A"] as Excel.Range;
            var range = worksheet.Range["A1"];
            range.Value = message;

            var range2 = worksheet.Range["A2"];
            range2.Value2 = message + "2";

            //range.Value = "Hello World!";

            workbook.SaveAs(@"C:\Temp\temp.xlsx");

            workbook.Close();

            excelApp.Quit();
        }

        private static void Test()
        {
            //XL.Range range;

            //range.end
        }
    }
}
