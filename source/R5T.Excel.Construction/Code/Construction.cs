using System;
using System.IO;
using System.Reflection;


namespace R5T.Excel.Construction
{
    public static class Construction
    {
        public static void SubMain()
        {
            //Construction.CreateTestWorkbook();
            Construction.CreateExampleWorkbook();
        }

        private static void CreateExampleWorkbook()
        {
            using (var app = new Application())
            {
                // Create a new workbook.
                var wkbk = app.NewWorkbook();

                // Get the first sheet.
                var ws = wkbk.GetWorksheet(Constants.Sheet1Name);

                // Rename the sheet.
                ws.Name = "Test Worksheet";

                // Get the A1 range.
                var a1 = ws.GetA1Range();

                // Set object array values.
                var data = new object[4, 2];
                data[0, 0] = "Machine:"; data[0, 1] = Environment.MachineName;
                data[1, 0] = "User:"; data[1, 1] = Environment.UserDomainName + Path.DirectorySeparatorChar + Environment.UserName;
                data[2, 0] = "DateTime:"; data[2, 1] = DateTime.UtcNow;
                data[3, 0] = "Process Name:"; data[3, 1] = Assembly.GetEntryAssembly().FullName;

                var upperLeft = a1;
                var size = new RangeSize().SetFrom(data);
                var dataRange = upperLeft.GetRange(size);
                dataRange.Values = data;

                // Set column widths.
                var columnWidths = new double[]
                {
                    20,
                    30,
                };
                ws.SetColumnWidths(columnWidths);

                // Make row-headers bold.
                var rowHeadersColumn = dataRange.GetColumn(0);
                rowHeadersColumn.Bold();

                // Align values horizontally-left.
                var valuesColumn = dataRange.GetColumn(1);
                valuesColumn.AlignHorizontalLeft();

                // Get object array values.

                // Test if range has value (is empty).

                // Change number formats.

                // Set named range.
                valuesColumn.SetName("Author_Values");

                // Get named range.

                // Test formulas.

                // Save the workbooks.
                var filePath = @"C:\Temp\temp.xlsx";
                wkbk.SaveAs(filePath);
            }
        }

        private static void CreateTestWorkbook()
        {
            var message = "Howdy there!";

            TestClass2.CreateTestWorkbook(message);
        }
    }
}
