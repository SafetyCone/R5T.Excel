using System;


namespace R5T.Excel.Construction
{
    public static class Construction
    {
        public static void SubMain()
        {
            Construction.CreateTestWorkbook();
        }

        private static void CreateExampleWorkbook()
        {
            // Create a new workbook.

            // Get the first sheet.

            // Test getting/setting object[,] values.

            // Test if range has value (is empty).

            // Change number formats.

            // Set named range.

            // Set named range values.

            // Get named range.

            // Get named range values.

            // Test formulas.

            // Test formatting: alignment, bold.
        }

        private static void CreateTestWorkbook()
        {
            var message = "Howdy there!";

            TestClass2.CreateTestWorkbook(message);
        }
    }
}
