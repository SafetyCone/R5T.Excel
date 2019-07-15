using System;


namespace R5T.Excel.Construction
{
    public static class Construction
    {
        public static void SubMain()
        {
            Construction.CreateTestWorkbook();
        }

        private static void CreateTestWorkbook()
        {
            var message = "Howdy there!";

            TestClass2.CreateTestWorkbook(message);
        }
    }
}
