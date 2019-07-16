using System;


namespace R5T.Excel
{
    public static class ApplicationExtensions
    {
        public static void Calculate(this Application application)
        {
            application.XlApplication.Calculate();
        }
    }
}
