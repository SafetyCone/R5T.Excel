using System;

using Xl = Microsoft.Office.Interop.Excel;

using R5T.NetStandard;


namespace R5T.Excel
{
    public static class ExcelCalculationModeExtensions
    {
        public static Xl.XlCalculation ToXlCalculation(this ExcelCalculationMode mode)
        {
            switch(mode)
            {
                case ExcelCalculationMode.Automatic:
                    return Xl.XlCalculation.xlCalculationAutomatic;

                case ExcelCalculationMode.Manual:
                    return Xl.XlCalculation.xlCalculationManual;

                case ExcelCalculationMode.SemiAutomatic:
                    return Xl.XlCalculation.xlCalculationSemiautomatic;

                default:
                    throw new ArgumentException(EnumHelper.UnexpectedEnumerationValueMessage(mode), nameof(mode));
            }
        }
    }
}
