using System;


namespace R5T.Excel
{
    public static class RangeSizeExtensions
    {
        public static RangeSize SetFrom(this RangeSize rangeSize, int rows, int columns)
        {
            rangeSize.Rows = rows;
            rangeSize.Columns = columns;

            return rangeSize;
        }

        public static RangeSize SetFrom(this RangeSize rangeSize, object[,] data)
        {
            int rows = data.GetLength(0);
            int columns = data.GetLength(1);

            rangeSize.SetFrom(rows, columns);

            return rangeSize;
        }
    }
}
