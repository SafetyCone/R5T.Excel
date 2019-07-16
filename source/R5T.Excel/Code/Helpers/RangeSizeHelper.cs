using System;


namespace R5T.Excel
{
    public static class RangeSizeHelper
    {
        public static RangeSize From(object[,] data)
        {
            var size = new RangeSize()
                .SetFrom(data);

            return size;
        }
    }
}
