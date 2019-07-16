using System;


namespace R5T.Excel
{
    public class RangeSize
    {
        public int Rows { get; set; }
        public int Columns { get; set; }


        public RangeSize()
        {
        }

        public RangeSize(int rows, int columns)
        {
            this.Rows = rows;
            this.Columns = columns;
        }
    }
}
