﻿using System;
using System.Collections.Generic;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.Excel
{
    public class Range
    {
        internal Xl.Range XlRange { get; private set; }

        public Worksheet Worksheet { get; private set; }

        public Workbook Workbook
        {
            get
            {
                var output = this.Worksheet.Workbook;
                return output;
            }
        }
        public Application Application
        {
            get
            {
                var output = this.Workbook.Application;
                return output;
            }
        }
        public int Row
        {
            get
            {
                var row = this.XlRange.Row;
                return row;
            }
        }
        public int RowCount
        {
            get
            {
                var rowCount = this.XlRange.Rows.Count;
                return rowCount;
            }
        }
        public int Column
        {
            get
            {
                var column = this.XlRange.Column;
                return column;
            }
        }
        public int ColumnCount
        {
            get
            {
                var columnCount = this.XlRange.Columns.Count;
                return columnCount;
            }
        }
        public IEnumerable<Range> Cells
        {
            get
            {
                foreach (Xl.Range xlCell in this.XlRange.Cells)
                {
                    var cell = new Range(xlCell, this.Worksheet);
                    yield return cell;
                }
            }
        }
        public object Value
        {
            get
            {
                var value = this.XlRange.Value2;
                return value;
            }
            set
            {
                this.XlRange.Value2 = value;
            }
        }
        public Decimal ValueDecimal
        {
            get
            {
                var value = Convert.ToDecimal(this.XlRange.Value2);
                return value;
            }
            set
            {
                this.XlRange.Value2 = value;
            }
        }
        public double ValueDouble
        {
            get
            {
                var value = Convert.ToDouble(this.XlRange.Value2);
                return value;
            }
            set
            {
                this.XlRange.Value2 = value;
            }
        }
        public int ValueInt
        {
            get
            {
                var value = Convert.ToInt32(this.XlRange.Value2);
                return value;
            }
            set
            {
                this.XlRange.Value2 = value;
            }
        }
        public string ValueString
        {
            get
            {
                var value = Convert.ToString(this.XlRange.Value2);
                return value;
            }
            set
            {
                this.XlRange.Value2 = value;
            }
        }
        public object[,] Values
        {
            get
            {
                var output = this.XlRange.Value as object[,];
                return output;
            }
            set
            {
                this.XlRange.Value = value;
            }
        }
        public bool IsEmpty
        {
            get
            {
                var output = this.XlRange.Value is null;
                return output;
            }
        }
        public bool IsNumeric
        {
            get
            {
                var output = this.XlRange.Application.WorksheetFunction.IsNumber(this.XlRange.Value);
                return output;
            }
        }
        public double ColumnWidth
        {
            get
            {
                var columnWidth = Convert.ToDouble(this.XlRange.ColumnWidth);
                return columnWidth;
            }
            set
            {
                this.XlRange.ColumnWidth = value;
            }
        }
        public string Formula
        {
            get
            {
                var formula = Convert.ToString(this.XlRange.Formula);
                return formula;
            }
            set
            {
                this.XlRange.Formula = value;
            }
        }
        public string NumberFormat
        {
            get
            {
                var numberFormat = Convert.ToString(this.XlRange.NumberFormat);
                return numberFormat;
            }
            set
            {
                this.XlRange.NumberFormat = value;
            }
        }
        public Range EndUp
        {
            get
            {
                var xlRange = this.XlRange.End[Xl.XlDirection.xlUp];

                var range = new Range(xlRange, this.Worksheet);
                return range;
            }
        }
        public Range EndDown
        {
            get
            {
                var xlRange = this.XlRange.End[Xl.XlDirection.xlDown];

                var range = new Range(xlRange, this.Worksheet);
                return range;
            }
        }
        public Range EndLeft
        {
            get
            {
                var xlRange = this.XlRange.End[Xl.XlDirection.xlToLeft];

                var range = new Range(xlRange, this.Worksheet);
                return range;
            }
        }
        public Range EndRight
        {
            get
            {
                var xlRange = this.XlRange.End[Xl.XlDirection.xlToRight];

                var range = new Range(xlRange, this.Worksheet);
                return range;
            }
        }
        public Range this[int row, int column]
        {
            get
            {
                var xlRange = this.XlRange[row + 1, column + 1] as Xl.Range; // Excel is 1-based.

                var range = new Range(xlRange, this.Worksheet);
                return range;
            }
        }


        internal Range(Xl.Range xlRange, Worksheet worksheet)
        {
            this.XlRange = xlRange;
            this.Worksheet = worksheet;
        }
    }
}
