using System;
using System.Collections.Generic;
using System.IO;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.Excel
{
    /// <summary>
    /// Represents an Excel workbook.
    /// </summary>
    /// <remarks>
    /// Not disposable since "disposing" a workbook would mean losing work unless the workbook was saved.
    /// Thus workbooks are saved then closed.
    /// </remarks>
    public class Workbook
    {
        internal Xl.Workbook XlWorkbook { get; private set; }

        public Application Application { get; private set; }

        public string Name
        {
            get
            {
                var output = this.XlWorkbook.Name;
                return output;
            }
            // Read-only.
        }
        public string FilePath
        {
            get
            {
                var output = this.XlWorkbook.FullName;
                return output;
            }
            // Read-only.
        }
        public ExcelFileFormat FileFormat
        {
            get
            {
                var xlFileFormat = this.XlWorkbook.FileFormat;
                
                var output = xlFileFormat.ToExcelFileFormat();
                return output;
            }
            // Read-only.
        }
        public int WorksheetCount
        {
            get
            {
                var output = this.XlWorkbook.Worksheets.Count;
                return output;
            }
        }
        public IEnumerable<Worksheet> Worksheets
        {
            get
            {
                foreach (Xl.Worksheet xlWorksheet in this.XlWorkbook.Worksheets)
                {
                    var worksheet = new Worksheet(xlWorksheet, this);
                    yield return worksheet;
                }
            }
        }



        internal Workbook(Xl.Workbook xlWorkbook, Application application)
        {
            this.XlWorkbook = xlWorkbook;
            this.Application = application;
        }

        /// <summary>
        /// Closes the Excel workbook without saving changes.
        /// </summary>
        public void Close()
        {
            this.XlWorkbook.Close(false);
        }

        public void SaveAs(string filePath, ExcelFileFormat fileFormat, bool overwrite = true)
        {
            // Workaround for Workbook.SaveAs() not having an easy overwrite option.
            if(overwrite && File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            var xlFileFormat = fileFormat.ToXlFileFormat();

            this.XlWorkbook.SaveAs(filePath, xlFileFormat);
        }

        public void SaveAs(string filePath, bool overwrite = true)
        {
            this.SaveAs(filePath, ExcelFileFormat.XLSX, overwrite);
        }

        public void Select()
        {
            this.XlWorkbook.Activate();
        }

        public Worksheet NewWorksheet()
        {
            var xlWorksheet = this.XlWorkbook.Worksheets.AddWorksheet();

            var worksheet = new Worksheet(xlWorksheet, this);
            return worksheet;
        }

        public bool HasWorksheet(string name)
        {
            var output = false;
            foreach (Xl.Worksheet worksheet in this.XlWorkbook.Worksheets)
            {
                if(name == worksheet.Name)
                {
                    output = true;
                    break;
                }
            }

            return output;
        }

        public Worksheet GetWorksheet(string name)
        {
            var xlWorksheet = this.XlWorkbook.Worksheets[name] as Xl.Worksheet;

            var worksheet = new Worksheet(xlWorksheet, this);
            return worksheet;
        }

        /// <summary>
        /// Gets an existing worksheet.
        /// </summary>
        /// <param name="index">The zero-based (0-based) worksheet index.</param>
        public Worksheet GetWorksheet(int index)
        {
            var xlWorksheet = this.XlWorkbook.Worksheets[index] as Xl.Worksheet;

            var worksheet = new Worksheet(xlWorksheet, this);
            return worksheet;
        }

        public void AddNamedRange(Range range, string name)
        {
            this.XlWorkbook.Names.Add(name, range.XlRange);
        }
    }
}
