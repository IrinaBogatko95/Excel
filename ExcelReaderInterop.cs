using System.Collections.Generic;
using System.Runtime.InteropServices;
using ImageConnect.Test.Func.Shared.Core.Utils;
using Microsoft.Office.Interop.Excel;

namespace ImageConnect.Test.Func.Shared.Core.Helpers.ExcelReaderInterop
{
    /// <summary>
    /// This class contains the Excel Interop code.
    /// </summary>
    public class ExcelReaderInterop
    {
        /// <summary>
        /// The Application.
        /// </summary>
        private Application _excelApp;

        /// <summary>
        /// The Workbook.
        /// </summary>
        private Workbook _workbook;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelReaderInterop"/> class.
        /// </summary>
        public ExcelReaderInterop()
        {
            _excelApp = new Application();
        }

        /// <summary>
        /// Opens an excel file.
        /// </summary>
        /// <param name="path">File path.</param>
        public void OpenFile(string path)
        {
            WaitHelper.WaitFor(() => FileHelper.IsFileReady(path), 20);

            _workbook = RetryUtils.RetryIfThrown<COMException, Workbook>(() => _excelApp.Workbooks.Open(path),
                ConfigHelper.DefaultRetriesNumber, () => _workbook.Close());
        }

        /// <summary>
        /// Sets a column values for selected rows.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="values">Key - number of row, value - value to set.</param>
        /// <param name="columnNumber">Number of column to set.</param>
        public void SetColumnValues(Worksheet worksheet, Dictionary<int, string> values, int columnNumber)
        {
            foreach (var value in values)
            {
                RetryUtils.RetryIfThrown<COMException, string>(() => worksheet.Cells[value.Key, columnNumber].Value = value.Value,
                    ConfigHelper.DefaultRetriesNumber, () => _workbook.Close());
            }
        }

        /// <summary>
        /// Saves a workbook and close it.
        /// </summary>
        public void SaveWorkbookAndClose()
        {
            RetryUtils.RetryIfThrown<COMException>(() => _workbook.Save(), ConfigHelper.DefaultRetriesNumber);
            RetryUtils.RetryIfThrown<COMException>(() => _workbook.Close(), ConfigHelper.DefaultRetriesNumber);
        }

        /// <summary>
        /// Gets first worksheet.
        /// </summary>
        /// <returns>The worksheet.</returns>
        public Worksheet GetFirstWorksheet()
        {
            var worksheet = RetryUtils.RetryIfThrown<COMException, Worksheet>(() => (Worksheet)_workbook.Worksheets.Item[1],
                ConfigHelper.DefaultRetriesNumber, () => _workbook.Close());

            return worksheet;
        }
    }
}
