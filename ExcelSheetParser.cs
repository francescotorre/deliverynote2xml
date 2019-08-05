using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace deliverynote2xml
{
    class ExcelSheetParser
    {
        //Path to Excel file...
        private string Path;

        public ExcelSheetParser(string path)
        {
            this.Path = @path;
        }

        public string GetDiscounts(string customerCode)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range xlRange = null;
            Excel.Range colRange = null;
            Excel.Range resultRange = null;
            string discount = "";

            try
            {
                xlApp = new Excel.Application();

                xlWorkBook = xlApp.Workbooks.Open(this.Path);

                xlWorkSheet = xlWorkBook.Sheets[1];

                xlRange = xlWorkSheet.UsedRange;

                // Get the range object where you want to search from
                colRange = xlWorkSheet.Columns[1, Type.Missing];

                // Search search String in the range, if find result, return a range
                resultRange = colRange.Find(What: customerCode,
                    LookIn: Excel.XlFindLookIn.xlValues,
                    LookAt: Excel.XlLookAt.xlWhole,
                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Excel.XlSearchDirection.xlNext);

                if (resultRange != null)
                {
                    int rowIndex = resultRange.Row;

                    discount = xlRange.Cells[rowIndex, 10].value2 ?? "";
                }
                else
                {
                    discount = "CUSTOMER_NOT_FOUND";
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                //Release com objects to fully kill excel process from running in the background
                CodeManager.ReleaseComObject(xlRange);
                xlRange = null;

                CodeManager.ReleaseComObject(xlWorkSheet);
                xlWorkSheet = null;

                CodeManager.ReleaseComObject(colRange);
                colRange = null;

                CodeManager.ReleaseComObject(resultRange);
                resultRange = null;

                //Close and release
                xlWorkBook.Close();
                CodeManager.ReleaseComObject(xlWorkBook);

                //Quit and release
                if (xlApp != null)
                {
                    xlApp.Quit();
                }

                CodeManager.ReleaseComObject(xlApp);
                xlApp = null;

                GC.Collect();
            }

            return discount;
        }

    }
}
