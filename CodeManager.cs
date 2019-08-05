using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using deliverynote2xml.tags;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace deliverynote2xml
{
    /// <summary>
    /// CodeManager Class
    /// 
    /// This is a service class. Implements common code utilities.
    /// </summary>
    public static class CodeManager
    {
        //Using Microsoft.Office.Interop, it takes 2.24 minutes to parse a sheet of
        //246 rows by 44 columns
        public static List<string> parseExcelDocument(string path)
        {
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(@path);

            Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[1];

            Excel.Range xlRange = xlWorkSheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<string> result = new List<string>();
            string line;

            for (int i = 1; i <= rowCount; i++)
            {
                line = "";

                for (int j = 1; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j] != null)
                    {
                        line += xlRange.Cells[i, j].value2 == null ? "NULL" : xlRange.Cells[i, j].value2;
                        line += j == colCount ? "" : ", ";
                    }
                    
                }

                result.Add(line);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            //Release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorkSheet);

            //Close and release
            xlWorkBook.Close();
            Marshal.ReleaseComObject(xlWorkBook);

            //Quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return result;
        }

        public static bool generateDeliveryNoteXml(DeliveryNotes deliveryNotes, string pathToXmlFile)
        {
            XmlSerializer xs = new XmlSerializer(typeof(DeliveryNotes));

            MemoryStream memory = new MemoryStream();

            xs.Serialize(memory, deliveryNotes);
            FileStream fs = null;

            try
            {
                fs = new FileStream(@pathToXmlFile, FileMode.Create);

                memory.WriteTo(fs);
                fs.Flush();

                return true;
            }
            catch (Exception)
            {

                return false;
            }
            finally
            {
                //fs.Close();
                memory.Close();
            }

        }

        public static List<string> parseWordDocument(string path)
        {
            Application word = new Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();

            object fileName = path;

            // Define an object to pass to the API for missing parameters
            object missing = System.Type.Missing;
            doc = word.Documents.Open(ref fileName,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

            List<string> data = new List<string>();


            Frames frames = doc.Frames;

            for (int i = 1; i < frames.Count; i++)
            {
                data.Add(string.Format("{0} -  {1}", i, frames[i].Range.Text ?? "NULL"));
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            //Release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(frames);

            //Close and release
            ((_Document)doc).Close();
            Marshal.ReleaseComObject(doc);

            //Quit and release
            ((_Application)word).Quit();
            Marshal.ReleaseComObject(word);

            return data;
        }

        //Initialized on form load.
        public static string CustomerDataFilePath { get; set; }

        public static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
        }


        public static bool rtfDocumentIsAlreadyOpen(string rtfPath)
        {
            Application wordApp = null;
            Microsoft.Office.Interop.Word.Document wordDoc = null;

            Process[] processes = Process.GetProcessesByName("WinWord");

            try
            {
                if (processes.Length > 0)
                {
                    wordApp = (Application)Marshal.GetActiveObject("Word.Application");

                    if (wordApp.Documents.Count > 0)
                    {
                        for (int i = 1; i <= wordApp.Documents.Count; i++)
                        {
                            wordDoc = wordApp.Documents[i];

                            if (wordDoc.FullName.ToLower() == rtfPath.ToLower())
                            {
                                return true;
                            }
                        }
                    }
                }

                return false;
            }
            catch (Exception)
            {  
                throw;
            }

        }

        public static bool excelWorkbookIsAlreadyOpen(string excelWorkbookPath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook excelWorkbook = null;

            Process[] processes = Process.GetProcessesByName("excel");

            try
            {
                if (processes.Length > 0)
                {
                    excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    
                    if (excelApp.Workbooks.Count > 0)
                    {
                        for (int i = 1; i <= excelApp.Workbooks.Count; i++)
                        {
                            excelWorkbook = excelApp.Workbooks[i];

                            if (excelWorkbook.FullName.ToLower() == excelWorkbookPath.ToLower())
                            {
                                return true;
                            }
                        }
                    }
                }

                return false;
            }
            catch (Exception)
            {
                throw;
            }

        }

    }
}