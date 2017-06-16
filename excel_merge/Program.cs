using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace excel_merge
{
    class Program
    {

        static void print(object value)
        {
            Console.WriteLine(value);
            Trace.WriteLine(value);
        }

        public void CreatDoc(String fileName, SpreadsheetDocumentType documentType)
        {
            using (SpreadsheetDocument excelSpreadsheet = SpreadsheetDocument.Create(fileName, documentType))
            {
                //adding a workbook to the excelSpreadsheet
                WorkbookPart workbook = excelSpreadsheet.AddWorkbookPart();
                workbook.Workbook = new Workbook();
                //add a worksheet to the wokrbook
                WorksheetPart worksheet = workbook.AddNewPart<WorksheetPart>();
                worksheet.Worksheet = new Worksheet(new SheetData());
            }
        }

        static int add(int x, int y)
        {
            return x + y;
        }


        static void Main(string[] args)
        {
            print("Running");
            print(add(5, 3));
        }
    }
}