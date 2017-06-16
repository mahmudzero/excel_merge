using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace excel_merge {
    class Program {

        static void print(object value) {
            Console.WriteLine(value);
            Trace.WriteLine(value);
        }

        public static void createDoc(String fileName, SpreadsheetDocumentType documentType) {
            using (SpreadsheetDocument excelSpreadsheet = SpreadsheetDocument.Create(fileName, documentType)) {
                //adding a workbook to the excelSpreadsheet
                WorkbookPart workbook = excelSpreadsheet.AddWorkbookPart();
                workbook.Workbook = new Workbook();
                //add a worksheet to the wokrbook
                WorksheetPart worksheet = workbook.AddNewPart<WorksheetPart>();
                worksheet.Worksheet = new Worksheet(new SheetData());
                //adding sheets to the workbook
                Sheets sheets = workbook.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbook.GetIdOfPart(worksheet), SheetId = 1, Name = "Combined Sheet" };
                sheets.Append(sheet);
                workbook.Workbook.Save();
                print("Finished making document");
            }
        }

        static void Main(string[] args) {
            print("Running");
            createDoc(@"./testDoc.xlsm", SpreadsheetDocumentType.MacroEnabledWorkbook);

        }
    }
}