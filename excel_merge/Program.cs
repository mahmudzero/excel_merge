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

        public void CreatDoc(String fileName, SpreadsheetDocumentType documentType)
        {
            using (SpreadsheetDocument excelSpreadsheet = SpreadsheetDocument.Create(fileName, documentType))
            {
                WorkbookPart workbookPart = excelSpreadsheet.AddWorkbookPart();
            }
        }

        static void Main(string[] args)
        {
            Debug.WriteLine("Running");
        }
    }
}
