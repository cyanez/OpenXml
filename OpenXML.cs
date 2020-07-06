using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLManager {
  class OpenXML {
    public static void CreateSpreadsheetWorkbook(string filepath, string sheetName) {

      using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)) {
        WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

        Sheet sheet = new Sheet() {
          Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
          SheetId = 1,
          Name = sheetName
        };
        sheets.Append(sheet);

        workbookpart.Workbook.Save();
      }

    }

    public static void InsertWorksheet(string filepath, string sheetName) {
      
      using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filepath, true)) {
        
        WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());

        Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
        string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

        
        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Count() > 0) {
          sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        }
       
        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);
      }

    }


  }
}
