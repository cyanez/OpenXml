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

    public static void UpdateCell(string fileName, string sheetName, string columnName, uint rowIndex, string cellValue) {

      OpenSrepadSheetDocument(fileName, sheetName, columnName, rowIndex, cellValue, new EnumValue<CellValues>(CellValues.String));
    }
    public static void OpenSrepadSheetDocument(string fileName, string sheetName, string columnName, uint rowIndex, string cellValue, EnumValue<CellValues> cellDataType) {
      using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true)) {
        WorkbookPart workbookPart = document.WorkbookPart;
        WorksheetPart worksheetPart = GetSheet(workbookPart, sheetName);

        if (worksheetPart == null) {
          throw new Exception("I cant find sheet: " + sheetName);
        }

        Row row = GetRow(worksheetPart.Worksheet, rowIndex);
        Cell cell = InsertCell(worksheetPart.Worksheet, row, columnName, rowIndex);

        cell.CellValue = new CellValue(cellValue);
                   
        cell.DataType = cellDataType;
                    
        worksheetPart.Worksheet.Save();











        // Cell cell = GetCell(worksheetPart, addressName);

        //  MessageBox.Show(GetCellValue(workbookPart,cell));

      }



    }

    private static WorksheetPart GetSheet(WorkbookPart workbookPart, string sheetName) {      
      Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            
      if (sheet == null) {
        throw new ArgumentException("I cant find sheet: " + sheetName);
      }

      return (WorksheetPart) (workbookPart.GetPartById(sheet.Id));
    }

    private static bool ExistRow(Worksheet worksheet, uint rowIndex) {
      int count = worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex).Count();

      if (count == 0) {
        return false;
      }

      return true;      
    }

    private static Row GetRow(Worksheet worksheet, uint rowIndex) {
      if (!ExistRow(worksheet, rowIndex)) {
        Row row = new Row() {
          RowIndex = rowIndex
        };
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        sheetData.Append(row);
      }

      return worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
    }

    private static Cell InsertCell(Worksheet worksheet, Row row, string columnName, uint rowIndex) {
      
      string cellReference = columnName + rowIndex;

      if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0) {
        return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
      } else {       
        Cell refCell = null;
        foreach (Cell cell in row.Elements<Cell>()) {
          if (string.Compare(cell.CellReference.Value, cellReference, true) > 0) {
            refCell = cell;
            break;
          }
        }

        Cell newCell = new Cell() { CellReference = cellReference };
        row.InsertBefore(newCell, refCell);

        worksheet.Save();
        return newCell;
      }

    }



  }
}
