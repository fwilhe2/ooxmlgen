using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlTesting
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "ComplexTestDocument.xlsx";
            CreateComplexExcelFile(filePath);
            Console.WriteLine($"Created complex testing file at: {filePath}");
        }

        public static void CreateComplexExcelFile(string filepath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {
                // 1. Setup Workbook
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

                // 2. Add Stylesheet (Required for dates and formatting)
                WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = CreateStylesheet();
                stylesPart.Stylesheet.Save();

                // 3. Add Shared String Table (How Excel actually stores text)
                SharedStringTablePart shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                shareStringPart.SharedStringTable = new SharedStringTable();

                // 4. Create Sheet 1: Data Types
                WorksheetPart worksheetPart1 = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart1.Worksheet = new Worksheet(new SheetData());
                Sheet sheet1 = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart1), SheetId = 1, Name = "DataTypes" };
                sheets.Append(sheet1);

                // 5. Create Sheet 2: Styles
                WorksheetPart worksheetPart2 = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart2.Worksheet = new Worksheet(new SheetData());
                Sheet sheet2 = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart2), SheetId = 2, Name = "StyledSheet" };
                sheets.Append(sheet2);

                // --- POPULATE SHEET 1 (Different Data Types) ---
                SheetData sheetData1 = worksheetPart1.Worksheet.GetFirstChild<SheetData>();
                Row row1 = new Row() { RowIndex = 1 };

                // A1: Shared String
                int stringIndex = InsertSharedStringItem("Hello Shared String", shareStringPart);
                row1.Append(CreateCell("A1", stringIndex.ToString(), CellValues.SharedString, 0));

                // B1: Number (Notice no specific CellValues type is needed for standard numbers)
                row1.Append(CreateCell("B1", "42.5", CellValues.Number, 0));

                // C1: Boolean (0 = False, 1 = True)
                row1.Append(CreateCell("C1", "1", CellValues.Boolean, 0));

                // D1: Date (Stored as a number, formatted as a date via StyleIndex 1)
                double excelDate = DateTime.Now.ToOADate();
                row1.Append(CreateCell("D1", excelDate.ToString(), CellValues.Number, 1));

                sheetData1.Append(row1);

                // --- POPULATE SHEET 2 (Styles) ---
                SheetData sheetData2 = worksheetPart2.Worksheet.GetFirstChild<SheetData>();
                Row row2 = new Row() { RowIndex = 1 };

                // A1: Bold Text with Yellow Fill (StyleIndex 2)
                int styleStringIndex = InsertSharedStringItem("Important Data!", shareStringPart);
                row2.Append(CreateCell("A1", styleStringIndex.ToString(), CellValues.SharedString, 2));

                sheetData2.Append(row2);

                // Save all changes
                workbookPart.Workbook.Save();
            }
        }

        // --- HELPER METHODS ---

        // Creates a basic cell
        private static Cell CreateCell(string reference, string value, CellValues dataType, uint styleIndex)
        {
            Cell cell = new Cell() { CellReference = reference, DataType = dataType, StyleIndex = styleIndex };
            cell.CellValue = new CellValue(value);
            return cell;
        }

        // Inserts a string into the SharedStringTable and returns its index
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            int i = 0;
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text) return i;
                i++;
            }
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            shareStringPart.SharedStringTable.Save();
            return i;
        }

        // Generates a minimal valid Stylesheet with custom styles
        private static Stylesheet CreateStylesheet()
        {
            return new Stylesheet(
                new Fonts(
                    new Font(), // Index 0 - Default
                    new Font(new Bold()) // Index 1 - Bold
                ),
                new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - Default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - Required by Excel
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }) { PatternType = PatternValues.Solid }) // Index 2 - Yellow Fill
                ),
                new Borders(
                    new Border() // Index 0 - Default
                ),
                new CellFormats(
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 }, // Index 0: Default Style
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0, NumberFormatId = 14, ApplyNumberFormat = true }, // Index 1: Date Format (14 is built-in m/d/yyyy)
                    new CellFormat() { FontId = 1, FillId = 2, BorderId = 0, ApplyFont = true, ApplyFill = true } // Index 2: Bold + Yellow Fill
                )
            );
        }
    }
}