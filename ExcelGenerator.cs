using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ElancoPimsDdsParser
{
    class ExcelGenerator
    {
        private BatchUnit batchUnitRef;
        public ExcelGenerator(BatchUnit batchUnit)
        {
            batchUnitRef = batchUnit;
        }

        public bool generateElancoExcel(string fullFileName)
        {
            string path = Path.GetDirectoryName(fullFileName);
            string filenameMinusExtension = Path.GetFileNameWithoutExtension(fullFileName);
            string outputFile = path + @"\" + filenameMinusExtension + ".xlsx";

            bool createFile = false, success = false;

            if (!testExcelFile(fullFileName, ref createFile))
            {
                Console.WriteLine("Press a key");
                Console.ReadKey();
                return false; // erred. show msg and exit
            }

            if (createFile)
            {
                success = createSpreadsheet(fullFileName);
            }
            else
            {
                success = openSpreadsheet(fullFileName);
            }
            return success;
        }

        private bool createSpreadsheet(string fileName)
        {
            // Create a spreadsheet document by using the file name.
            SpreadsheetDocument spreadsheetDocument =
                 SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart and Workbook objects.
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a WorksheetPart.
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            // Add a Sheets object.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook
                .AppendChild<Sheets>(new Sheets());

            bool success = sendDataToExcel(workbookPart);            

            // Close the document.
            spreadsheetDocument.Close();

            return success;
        }

        private bool openSpreadsheet(string fileName)
        {
            bool success = false;

            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart workbookPart = spreadsheetDoc.WorkbookPart;

                success = sendDataToExcel(workbookPart);

                workbookPart.Workbook.Save();
            } // end of using 

            return success;
        }

        /// <summary>
        /// This method assumes the input parameter is from a
        /// successfully opened Excel spreadsheet.
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private bool sendDataToExcel(WorkbookPart workbookPart)
        {
            // good place to determine which sections are to be processed
            // (e.g. Sim or not)
            //addCellData(workbookPart, sheetName, row, col, cellString);
            bool success = ScriptingOutput.ScriptAreaStructure.generateAreaStructure(batchUnitRef.getOperationList(), workbookPart);

            return success;
        }

        /// <summary>
        /// Returns true if output file exists and there
        /// were no problems opening it.
        /// Returns true if output file does not exist.
        /// Returns false for errors on attempt to open file.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static bool testExcelFile(string fileName, ref bool createFile)
        {
            /*
             * If file exists and it fails this test - then user should
             * be prompted with error.
             */
            bool exists = checkIfFileExists(fileName);
            if (!exists)
            {
                createFile = true;
                return true;
            }
            // file exists so test it for validity
            try
            {
                using (SpreadsheetDocument workBook = SpreadsheetDocument.Open(fileName, true))
                { }
            }
            catch (OpenXmlPackageException e)
            {
                Console.WriteLine("File ({0}) is an invalid Excel format", fileName);
                return false;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

            return true;
        }

        public static bool checkIfFileExists(string fileName)
        {
            return File.Exists(fileName) ? true : false;
        }
        
        public static void addCellData(WorkbookPart workbookPart,
            string sheetName,
            uint row,
            uint col,
            string cellString)
        {
            StringValue sheetId = getSheetRelationshipId(workbookPart, sheetName);
            if (sheetId.Value.Equals("-1"))
            {
                Console.WriteLine("Creating new sheet");
                sheetId = createNewSheet(workbookPart, sheetName);
            }

            addCellToSheet(workbookPart,
                sheetId,
                row,
                col,
                cellString);
        }

        /// <summary>
        /// Returns a StringValue of "-1" if Sheet of given 
        /// name is not found.
        /// Otherwise returns the StringValue relationshipId of the Sheet.
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private static StringValue getSheetRelationshipId(WorkbookPart workbookPart, string sheetName)
        {
            Workbook workbook = workbookPart.Workbook;
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null)
            {
                return new StringValue("-1");
            }
            return sheet.Id;
        }

        private static void addCellToSheet(WorkbookPart workbookPart,
            StringValue sheetId,
            uint row,
            uint col,
            string cellString)
        {
            SheetData sheetData = getSheetData(workbookPart, sheetId);
            Row rowObj = getRow(sheetData, row);

            addCellToRow(rowObj, col, cellString);
        }

        /// <summary>
        /// This method assumes the SheetData with a matching relationship Id, exists.
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="relationshipId"></param>
        /// <returns></returns>
        private static SheetData getSheetData(WorkbookPart workbookPart, StringValue relationshipId)
        {
            var worksheetIE =
                from t in workbookPart.Parts
                where t.RelationshipId.Equals(relationshipId.Value)
                select ((WorksheetPart)t.OpenXmlPart).Worksheet;

            //   sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            var worksheet = worksheetIE.Single<Worksheet>();

            return
                (SheetData)(from t in worksheet.ChildElements
                            where t.GetType().Equals(typeof(SheetData))
                            select t).First();
        }

        private static StringValue createNewSheet(WorkbookPart workbookPart, string sheetName)
        {
            Sheets sheets = workbookPart.Workbook.Descendants<Sheets>().FirstOrDefault();
            // will Sheets ever be null ever on an existing workbook? No.
            uint currentMaxSheetId;
            if (sheets.Count() == 0)
            {
                currentMaxSheetId = 0;
            }
            else
            {
                currentMaxSheetId =
                    (from t in sheets.Elements()
                     select ((uint)((Sheet)t).SheetId)).Max();
            }

            // Add a WorksheetPart to the WorkbookPart. 
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            StringValue relationshipId = workbookPart.GetIdOfPart(worksheetPart);
            // Append a new worksheet and associate it with the workbook. 
            Sheet sheet = new Sheet()
            {
                Id = relationshipId,
                SheetId = (UInt32Value)currentMaxSheetId + 1,
                Name = sheetName
            };
            sheets.Append(sheet);

            return relationshipId;
        }

        private static Row getRow(SheetData sheetData, uint row)
        {
            Row trgtRow = null;

            foreach (var elem in sheetData.ChildElements)
            {
                if (elem.GetType().Equals(typeof(Row)))
                {
                    if (((Row)elem).RowIndex == row)
                    {
                        trgtRow = (Row)elem;
                        break;
                    }
                }
            }
            if (trgtRow == null)
            {
                trgtRow = new Row() { RowIndex = row };
                sheetData.AppendChild(trgtRow);
            }
            return trgtRow;
        }



        private static void addCellToRow(Row row, uint col, string cellString)
        {
            string colName = getColumnName((int)col);
            StringValue newCellReference = colName + row.RowIndex.ToString();

            foreach (Cell cell in row)
            {
                if (cell.CellReference.Value.Equals(newCellReference.Value))
                {
                    cell.Remove();
                }
                //Console.WriteLine("row val: {0}", cell.CellReference);

            }

            // create a new cell
            Cell newCell = createTextCell(newCellReference, cellString);
            insertCellInRow(row, newCell, (int)col);
        }

        public static void insertCellInRow(Row row, Cell cell, int col)
        {
            bool wasInserted = false;
            foreach (Cell existingCell in row)
            {
                int? currentCol = GetColumnIndex(existingCell.CellReference); // will not ever be null
                if (currentCol > col)
                {
                    row.InsertBefore<Cell>(cell, existingCell);
                    wasInserted = true;
                    break;
                }
            }
            if (!wasInserted)
            {
                row.Append(cell);
            }
        }

        /// <summary>
        /// 
        /// from:
        ///   http://stackoverflow.com/questions/28875815/get-the-column-index-of-a-cell-in-excel-using-openxml-c-sharp
        /// </summary>
        /// <param name="cellReference"></param>
        /// <returns></returns>
        private static int? GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return null;
            }

            //remove digits
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            //then multiply that number by our multiplier (which starts at 1)
            //multiply our multiplier by 26 as there are 26 letters
            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);

                mulitplier = mulitplier * 26;
            }

            //the result is zero based so return columnnumber + 1 for a 1 based answer
            //this will match Excel's COLUMN function
            return columnNumber + 1;
        }

        /// <summary> 
        /// This method creates text cell 
        /// </summary> 
        /// <param name="columnIndex"></param> 
        /// <param name="rowIndex"></param> 
        /// <param name="cellValue"></param> 
        /// <returns></returns> 
        private static Cell createTextCell(string cellReference, string cellText)
        {
            Text t = new Text();
            t.Text = cellText;

            InlineString inlineString = new InlineString();
            inlineString.AppendChild(t);

            Cell cell = new Cell();
            cell.CellReference = cellReference;
            cell.DataType = CellValues.InlineString;
            cell.AppendChild(inlineString);

            return cell;
        }

        private static void overwriteTextOfTextCell(Cell cell, string newCellText)
        {
            Text t = new Text();
            t.Text = newCellText;

            InlineString inlineString = new InlineString();
            inlineString.AppendChild(t);
            
            cell.DataType = CellValues.InlineString;
            int count = cell.ChildElements.Count();
            if (count > 0)
            {
                if (count > 1)
                {
                    Console.WriteLine("WARNING! - Overwriting cell with {0:D} children", count);
                }
                //Console.WriteLine("replacing inlinstring child");
                cell.ReplaceChild(inlineString, cell.GetFirstChild<InlineString>());
            }
            else
            {
                cell.AppendChild(inlineString);
            }
        }


        /// <summary> 
        /// Uses a base 10 number and transforms that into
        /// a atring of the Excel column name format
        ///   (e.g. one of: A, B, C,...AA, AB, AC,...)
        ///   from:
        ///     https://code.msdn.microsoft.com/How-to-convert-Word-table-0cb4c9c3
        /// </summary> 
        /// <param name="columnIndex"></param> 
        /// <returns></returns> 
        private static string getColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName =
                    Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }

            return columnName;
        }

    } // end of class ExelGenerator
}

/*
 * ref: http://www.prowareness.com/blog/writing-data-into-excel-document-using-openxml/
 https://code.msdn.microsoft.com/How-to-convert-Word-table-0cb4c9c3
 */
