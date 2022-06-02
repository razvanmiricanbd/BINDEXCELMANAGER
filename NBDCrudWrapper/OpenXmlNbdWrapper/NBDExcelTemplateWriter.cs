using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Data;
using NBDCrudWrapper;

namespace OpenXmlNbdWrapper
{
    public class NBDExcelTemplateWriter
    {
        private string dbconnection;
        private Dictionary<string, uint> _sharedStrings;
        public NBDExcelTemplateWriter(string dbconnection) {
            this.dbconnection = dbconnection;
           
        }
        public void WriteExcelOutput(string templateFilePath, string finalFilePath,DateTime refDate, int? executionid)
        {
            string fileName = templateFilePath;
            string destFile = finalFilePath;
            System.IO.File.Copy(fileName, destFile, true);
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(destFile, true))
            {
                _sharedStrings = new Dictionary<string,uint>();
                Sheets sheets = doc.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                SharedStringTablePart shareStringPart;

                
                if (doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = doc.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                uint i = 0;

                // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
                foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
                {
                    _sharedStrings.Add(item.InnerText, i);
                    i++;
                }
                //Read the first Sheet from Excel file.
                foreach (Sheet sheet in sheets)
                {

                    //Get the Worksheet instance.
                    WorksheetPart workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                    
                    //Fetch all the rows present in the Worksheet.
                    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

                    Dictionary<string, string> sqlBlocks = new Dictionary<string, string>();

                    //Loop through the Worksheet rows.
                    foreach (Row row in rows)
                    {


                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            string val = GetValue(doc, cell);
                            /*
                            if (cell.CellFormula != null && cell.CellFormula.InnerText.Contains("VALUE"))
                            {
                                CellFormula cellformula = new CellFormula();
                                cellformula.Text = "VALUE(1)";
                                //CellValue cellValue = new CellValue();
                                //cellValue.Text = "1";
                                cell.CellFormula = cellformula;
                                //cell.CellValue =cellValue;
                            }
                            */
                            if (val != null && val.StartsWith("<#") && val.EndsWith("#>"))
                            {


                                sqlBlocks.Add(cell.CellReference, val.Replace("<#", "").Replace("#>", ""));

                            }
                        }

                    }

                    uint rowsAdded = 0;
                    foreach (string key in sqlBlocks.Keys)
                    {
                        string sqlCode = sqlBlocks[key];
                        moveRowsInExcel(key, sqlCode, doc, sheetData, shareStringPart, refDate, executionid, ref rowsAdded);

                       
                    }
                    rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                    foreach (Row row in rows)
                    {
                        int x = 0;
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            cell.CellReference = (char)(65 + x) + row.RowIndex.ToString();
                            x += 1;
                        }

                    }

                    doc.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                    doc.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                    worksheet.Save();
                    

                }
            }

        }



        public void WriteExcelOutputSAX(string templateFilePath, string finalFilePath, DateTime refDate, int? executionid)
        {
            string fileName = templateFilePath;
            string destFile = finalFilePath;
            /*
            //System.IO.File.Copy(fileName, destFile, true);
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
            {
                _sharedStrings = new Dictionary<string, uint>();
                Sheets sheets = doc.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                SharedStringTablePart shareStringPart;


                if (doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = doc.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                uint i = 0;

                // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
                foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
                {
                    _sharedStrings.Add(item.InnerText, i);
                    i++;
                }
                //Read the first Sheet from Excel file.
                foreach (Sheet sheet in sheets)
                {

                    //Get the Worksheet instance.
                    WorksheetPart workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                    //Fetch all the rows present in the Worksheet.
                    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

                    Dictionary<string, string> sqlBlocks = new Dictionary<string, string>();

                    //Loop through the Worksheet rows.
                    foreach (Row row in rows)
                    {


                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            string val = GetValue(doc, cell);
                           
                            if (val != null && val.StartsWith("<#") && val.EndsWith("#>"))
                            {


                                sqlBlocks.Add(cell.CellReference, val.Replace("<#", "").Replace("#>", ""));

                            }
                        }

                    }

                    uint rowsAdded = 0;
                    foreach (string key in sqlBlocks.Keys)
                    {
                        string sqlCode = sqlBlocks[key];
                        moveRowsInExcel(key, sqlCode, doc, sheetData, shareStringPart, refDate, executionid, ref rowsAdded);
                    }
                    rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                    foreach (Row row in rows)
                    {
                        int x = 0;
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            cell.CellReference = (char)(65 + x) + row.RowIndex.ToString();
                            x += 1;
                        }

                    }

                    //doc.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                    //doc.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                    //worksheet.Save();


                }
                */
            MSSqlEngine engine = new MSSqlEngine(dbconnection);
            MSSqlParameter[] parameters = null;

         
                parameters = new MSSqlParameter[1];
                parameters[0] = new MSSqlParameter
                {
                    Name = "@ref_date",
                    Value = (refDate).ToString("yyyy-MM-dd")
                };
            
            DataTable excelData = engine.RunProcedureQuery("BINDEXCEL$GET_PRV_OUT_ECL", parameters);
            using (SpreadsheetDocument xl = SpreadsheetDocument.Create(destFile, SpreadsheetDocumentType.Workbook))
                {
                    List<OpenXmlAttribute> oxa;
                    OpenXmlWriter oxw;

                    xl.AddWorkbookPart();
                    WorksheetPart wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>();

                    oxw = OpenXmlWriter.Create(wsp);
                    oxw.WriteStartElement(new Worksheet());
                    oxw.WriteStartElement(new SheetData());
                    int i = 0;
                foreach (DataRow excelrow in excelData.Rows)
                {
                    i++;
                    oxa = new List<OpenXmlAttribute>();
                    // this is the row index
                    oxa.Add(new OpenXmlAttribute("r", null, i.ToString()));

                    oxw.WriteStartElement(new Row(), oxa);

                    foreach (DataColumn column in excelData.Columns)
                    {
                        oxa = new List<OpenXmlAttribute>();
                        // this is the data type ("t"), with CellValues.String ("str")
                        oxa.Add(new OpenXmlAttribute("t", null, "str"));

                        // it's suggested you also have the cell reference, but
                        // you'll have to calculate the correct cell reference yourself.
                        // Here's an example:
                        //oxa.Add(new OpenXmlAttribute("r", null, "A1"));

                        oxw.WriteStartElement(new Cell(), oxa);

                        oxw.WriteElement(new CellValue(excelrow[column.ColumnName].ToString()));

                        // this is for Cell
                        oxw.WriteEndElement();
                    }

                    // this is for Row
                    oxw.WriteEndElement();
                }

                    // this is for SheetData
                    oxw.WriteEndElement();
                    // this is for Worksheet
                    oxw.WriteEndElement();
                    oxw.Close();
                    
                    oxw = OpenXmlWriter.Create(xl.WorkbookPart);
                    oxw.WriteStartElement(new Workbook());
                    oxw.WriteStartElement(new Sheets());

                    // you can use object initialisers like this only when the properties
                    // are actual properties. SDK classes sometimes have property-like properties
                    // but are actually classes. For example, the Cell class has the CellValue
                    // "property" but is actually a child class internally.
                    // If the properties correspond to actual XML attributes, then you're fine.
                    oxw.WriteElement(new Sheet()
                    {
                        Name = "Sheet1",
                        SheetId = 1,
                        Id = xl.WorkbookPart.GetIdOfPart(wsp)
                    });

                    // this is for Sheets
                    oxw.WriteEndElement();
                    // this is for Workbook
                    oxw.WriteEndElement();
                    oxw.Close();
                    
                    xl.Close();
                }

            
        }

        private void populateCodeInExcel(string columnRef, string sqlBlock, SpreadsheetDocument doc,SheetData sheetData, SharedStringTablePart shareStringPart, DateTime refDate, int? executionid, ref uint rowsAdded )
        {

            string columnReference = Regex.Replace(columnRef.ToUpper(), @"[\d]", string.Empty);
            uint rowIndex = (uint)Int32.Parse(Regex.Replace(columnRef.ToUpper(), @"[A-Z]", string.Empty));
            rowIndex += rowsAdded;


            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            MSSqlEngine engine = new MSSqlEngine(dbconnection);
            MSSqlParameter[] parameters = null;

            if (executionid.HasValue)
            {
                parameters = new MSSqlParameter[2];
                parameters[0] = new MSSqlParameter
                {
                    Name = "@ref_date",
                    Value = (refDate).ToString("yyyy-MM-dd")
                };
                parameters[1] = new MSSqlParameter
                {
                    Name = "@executionid",
                    Value = executionid
                };
            }
            else
            {
                parameters = new MSSqlParameter[1];
                parameters[0] = new MSSqlParameter
                {
                    Name = "@ref_date",
                    Value = (refDate).ToString("yyyy-MM-dd")
                };
            }
            DataTable excelData = engine.RunProcedureQuery(sqlBlock, parameters);
            UpdateRowIndex(sheetData, rowIndex,(uint) excelData.Rows.Count);

            Cell cell = null;
            // Cell cell = InsertCellInWorksheet(columnReference, rowIndex, workpart);
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnRef).Count() > 0)
            {
                cell = row.Elements<Cell>().Where(c => c.CellReference.Value == columnRef).First();
            }
            if (cell != null)
                row.RemoveChild<Cell>(cell);
            Row originalRow = (Row)row.Clone();
            Row newRow = row;
            Row lastRow = row;
            bool first = true;

            foreach (DataRow excelrow in excelData.Rows)
            {
                if (!first)
                {
                    rowIndex += 1;
                    newRow = (Row)originalRow.Clone();
                    newRow.RowIndex = rowIndex;
                    sheetData.InsertAfter(newRow, lastRow);
                    lastRow = newRow;
                }
                else first = false;
                foreach (DataColumn column in excelData.Columns)
                {
                    int val = InsertSharedStringItem(excelrow[column.ColumnName].ToString(), shareStringPart);
                    newRow.Append(
                        // ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.String)
                        ConstructCell(val.ToString(), CellValues.SharedString)
                        );
                }
                
            }
            shareStringPart.SharedStringTable.Save();
            /*
            int val = InsertSharedStringItem("X1", shareStringPart);
            // row.Append(ConstructCell("A"+sqlCode, CellValues.String,"A"+ rowIndex.ToString()));
            row.Append(ConstructCell(val.ToString(), CellValues.SharedString));
            val = InsertSharedStringItem("Y1", shareStringPart);
            row.Append(ConstructCell(val.ToString(), CellValues.SharedString));
            val = InsertSharedStringItem("Z1", shareStringPart);
            row.Append(ConstructCell(val.ToString(), CellValues.SharedString));
            //
            //cell.CellValue = new CellValue("X");
            
            rowIndex += 1;
            Row lastRow = row;
            // Row newRow = new Row() { RowIndex = rowIndex };
            Row newRow = (Row)originalRow.Clone();
            newRow.RowIndex = rowIndex;
            sheetData.InsertAfter(newRow, lastRow);
            lastRow = newRow;
            */
        }



        private void moveRowsInExcel(string columnRef, string sqlBlock, SpreadsheetDocument doc, SheetData sheetData, SharedStringTablePart shareStringPart, DateTime refDate, int? executionid, ref uint rowsAdded)
        {

            string columnReference = Regex.Replace(columnRef.ToUpper(), @"[\d]", string.Empty);
            uint rowIndex = (uint)Int32.Parse(Regex.Replace(columnRef.ToUpper(), @"[A-Z]", string.Empty));
            rowIndex += rowsAdded;


            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            MSSqlEngine engine = new MSSqlEngine(dbconnection);
            MSSqlParameter[] parameters = null;

            if (executionid.HasValue)
            {
                parameters = new MSSqlParameter[2];
                parameters[0] = new MSSqlParameter
                {
                    Name = "@ref_date",
                    Value = (refDate).ToString("yyyy-MM-dd")
                };
                parameters[1] = new MSSqlParameter
                {
                    Name = "@executionid",
                    Value = executionid
                };
            }
            else
            {
                parameters = new MSSqlParameter[1];
                parameters[0] = new MSSqlParameter
                {
                    Name = "@ref_date",
                    Value = (refDate).ToString("yyyy-MM-dd")
                };
            }
            DataTable excelData = engine.RunProcedureQuery(sqlBlock, parameters);
            UpdateRowIndex(sheetData, rowIndex, (uint)excelData.Rows.Count);
            rowsAdded += (uint)excelData.Rows.Count;



        }

        /*
         * Using DOM
         * 
         * 
          private void populateCodeInExcel(string columnRef, string sqlBlock, SheetData sheetData, SharedStringTablePart shareStringPart, DateTime refDate, int? executionid, ref uint rowsAdded )
        {

            string columnReference = Regex.Replace(columnRef.ToUpper(), @"[\d]", string.Empty);
            uint rowIndex = (uint)Int32.Parse(Regex.Replace(columnRef.ToUpper(), @"[A-Z]", string.Empty));
            rowIndex += rowsAdded;


            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            MSSqlEngine engine = new MSSqlEngine(dbconnection);
            MSSqlParameter[] parameters = null;

            if (executionid.HasValue)
            {
                parameters = new MSSqlParameter[2];
                parameters[0] = new MSSqlParameter
                {
                    Name = "@ref_date",
                    Value = (refDate).ToString("yyyy-MM-dd")
                };
                parameters[1] = new MSSqlParameter
                {
                    Name = "@executionid",
                    Value = executionid
                };
            }
            else
            {
                parameters = new MSSqlParameter[1];
                parameters[0] = new MSSqlParameter
                {
                    Name = "@ref_date",
                    Value = (refDate).ToString("yyyy-MM-dd")
                };
            }
            DataTable excelData = engine.RunProcedureQuery(sqlBlock, parameters);
            UpdateRowIndex(sheetData, rowIndex,(uint) excelData.Rows.Count);

            Cell cell = null;
            // Cell cell = InsertCellInWorksheet(columnReference, rowIndex, workpart);
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnRef).Count() > 0)
            {
                cell = row.Elements<Cell>().Where(c => c.CellReference.Value == columnRef).First();
            }
            if (cell != null)
                row.RemoveChild<Cell>(cell);
            Row originalRow = (Row)row.Clone();
            Row newRow = row;
            Row lastRow = row;
            bool first = true;

            foreach (DataRow excelrow in excelData.Rows)
            {
                if (!first)
                {
                    rowIndex += 1;
                    newRow = (Row)originalRow.Clone();
                    newRow.RowIndex = rowIndex;
                    sheetData.InsertAfter(newRow, lastRow);
                    lastRow = newRow;
                }
                else first = false;
                foreach (DataColumn column in excelData.Columns)
                {
                    int val = InsertSharedStringItem(excelrow[column.ColumnName].ToString(), shareStringPart);
                    newRow.Append(
                        // ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.String)
                        ConstructCell(val.ToString(), CellValues.SharedString)
                        );
                }
                
            }
            shareStringPart.SharedStringTable.Save();
           
        }

         **/
        private int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.

            if (_sharedStrings.ContainsKey(text))
                return (int) _sharedStrings[text];
            int i = _sharedStrings.Count-1;
            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            //shareStringPart.SharedStringTable.Save();
            _sharedStrings.Add(text, (uint)(i + 1));
            return i+1;
        }
        private  string GetValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = null;
            if (cell.CellFormula == null && cell.CellValue != null)
            {
                value = cell.CellValue.InnerText;
                if (value != "#VALUE")
                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        value = doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
                    }

            }
            if (value == null && cell.CellFormula != null)
                value = cell.CellFormula.InnerText;
            return value;
        }
        private  Cell ConstructCell(string value, CellValues dataType, string cellref)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                CellReference = cellref
            };
        }
        private  Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }

        private  Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
        private   void UpdateRowIndex(SheetData sheetData, uint start, uint rowNo)
        {

            foreach (Row row in sheetData.Elements<Row>().Where(r => r.RowIndex > start))
            {
                row.RowIndex += rowNo;
            }

        }
    }
}
