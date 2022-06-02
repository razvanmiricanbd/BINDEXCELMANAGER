using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NBDCrudWrapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
namespace OpenXmlNbdWrapper
{
    public class OpemXmlExcelReader
    {
        private string dbconnection;

        public OpemXmlExcelReader(string dbconnection)
        {
            this.dbconnection = dbconnection;

        }

        public int ReadeExcel(string filePath, out string errmessage)
        {
            errmessage = null;
            int importId = -1;
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, true))
                {
                    var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
                    var sheet = sheets.First();
                    WorksheetPart workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;

                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> rows = sheetData.Elements<Row>();
                    MSSqlEngine engine = new MSSqlEngine(dbconnection);
                    MSSqlParameter[] parameters = null;

                    DataTable tb = engine.RunProcedureQuery("BINDAPP$GetExcelFileImportSeq", parameters);
                    importId = Int32.Parse(tb.Rows[0]["id"].ToString());
                    Dictionary<string, object> excelRow = null;
                    int rowNo = 0;
                    int firstcolumnNo = 0, columnNo = 0;
                    foreach (Row row in rows)
                    {
                        columnNo = 0;
                        excelRow = new Dictionary<string, object>();
                        foreach (Cell cell in row.Descendants<Cell>())
                        {

                            string cellValue = null;

                            if ((cell.DataType != null) && (cell.DataType.Value == CellValues.SharedString))
                            {
                                int ssid = int.Parse(cell.CellValue.Text);
                                cellValue = sst.ChildElements[ssid].InnerText;

                            }
                            else
                            if (((cell.DataType != null) && (cell.DataType.Value == CellValues.Date))
                                || ((cell.DataType == null) && (cell.StyleIndex == 4))
                                    )
                            {
                                cellValue = DateTime.FromOADate(Int32.Parse(cell.CellValue.Text)).ToString("dd-MM-yyyy");

                            }
                            else if (cell.CellValue != null)
                            {
                                cellValue = cell.CellValue.Text;
                            }
                            if (cellValue != null)
                            {
                                columnNo++;
                                excelRow.Add("Column" + columnNo, cellValue);
                            }
                            // Column Number based on the first row.


                        }
                        if (rowNo == 0)
                            firstcolumnNo = columnNo;
                        if (columnNo > 0)
                        {
                            //Insert row
                            parameters = new MSSqlParameter[columnNo + 4];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@import_id",
                                Value = importId
                            };

                            parameters[1] = new MSSqlParameter
                            {
                                Name = "@row_no",
                                Value = rowNo
                            };
                            parameters[2] = new MSSqlParameter
                            {
                                Name = "@file_name",
                                Value = filePath
                            };

                            parameters[3] = new MSSqlParameter
                            {
                                Name = "@isheader",
                                Value = (rowNo == 0) ? "Y" : "N"
                            };
                            int x = 4;
                            foreach (string key in excelRow.Keys)
                            {
                                parameters[x] = new MSSqlParameter
                                {
                                    Name = "@" + key,
                                    Value = excelRow[key].ToString()
                                };
                                x++;
                            }
                            engine.RunProcedureStatment("CRUD$add_excel_file_temp_buffer", parameters);
                            rowNo++;
                        }

                    }

                }
            }
            catch (Exception e)
            { errmessage = e.Message; }


            return importId;

        }


        public int ReadeExcelImport(int fileId, string filePath, out string errmessage)
        {
            errmessage = null;
            int importId = -1;
            MSSqlEngine engine = new MSSqlEngine(dbconnection);
            MSSqlParameter[] parameters =  new MSSqlParameter[1];
            parameters[0] = new MSSqlParameter
            {
                Name = "@import_id",
                Value = fileId
            };
            DataTable tb = engine.RunProcedureQuery("BINDAPP$GET_XLS_IMPORT_DET", parameters);
            Dictionary<int, string> columntype = new Dictionary<int, string>(); 
            foreach (DataRow row in tb.Rows)
            {
                columntype.Add(int.Parse(row["OrderNo"].ToString())-1, row["Column_Type"].ToString());
            }
            Dictionary<string, object> excelRow = null;
            int rowNo = 0;
            int firstcolumnNo = 0, columnNo = 0;
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, true))
                {
                    var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
                    var sheet = sheets.First();
                    WorksheetPart workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;

                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> rows = sheetData.Elements<Row>();
                    
                     parameters = null;

                     tb = engine.RunProcedureQuery("BINDAPP$GetExcelFileImportSeq", parameters);
                    importId = Int32.Parse(tb.Rows[0]["id"].ToString());
                   
                    foreach (Row row in rows)
                    {
                        columnNo = 0;
                        excelRow = new Dictionary<string, object>();
                        foreach (Cell cell in row.Descendants<Cell>())
                        {

                            string cellValue = null;

                            if ((cell.DataType != null) && (cell.DataType.Value == CellValues.SharedString))
                            {
                                int ssid = int.Parse(cell.CellValue.Text);
                                cellValue = sst.ChildElements[ssid].InnerText;

                            }
                            else
                            if (((cell.DataType != null) && (cell.DataType.Value == CellValues.Date))
                                || ((cell.DataType == null) && (columntype[columnNo] =="DI"))
                                    )
                            {
                                cellValue = DateTime.FromOADate(Int32.Parse(cell.CellValue.Text)).ToString("dd-MM-yyyy");

                            }
                            else if (cell.CellValue != null)
                            {
                                cellValue = cell.CellValue.Text;
                            }
                            if (cellValue != null)
                            {
                                columnNo++;
                                excelRow.Add("Column" + columnNo, cellValue);
                            }
                            // Column Number based on the first row.


                        }
                        if (rowNo == 0)
                            firstcolumnNo = columnNo;
                        if (columnNo > 0)
                        {
                            //Insert row
                            parameters = new MSSqlParameter[columnNo + 4];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@import_id",
                                Value = importId
                            };

                            parameters[1] = new MSSqlParameter
                            {
                                Name = "@row_no",
                                Value = rowNo
                            };
                            parameters[2] = new MSSqlParameter
                            {
                                Name = "@file_name",
                                Value = filePath
                            };

                            parameters[3] = new MSSqlParameter
                            {
                                Name = "@isheader",
                                Value = (rowNo == 0) ? "Y" : "N"
                            };
                            int x = 4;
                            foreach (string key in excelRow.Keys)
                            {
                                parameters[x] = new MSSqlParameter
                                {
                                    Name = "@" + key,
                                    Value = excelRow[key].ToString()
                                };
                                x++;
                            }
                            engine.RunProcedureStatment("CRUD$add_excel_file_temp_buffer", parameters);
                            rowNo++;
                        }

                    }

                }
            }
            catch (Exception e)
            { errmessage = e.Message; }


            return importId;

        }

        public int ReadeExcelSection(int fileId, int sectionNo,string filePath, out string errmessage)
        {
            errmessage = null;
            int importId = -1;
            MSSqlEngine engine = new MSSqlEngine(dbconnection);
            MSSqlParameter[] parameters = new MSSqlParameter[2];
            parameters[0] = new MSSqlParameter
            {
                Name = "@import_id",
                Value = fileId
            };

            parameters[1] = new MSSqlParameter
            {
                Name = "@Section_No",
                Value = sectionNo
            };
            int start_Row = 0, end_Row = 0, start_Column = 0, end_Column = 0, tab_No = 0;
            //string filePath = "";
            DataTable tb = engine.RunProcedureQuery("BINDAPP$GET_XLS_IMPORT", parameters);
            foreach (DataRow row in tb.Rows)
            {
                start_Row = int.Parse(row["Start_Row"].ToString());
                end_Row = int.Parse(row["End_Row"].ToString());
                start_Column = int.Parse(row["Start_Column"].ToString());
                end_Column = int.Parse(row["End_Column"].ToString());
                tab_No  = int.Parse(row["Tab_No"].ToString())-1;
               // filePath = row["File_Import_Path"].ToString();

            }



            //Parameters
            parameters = new MSSqlParameter[1];
            parameters[0] = new MSSqlParameter
            {
                Name = "@import_id",
                Value = fileId
            };
           

            tb = engine.RunProcedureQuery("BINDAPP$GET_XLS_IMPORT_DET", parameters);
            Dictionary<int, string> columntype = new Dictionary<int, string>();
            foreach (DataRow row in tb.Rows)
            {
                columntype.Add(int.Parse(row["OrderNo"].ToString()) - 1, row["Column_Type"].ToString());
            }
            Dictionary<string, object> excelRow = null;
            int rowNo = 0;
            int firstcolumnNo = 0, columnNo = 0;
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, true))
                {
                    var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
                    var sheet = sheets.ElementAt(tab_No); 
                    WorksheetPart workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;

                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex >=start_Row && r.RowIndex <= end_Row); ;

                    parameters = null;

                    tb = engine.RunProcedureQuery("BINDAPP$GetExcelFileImportSeq", parameters);
                    importId = Int32.Parse(tb.Rows[0]["id"].ToString());

                    foreach (Row row in rows)
                    {
                        columnNo = 0;
                        int columnIndex = 0;
                        excelRow = new Dictionary<string, object>();
                        foreach (Cell cell in row.Descendants<Cell>())


                        {
                            // Filter Columns
                            columnIndex++;
                            if (columnIndex >= start_Column && columnIndex <= end_Column)
                            {

                                string cellValue = null;

                                if ((cell.DataType != null) && (cell.DataType.Value == CellValues.SharedString))
                                {
                                    int ssid = int.Parse(cell.CellValue.Text);
                                    cellValue = sst.ChildElements[ssid].InnerText;

                                }
                                else
                                if (((cell.DataType != null) && (cell.DataType.Value == CellValues.Date))
                                    || ((cell.DataType == null) && (columntype[columnNo] == "DI"))
                                        )
                                {
                                    cellValue = DateTime.FromOADate(Int32.Parse(cell.CellValue.Text)).ToString("dd-MM-yyyy");

                                }
                                else if (cell.CellValue != null)
                                {
                                    cellValue = cell.CellValue.Text;
                                }
                                if (cellValue != null)
                                {
                                    columnNo++;
                                    excelRow.Add("Column" + columnNo, cellValue);
                                }
                                // Column Number based on the first row.

                            }
                        }
                        
                        if (rowNo == 0)
                            firstcolumnNo = columnNo;
                        if (columnNo > 0)
                        {
                            //Insert row
                            parameters = new MSSqlParameter[columnNo + 4];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@import_id",
                                Value = importId
                            };

                            parameters[1] = new MSSqlParameter
                            {
                                Name = "@row_no",
                                Value = rowNo
                            };
                            parameters[2] = new MSSqlParameter
                            {
                                Name = "@file_name",
                                Value = filePath
                            };

                            parameters[3] = new MSSqlParameter
                            {
                                Name = "@isheader",
                                Value = (rowNo == 0) ? "Y" : "N"
                            };
                            int x = 4;
                            foreach (string key in excelRow.Keys)
                            {
                                parameters[x] = new MSSqlParameter
                                {
                                    Name = "@" + key,
                                    Value = excelRow[key].ToString()
                                };
                                x++;
                            }
                            engine.RunProcedureStatment("CRUD$add_excel_file_temp_buffer", parameters);
                            rowNo++;
                        }

                    }

                }
            }
            catch (Exception e)
            { errmessage = e.Message; }


            return importId;

        }
    }
}

