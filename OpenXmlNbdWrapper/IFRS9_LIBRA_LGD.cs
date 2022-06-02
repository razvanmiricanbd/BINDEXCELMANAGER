using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using NBDCrudWrapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlNbdWrapper
{
    public class IFRS9_LIBRA_LGD
    {
        public void StartFile(string templateFilePath,bool createCopy, string finalFilePath,DataTable data
            , DataTable datapi, DataTable dataweight, DataTable datapd_cdr, DataTable datapd_cdr_y, DataTable data_ldg, DataTable datapd_cdr_y_pi
          )
        {
            string fileName = templateFilePath;
            string destFile = finalFilePath;
            try
            {
                System.IO.File.Copy(fileName, destFile, true);
           
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(destFile, true))
                {
                    var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
                    var sheet = sheets.ElementAt(3);
                    WorksheetPart workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    int i = 1;
                    foreach (DataRow dataRow in data.Rows)
                    {
                        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                        IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i + 1);

                        foreach (Row row in rows)
                        {
                            int j = 0;
                            foreach (Cell cell in row.Descendants<Cell>())
                            {
                                if (cell.CellReference.ToString().Substring(0, 1) == "A" && j == 0)
                                {
                                    cell.CellValue = new CellValue(dataRow["DATA"].ToString());
                                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                if (cell.CellReference.ToString().Substring(0, 1) == "B" && j == 1)
                                {
                                    cell.CellValue = new CellValue(dataRow["GDP_nominal"].ToString());
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                if (cell.CellReference.ToString().Substring(0, 1) == "C" && j == 2)
                                {
                                    cell.CellValue = new CellValue(dataRow["Relative_yearly_change_GDP"].ToString());
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                if (cell.CellReference.ToString().Substring(0, 1) == "D" && j == 3)
                                {
                                    cell.CellValue = new CellValue(dataRow["DR_CORPORATE"].ToString());
                                    // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }
                                j++;
                            }


                        }
                        i++;
                    }
                    //AddPi

                    sheet = sheets.ElementAt(4);
                    workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    i = 1;
                    foreach (DataRow dataRow in datapi.Rows)
                    {
                        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                        IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i + 1);

                        foreach (Row row in rows)
                        {
                            int j = 0;
                            foreach (Cell cell in row.Descendants<Cell>())
                            {
                                if (cell.CellReference.ToString().Substring(0, 1) == "A" && j == 0)
                                {
                                    cell.CellValue = new CellValue(dataRow["DATA"].ToString());
                                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                if (cell.CellReference.ToString().Substring(0, 1) == "B" && j == 1)
                                {
                                    cell.CellValue = new CellValue(dataRow["GDP_nominal"].ToString());
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }


                                if (cell.CellReference.ToString().Substring(0, 1) == "C")
                                {
                                    cell.CellValue = new CellValue(dataRow["Relative_yearly_change_GDP"].ToString());
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                if (cell.CellReference.ToString().Substring(0, 1) == "D")
                                {
                                    cell.CellValue = new CellValue(dataRow["DR_PI"].ToString());
                                    // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }
                                j++;

                            }


                        }
                        i++;
                    }
                    // Add scenario Weight

                    sheet = sheets.ElementAt(0);
                    workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    i = 1;
                    foreach (DataRow dataRow in dataweight.Rows)
                    {
                        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                        IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i + 1);

                        foreach (Row row in rows)
                        {
                            foreach (Cell cell in row.Descendants<Cell>())
                            {



                                if (cell.CellReference.ToString().Substring(0, 2) == "BN")
                                {
                                    cell.CellValue = new CellValue(dataRow["MacroScenario1"].ToString());
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                if (cell.CellReference.ToString().Substring(0, 2) == "BO")
                                {
                                    cell.CellValue = new CellValue(dataRow["MacroScenario2"].ToString());
                                    // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                if (cell.CellReference.ToString().Substring(0, 2) == "BP")
                                {
                                    cell.CellValue = new CellValue(dataRow["MacroScenario3"].ToString());
                                    // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                            }


                        }
                        i++;
                    }


                    // Need to write on rows 22 ( latter it adds +1)
                    i = 21;
                    foreach (DataRow dataRow in dataweight.Rows)
                    {
                        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                        IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i + 1);

                        foreach (Row row in rows)
                        {
                            int j = 0;
                            foreach (Cell cell in row.Descendants<Cell>())
                            {



                                if (cell.CellReference.ToString().Substring(0, 2) == "BJ")
                                {
                                    cell.CellValue = new CellValue(dataRow["Weight1"].ToString());
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                if (cell.CellReference.ToString().Substring(0, 2) == "BK")
                                {
                                    cell.CellValue = new CellValue(dataRow["Weight2"].ToString());
                                    // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                if (cell.CellReference.ToString().Substring(0, 2) == "BL")
                                {
                                    cell.CellValue = new CellValue(dataRow["Weight3"].ToString());
                                    // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }
                                j++;
                            }


                        }
                        i++;
                    }


                    //sheet = sheets.ElementAt(0);
                    //workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    //worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    i = 1;
                    foreach (DataRow dataRow in datapd_cdr.Rows)
                    {
                        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                        IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i + 1);

                        foreach (Row row in rows)
                        {
                            int j = 0;
                            foreach (Cell cell in row.Descendants<Cell>())
                            {

                                if (cell.CellReference.ToString().Substring(0, 1) == "C" && j == 2)
                                {
                                    cell.CellValue = new CellValue(dataRow["cDR"].ToString());
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                j++;

                            }


                        }
                        i++;
                    }


                    sheet = sheets.ElementAt(1);
                    workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    i = 1;
                    foreach (DataRow dataRow in data_ldg.Rows)
                    {
                        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                        IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i + 1);

                        foreach (Row row in rows)
                        {
                            int j = 0;
                            foreach (Cell cell in row.Descendants<Cell>())
                            {



                                if (cell.CellReference.ToString().Substring(0, 1) == "B" && j == 1)
                                {
                                    cell.CellValue = new CellValue(dataRow["TRR"].ToString());
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }

                                // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                j++;

                            }


                        }
                        i++;
                    }

                    sheet = sheets.ElementAt(3);
                    workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    i = 1;
                    foreach (DataRow dataRow in datapd_cdr_y.Rows)
                    {
                        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                        IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i + 1);

                        foreach (Row row in rows)
                        {
                            int j = 0;
                            foreach (Cell cell in row.Descendants<Cell>())
                            {



                                if (cell.CellReference.ToString().Substring(0, 2) == "AA")
                                {
                                    cell.CellValue = new CellValue(dataRow["TimeofData"].ToString());
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }
                                if (cell.CellReference.ToString().Substring(0, 2) == "AD")
                                {
                                    cell.CellValue = new CellValue(dataRow["cDRY"].ToString());
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }
                                // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                j++;

                            }


                        }
                        i++;
                    }

                    if (datapd_cdr_y_pi != null)
                    {
                        sheet = sheets.ElementAt(4);
                        workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                        worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                        i = 1;
                        foreach (DataRow dataRow in datapd_cdr_y_pi.Rows)
                        {
                            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i + 1);

                            foreach (Row row in rows)
                            {
                                int j = 0;
                                foreach (Cell cell in row.Descendants<Cell>())
                                {



                                    if (cell.CellReference.ToString().Substring(0, 2) == "AA")
                                    {
                                        cell.CellValue = new CellValue(dataRow["TimeofData"].ToString());
                                        //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    }
                                    if (cell.CellReference.ToString().Substring(0, 2) == "AD")
                                    {
                                        cell.CellValue = new CellValue(dataRow["cDRY"].ToString());
                                        //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    }
                                    // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    j++;

                                }


                            }
                            i++;
                        }

                    }
                    doc.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                    doc.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                    worksheet.Save();
                }
            }
            catch (Exception e)
            { throw e; }
        }


        public void AddPI( string finalFilePath, DataTable data
          )
        {
   
            string destFile = finalFilePath;

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(destFile, true))
            {
                var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
                var sheet = sheets.ElementAt(5);
                WorksheetPart workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                int i = 1;
                foreach (DataRow dataRow in data.Rows)
                {
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i + 1);

                    foreach (Row row in rows)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                           
                           

                            if (cell.CellReference.ToString().Substring(0, 1) == "C")
                            {
                                cell.CellValue = new CellValue(dataRow["Relative_yearly_change_GDP"].ToString());
                                //cell.DataType = new EnumValue<CellValues>(CellValues.String);
                            }

                            if (cell.CellReference.ToString().Substring(0, 1) == "D")
                            {
                                cell.CellValue = new CellValue(dataRow["DR_PI"].ToString());
                                // cell.DataType = new EnumValue<CellValues>(CellValues.String);
                            }

                        }


                    }
                    i++;
                }



                doc.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                doc.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                worksheet.Save();
            }
        }
        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
    }
}
