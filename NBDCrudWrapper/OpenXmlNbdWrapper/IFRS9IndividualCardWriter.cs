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
    public class IFRS9IndividualCardWriter
    {
        public IFRS9IndividualCardWriter() {

        }

        public void WriteCard(string templateFilePath, string finalFilePath,
            DateTime refDate, string recId, string customerId, decimal Eir, decimal K, decimal EAD,
            string ccy,
            decimal scenario1Percent, decimal scenario2Percent, decimal scenario3Percent, decimal scenario4Percent,
            string hasChanges,
            DataTable data, DataTable cashFlow)
        {
            string fileName = templateFilePath;
            string destFile = finalFilePath;
            System.IO.File.Copy(fileName, destFile, true);
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(destFile, true))
            {
                var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
                var sheet = sheets.ElementAt(0);


                WorksheetPart workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex <9);
                foreach (Row row in rows)
                { foreach (Cell cell in row.Descendants<Cell>())
                    {

                        if (cell.CellReference == "B1")
                        {
                            cell.CellValue = new CellValue(refDate.ToOADate().ToString());
                        }
                          
                        if (cell.CellReference == "B2")
                        {
                            cell.CellValue = new CellValue(recId);
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        
                        if (cell.CellReference == "B3")
                        {
                            cell.CellValue = new CellValue(customerId);
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        if (cell.CellReference == "B4")
                        {
                            cell.CellValue = new CellValue(Eir.ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        if (cell.CellReference == "B5")
                        {
                            cell.CellValue = new CellValue(K.ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        if (cell.CellReference == "B6")
                        {
                            cell.CellValue = new CellValue(EAD.ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                     

                    }
                
                }
                if (data != null) {

                    rows = sheetData.Elements<Row>().Where(r => (r.RowIndex >= 9 && r.RowIndex <= 13));
                    int i = 8;
                    int j = -1;
                    foreach (Row row in rows)
                    {
                        i++;
                        j++;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            /*
                            if (cell.CellReference == "A" + i.ToString())
                            {
                                cell.CellValue = new CellValue("A" + i.ToString());
                            }
                            */
                            if (cell.CellReference == "B" + i.ToString())
                            {
                                cell.CellValue = new CellValue(data.Rows[j]["ActualValue"].ToString());
                            }

                            if (cell.CellReference == "C" + i.ToString())
                            {
                                cell.CellValue = new CellValue(data.Rows[j]["Scenario 1"].ToString());
                            }
                            if (cell.CellReference == "D" + i.ToString())
                            {
                                cell.CellValue = new CellValue(data.Rows[j]["Scenario 2"].ToString());
                            }
                            if (cell.CellReference == "E" + i.ToString())
                            {
                                cell.CellValue = new CellValue(data.Rows[j]["Scenario 3"].ToString());
                            }
                            if (cell.CellReference == "F" + i.ToString())
                            {
                                cell.CellValue = new CellValue(data.Rows[j]["Scenario1Period"].ToString());
                            }

                            if (cell.CellReference == "G" + i.ToString())
                            {
                                cell.CellValue = new CellValue(data.Rows[j]["Scenario2Period"].ToString());
                            }
                            if (cell.CellReference == "H" + i.ToString())
                            {
                                cell.CellValue = new CellValue(data.Rows[j]["Scenario3Period"].ToString());
                            }
                        }
                    }
                }// 
                if (hasChanges == "N")
                {
                    rows = sheetData.Elements<Row>().Where(r => (r.RowIndex == 23));
                    foreach (Row row in rows)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            if (cell.CellReference == "C" + row.RowIndex.ToString())
                            {
                                cell.CellValue = new CellValue(scenario1Percent.ToString());
                            }
                            if (cell.CellReference == "D" + row.RowIndex.ToString())
                            {
                                cell.CellValue = new CellValue(scenario2Percent.ToString());
                            }
                            if (cell.CellReference == "E" + row.RowIndex.ToString())
                            {
                                cell.CellValue = new CellValue(scenario3Percent.ToString());
                            }
                        }
                    }

                    rows = sheetData.Elements<Row>().Where(r => (r.RowIndex == 26));
                    foreach (Row row in rows)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            if (cell.CellReference == "D" + row.RowIndex.ToString())
                            {
                                cell.CellValue = new CellValue(ccy);
                                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                            }
                        }
                    }

                    worksheet.Save();

                }
                else
                {

                    rows = sheetData.Elements<Row>().Where(r => (r.RowIndex == 18));
                    foreach (Row row in rows)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            if (cell.CellReference == "C" + row.RowIndex.ToString())
                            {
                                cell.CellValue = new CellValue(scenario1Percent.ToString());
                            }
                            if (cell.CellReference == "D" + row.RowIndex.ToString())
                            {
                                cell.CellValue = new CellValue(scenario2Percent.ToString());
                            }
                            if (cell.CellReference == "E" + row.RowIndex.ToString())
                            {
                                cell.CellValue = new CellValue(scenario3Percent.ToString());
                            }
                            if (cell.CellReference == "F" + row.RowIndex.ToString())
                            {
                                cell.CellValue = new CellValue(scenario4Percent.ToString());
                            }
                        }
                    }

                    rows = sheetData.Elements<Row>().Where(r => (r.RowIndex == 20));
                    foreach (Row row in rows)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            if (cell.CellReference == "D" + row.RowIndex.ToString())
                            {
                                cell.CellValue = new CellValue(ccy);
                                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                            }
                        }
                    }
                    worksheet.Save();
                    //CashFlow
                    sheet = sheets.ElementAt(1);
                    workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    sheetData = worksheet.GetFirstChild<SheetData>();

                    foreach (DataRow excelrow in cashFlow.Rows)
                    {
                        Row exrow = new Row();
                        foreach (DataColumn column in cashFlow.Columns)
                        {
                            if (column.ColumnName == "Payment_Date"
                                || column.ColumnName == "Payment_Description"
                                || column.ColumnName == "Payment_Ccy"
                                || column.ColumnName == "Payment_value_ccy"
                                )
                            {

                                    var s = column.DataType;
                                    if (s.Name == "Int32" || s.Name == "Int16" || s.Name == "Decimal" || s.Name == "Double" || s.Name == "Float")
                                        exrow.Append(
                                            ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.Number)
                                            );
                                    else
                                        exrow.Append(
                                            ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.String)
                                            );
                                
                            }

                        }
                        sheetData.AppendChild(exrow);
                    }

                    worksheet.Save();
                }
                        doc.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                doc.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                
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
