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
    public class IFRS9_PATR_Card_Indiv_Prov_Writer
    {
        public IFRS9_PATR_Card_Indiv_Prov_Writer()
        {

        }

        public void WriteCard(string templateFilePath, string finalFilePath,
            DateTime refDate, DataTable dataCustomer, DataTable dataExposure
            , DataTable dataColateral
            , DataTable dataAlocation)
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

                DataRow customerRow = dataCustomer.Rows[0];
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex <= 6);

                foreach (Row row in rows)
                {
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        if (cell.CellReference == "A1")
                        {
                            cell.CellValue = new CellValue(customerRow["Customer_Name"].ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }

                        if (cell.CellReference == "B2")
                        {
                            cell.CellValue = new CellValue(refDate.ToOADate().ToString());
                        }

                        

                        if (cell.CellReference == "B3")
                        {
                            DateTime dataAnaliza = (DateTime)customerRow["Analysis_Start_Date"];
                            cell.CellValue = new CellValue(dataAnaliza.ToOADate().ToString());
                        }
                        if (cell.CellReference == "B4")
                        {
                            cell.CellValue = new CellValue(customerRow["Reason_For_Analysis"].ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                       
                        if (cell.CellReference == "A6")
                        {
                            cell.CellValue = new CellValue(customerRow["Customer_id"].ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        if (cell.CellReference == "B6")
                        {
                            cell.CellValue = new CellValue(customerRow["Total_Exposure_LCY"].ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        if (cell.CellReference == "C6")
                        {
                            cell.CellValue = new CellValue(customerRow["Insolvency"].ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        if (cell.CellReference == "D6")
                        {
                            cell.CellValue = new CellValue(customerRow["Is_Foreclosure"].ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }

                    }

                    
                }

                worksheet.Save();
                // Exposure
                sheet = sheets.ElementAt(1);
                workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                sheetData = worksheet.GetFirstChild<SheetData>();


                foreach (DataRow excelrow in dataExposure.Rows)
                {
                    Row exrow = new Row();
                    foreach (DataColumn column in dataExposure.Columns)
                    {
                        /*
                        if (column.ColumnName != "Ref_date"
                            && column.ColumnName != "TotalRows"
                            && column.ColumnName != "row_number"
                            && column.ColumnName != "TotalPages")
                        {
                            if(column.ColumnName != "Loan_ID")
                            exrow.Append(
                                ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.Number)
                                );
                            else exrow.Append(
                                ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.String)
                                );

    */
                        if (column.ColumnName != "Ref_date"
                                                && column.ColumnName != "TotalRows"
                                                && column.ColumnName != "row_number"
                                                && column.ColumnName != "TotalPages"
                                                && column.ColumnName != "row_key")
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
                //Collateral
                sheet = sheets.ElementAt(2);
                workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                sheetData = worksheet.GetFirstChild<SheetData>();


                foreach (DataRow excelrow in dataColateral.Rows)
                {
                    Row exrow = new Row();
                    foreach (DataColumn column in dataColateral.Columns)
                    {
                        if (column.ColumnName != "Ref_date"
                            && column.ColumnName != "TotalRows"
                            && column.ColumnName != "row_number"
                            && column.ColumnName != "TotalPages"
                            && column.ColumnName != "row_key")
                        {
                            /*
                            if (column.ColumnName != "Loan_ID" && column.ColumnName != "Collateral_ID"
                                 && column.ColumnName != "Collateral_Code"
                                 &&   column.ColumnName != "Collateral_Desc"
                                    )
                                exrow.Append(
                                    ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.Number)
                                    );
                            else exrow.Append(
                                ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.String)
                                );
                                */
                            if (column.ColumnName != "Ref_date"
                                                && column.ColumnName != "TotalRows"
                                                && column.ColumnName != "row_number"
                                                && column.ColumnName != "TotalPages"
                                                && column.ColumnName != "row_key")
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

                    }
                    sheetData.AppendChild(exrow);
                }

                worksheet.Save();
                // Allocation
                sheet = sheets.ElementAt(3);
                workpart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                sheetData = worksheet.GetFirstChild<SheetData>();


                foreach (DataRow excelrow in dataAlocation.Rows)
                {
                    Row exrow = new Row();
                    foreach (DataColumn column in dataAlocation.Columns)
                    {
                        if (column.ColumnName != "Ref_date"
                            && column.ColumnName != "TotalRows"
                            && column.ColumnName != "row_number"
                            && column.ColumnName != "TotalPages"
                            && column.ColumnName != "row_key")
                        {

                            /*if (column.ColumnName != "Loan_ID" && column.ColumnName != "Collateral_ID")
                                exrow.Append(
                                    ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.Number)
                                    );
                            else exrow.Append(
                                ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.String)
                                );
                                */
                            if (column.ColumnName != "Ref_date"
                                                && column.ColumnName != "TotalRows"
                                                && column.ColumnName != "row_number"
                                                && column.ColumnName != "TotalPages")
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

                    }
                    sheetData.AppendChild(exrow);
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
