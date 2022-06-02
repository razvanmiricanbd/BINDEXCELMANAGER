using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Data;
using System;
namespace OpenXmlNbdWrapper
{
    public class NbdOpenXmlExcelWriter
    {


        private Dictionary<string, uint> _shareStringDictionary;
        private uint _shareStringMaxIndex ;
        public NbdOpenXmlExcelWriter()
        {
            _shareStringDictionary = new Dictionary<string, uint>();
            _shareStringMaxIndex = 0;
        }
        public static void WriteExcelOutputForTable(string FilePath, DataTable excelData, string excelSheetName = null)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                string sheetName = "Sheet1";

                if (excelSheetName != null)
                    sheetName = excelSheetName;
                else if (excelData.TableName != null)
                    sheetName = excelData.TableName;

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();



                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Constructing header
                Row row = new Row();
                foreach (DataColumn column in excelData.Columns)
                    row.Append(
                        ConstructCell(column.ColumnName, CellValues.String)
                       );

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Inserting each employee
                foreach (DataRow excelrow in excelData.Rows)
                {
                    row = new Row();
                    foreach (DataColumn column in excelData.Columns)
                    {
                        row.Append(
                            ConstructCell(excelrow[column.ColumnName].ToString(), CellValues.String)
                            );
                    }
                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
            }
        }
        private static Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }

        public void WriteExcelOutputForTableSAX(string FilePath, DataTable excelData, string excelSheetName = null)
        {
            using (SpreadsheetDocument xl = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook))
            {
                List<OpenXmlAttribute> oxa;
                OpenXmlWriter oxw;

                WorkbookPart workbookPart = xl.AddWorkbookPart();
                WorksheetPart wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>();

                oxw = OpenXmlWriter.Create(wsp);
                oxw.WriteStartElement(new Worksheet());
                oxw.WriteStartElement(new SheetData());
                int i = 0;


                string sheetName = "Sheet1";

                if (excelSheetName != null)
                    sheetName = excelSheetName;
                else if (excelData.TableName != null)
                    sheetName = excelData.TableName;
                // Write Header
                i++;
                oxa = new List<OpenXmlAttribute>();
                // this is the row index
                oxa.Add(new OpenXmlAttribute("r", null, i.ToString()));

                oxw.WriteStartElement(new Row(), oxa);

                foreach (DataColumn column in excelData.Columns)
                {
                    if (column.ColumnName != "row_number" &&
                        column.ColumnName != "TotalRows" &&
                        column.ColumnName != "TotalPages" &&
                        column.ColumnName != "row_key")
                    {
                        // oxa = new List<OpenXmlAttribute>();
                        // this is the data type ("t"), with CellValues.String ("str")
                        //oxa.Add(new OpenXmlAttribute("t", null, "str"));

                        // it's suggested you also have the cell reference, but
                        // you'll have to calculate the correct cell reference yourself.
                        // Here's an example:
                        //oxa.Add(new OpenXmlAttribute("r", null, "A1"));

                        // oxw.WriteStartElement(new Cell(), oxa);

                        // oxw.WriteElement(new CellValue(column.ColumnName));
                        List<OpenXmlAttribute> attributes = new List<OpenXmlAttribute>();
                        attributes.Add(new OpenXmlAttribute("s", null, "2"));
                        WriteCellValueSax(oxw, column.ColumnName, CellValues.SharedString, attributes);
                        // this is for Cell
                        // oxw.WriteEndElement();
                    }
                }

                // this is for Row
                oxw.WriteEndElement();

                //End Write Header



                foreach (DataRow excelrow in excelData.Rows)
                {
                    i++;
                    oxa = new List<OpenXmlAttribute>();
                    // this is the row index
                    oxa.Add(new OpenXmlAttribute("r", null, i.ToString()));

                    oxw.WriteStartElement(new Row(), oxa);

                    foreach (DataColumn column in excelData.Columns)
                    {
                        if (column.ColumnName != "row_number" &&
                        column.ColumnName != "TotalRows" &&
                        column.ColumnName != "TotalPages" &&
                        column.ColumnName != "row_key")
                        {
                            //oxa = new List<OpenXmlAttribute>();
                            // this is the data type ("t"), with CellValues.String ("str")
                            //oxa.Add(new OpenXmlAttribute("t", null, "str"));

                            // it's suggested you also have the cell reference, but
                            // you'll have to calculate the correct cell reference yourself.
                            // Here's an example:
                            //oxa.Add(new OpenXmlAttribute("r", null, "A1"));

                            //oxw.WriteStartElement(new Cell(), oxa);

                            // oxw.WriteElement(new CellValue(excelrow[column.ColumnName].ToString()));

                            var s = column.DataType;
                            if (s.Name == "Int32" || s.Name == "Int16" || s.Name == "Decimal" || s.Name == "Double" || s.Name == "Float")
                                WriteCellValueSax(oxw, excelrow[column.ColumnName].ToString(), CellValues.Number);
                            else if (s.Name == "Date" || s.Name == "DateTime")
                            {
                                if (excelrow[column.ColumnName].ToString() != "")
                                    WriteCellValueSax(oxw, ((DateTime)excelrow[column.ColumnName]).ToString("dd-MM-yyyy"), CellValues.Date);
                                else WriteCellValueSax(oxw, excelrow[column.ColumnName].ToString(), CellValues.SharedString);
                            }
                            else WriteCellValueSax(oxw, excelrow[column.ColumnName].ToString(), CellValues.SharedString);
                            // this is for Cell
                            //oxw.WriteEndElement();
                        }
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
                    Name = sheetName,
                    SheetId = 1,
                    Id = xl.WorkbookPart.GetIdOfPart(wsp)
                });

                // this is for Sheets
                oxw.WriteEndElement();
                // this is for Workbook
                oxw.WriteEndElement();
                oxw.Close();


                CreateShareStringPart(xl.WorkbookPart);
                SaveCustomStylesheet(xl.WorkbookPart);
                xl.Close();
            }
        }

        private void CreateShareStringPart(WorkbookPart workbookPart)
        {
            if (_shareStringMaxIndex > 0)
            {
                var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                using (var writer = OpenXmlWriter.Create(sharedStringPart))
                {
                    writer.WriteStartElement(new SharedStringTable());
                    foreach (var item in _shareStringDictionary)
                    {
                        writer.WriteStartElement(new SharedStringItem());
                        writer.WriteElement(new Text(item.Key));
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }
            }
        }


        private void WriteCellValueSax(OpenXmlWriter writer, string cellValue,
            CellValues dataType, List<OpenXmlAttribute> attributes = null)
        {
            switch (dataType)
            {
                case CellValues.InlineString:
                    {
                        if (attributes == null)
                        {
                            attributes = new List<OpenXmlAttribute>();
                        }
                        attributes.Add(new OpenXmlAttribute("t", null, "inlineStr"));
                        writer.WriteStartElement(new Cell(), attributes);
                        writer.WriteElement(new InlineString(new Text(cellValue)));
                        writer.WriteEndElement();
                        break;
                    }
                case CellValues.SharedString:
                    {
                        if (attributes == null)
                        {
                            attributes = new List<OpenXmlAttribute>();
                        }
                        attributes.Add(new OpenXmlAttribute("t", null, "s"));//shared string type
                        writer.WriteStartElement(new Cell(), attributes);
                        if (!_shareStringDictionary.ContainsKey(cellValue))
                        {
                            _shareStringDictionary.Add(cellValue, _shareStringMaxIndex);
                            _shareStringMaxIndex++;
                        }

                        //writing the index as the cell value
                        writer.WriteElement(new CellValue(_shareStringDictionary[cellValue].ToString()));

                        writer.WriteEndElement();//cell

                        break;
                    }
                case CellValues.Date:
                    {
                        if (attributes == null)
                        {
                            attributes = new List<OpenXmlAttribute>();
                            attributes.Add(new OpenXmlAttribute("s", null, "1"));//data style
                            writer.WriteStartElement(new Cell() { DataType = CellValues.Number }, attributes);
                        }
                        else
                        {
                            writer.WriteStartElement(new Cell() { DataType = CellValues.Number }, attributes);
                        }

                        writer.WriteElement(new CellValue(DateTime.ParseExact(cellValue, "dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture).ToOADate().ToString()));

                        writer.WriteEndElement();

                        break;
                    }
                case CellValues.Boolean:
                    {
                        if (attributes == null)
                        {
                            attributes = new List<OpenXmlAttribute>();
                        }
                        attributes.Add(new OpenXmlAttribute("t", null, "b"));//boolean type
                        writer.WriteStartElement(new Cell(), attributes);
                        writer.WriteElement(new CellValue(cellValue == "True" ? "1" : "0"));
                        writer.WriteEndElement();
                        break;
                    }
                case CellValues.Number:
                    {
                        if (attributes == null)
                        {
                            attributes = new List<OpenXmlAttribute>();
                        }
                    
                        //attributes.Add(new OpenXmlAttribute("s", null, "3"));
                        writer.WriteStartElement(new Cell(), attributes);
                        writer.WriteElement(new CellValue(cellValue.ToString()));
                        writer.WriteEndElement();
                        break;
                    }
                default:
                    {
                        if (attributes == null)
                        {
                            writer.WriteStartElement(new Cell() { DataType = dataType });
                        }
                        else
                        {
                            writer.WriteStartElement(new Cell() { DataType = dataType }, attributes);
                        }
                        writer.WriteElement(new CellValue(cellValue));

                        writer.WriteEndElement();

                        break;
                    }
            }
        }



        private Stylesheet CreateDefaultStylesheet()
        {

            Stylesheet ss = new Stylesheet();

            Fonts fts = new Fonts();
            DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
            FontName ftn = new FontName();
            ftn.Val = "Times New Roman";
            FontSize ftsz = new FontSize();
            ftsz.Val = 11;
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            //ft.Bold = new Bold();
            fts.Append(ft);

            ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
            ftn = new FontName();
            ftn.Val = "Times New Roman";
            ftsz = new FontSize();
            ftsz.Val = 12;
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            ft.Bold = new Bold();
            fts.Append(ft);
            fts.Count = (uint)fts.ChildElements.Count;

            Fills fills = new Fills();
            Fill fill;
            PatternFill patternFill;

            //default fills used by Excel, don't changes these

            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.None;
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Gray125;
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);



            fills.Count = (uint)fills.ChildElements.Count;

            Borders borders = new Borders();
            Border border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.BottomBorder = new BottomBorder();
            border.DiagonalBorder = new DiagonalBorder();
            borders.Append(border);
            borders.Count = (uint)borders.ChildElements.Count;

            CellStyleFormats csfs = new CellStyleFormats();
            CellFormat cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            csfs.Append(cf);
            csfs.Count = (uint)csfs.ChildElements.Count;


            CellFormats cfs = new CellFormats();

            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cfs.Append(cf);



            var nfs = new NumberingFormats();



            nfs.Count = (uint)nfs.ChildElements.Count;
            cfs.Count = (uint)cfs.ChildElements.Count;

            ss.Append(nfs);
            ss.Append(fts);
            ss.Append(fills);
            ss.Append(borders);
            ss.Append(csfs);
            ss.Append(cfs);

            CellStyles css = new CellStyles(
                new CellStyle()
                {
                    Name = "Normal",
                    FormatId = 0,
                    BuiltinId = 0,
                }
                );

            css.Count = (uint)css.ChildElements.Count;
            ss.Append(css);

            DifferentialFormats dfs = new DifferentialFormats();
            dfs.Count = 0;
            ss.Append(dfs);

            TableStyles tss = new TableStyles();
            tss.Count = 0;
            tss.DefaultTableStyle = "TableStyleMedium9";
            tss.DefaultPivotStyle = "PivotStyleLight16";
            ss.Append(tss);
            return ss;
        }


        public void SaveCustomStylesheet(WorkbookPart workbookPart)
        {

            //get a copy of the default excel style sheet then add additional styles to it
            var stylesheet = CreateDefaultStylesheet();

            // ***************************** Fills *********************************
            var fills = stylesheet.Fills;

            //header fills background color
            var fill = new Fill();
            var patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            patternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("D0D0D0") };
            //patternFill.BackgroundColor = new BackgroundColor() { Indexed = 64 };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);
            fills.Count = (uint)fills.ChildElements.Count;

            // *************************** numbering formats ***********************
            var nfs = stylesheet.NumberingFormats;
            //number less than 164 is reserved by excel for default formats
            uint iExcelIndex = 165;
            NumberingFormat nf;
            nf = new NumberingFormat();
            nf.NumberFormatId = iExcelIndex++;
            nf.FormatCode = @"[$-409]m/d/yy\ h:mm\ AM/PM;@";
            nfs.Append(nf);

            nfs.Count = (uint)nfs.ChildElements.Count;

            //************************** cell formats ***********************************
            var cfs = stylesheet.CellFormats;//this should already contain a default StyleIndex of 0

            var cf = new CellFormat();// Date time format is defined as StyleIndex = 1
            cf.NumberFormatId = nf.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = true;
            cfs.Append(cf);

            cf = new CellFormat();// Header format is defined as StyleINdex = 2
            cf.NumberFormatId = 0;
            cf.FontId = 1;
            cf.FillId = 2;
            cf.ApplyFill = true;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cfs.Append(cf);


            cfs.Count = (uint)cfs.ChildElements.Count;

            var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var style = workbookStylesPart.Stylesheet = stylesheet;
            style.Save();

        }
    }
}
