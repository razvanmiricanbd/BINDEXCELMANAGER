using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Data.SqlTypes;
using System.Text.RegularExpressions;
using System.Data;
using NPOI.SS.Util;

namespace GenerareMachete
{
    public class GeneratorMacheteCorep
    {
        private  bool isCursor;
        private  bool isLastRow = false;
        private  int lastRow;
        private  int firstRow;
        private  int errorRowPos;
        private  int lastCell;
        private  int errorCellPos;
        private  int startIndex, endIndex, endCursorIndex, startCursorIndex;
        public  int i, j, q, sheetNumber, columnIndex;
        public  int rowCounter = 0;                                                            // Numarul de randuri noi inserate in sheet in cazul unui cursor
        public  int rowShiftAmount = 10000;

        private  string sheetName;
        private  string errorSheetName;
        private  string cellContent;
        private  string cellFormula;
        private  string cellPrefix;
        private  string cellSufix = "";
        private  string cursorName = "";
        private  string numberRegExp = "^(((-?(([1-9]+[0-9]*)|0))?\\.[0-9]+)|((-?[1-9]+[0-9]*)|0))$";

        private  string procedureCallString;
        private  string executionId;
        private  string sql;
        private  string cursorSql;
        private  string connectionString;

        private  string fileOut;
        private  string folderOut;
        private  string fileIn;
        private  string folderIn;


        private  CellType cellType;
        private  string columnName;
        private  string columnType;
        private  string sqlResult;
        private  SqlDateTime data;
        private  DataTable rsmd;

        private string cloneSheetQuerry;
        private static Dictionary<string, string> sheetNames = new Dictionary<string, string>();
        private static string[] cloneReportArray;
        private static String sheetNameClone = "";

        public GeneratorMacheteCorep() {

        }

        public void Generate(string pfolderIn, string pfileIn, string pfolderOut, string pfileOut, string pexecutionId, string pConnection)
        {


            List<ICell> cellList = new List<ICell>();                                       // Lista de celule pentru procesarea cursorului
            List<String> cellListFormula = new List<String>();                              // Lista de formule aferente celulelor
            List<CellType> cellListType = new List<CellType>();

            SqlCommand command = null;
            SqlCommand commandCursor = null;
            SqlDataReader rs = null;
            SqlDataReader rsCursor = null;
            SqlDataReader rsCloneSheets = null;
            SqlDataReader rsCountryCodes = null;
            SqlConnection sqlConnection = null;
            FileStream outputFile = null;


            //throw new Exception("Test daca merge ceva !!");


            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            IRow newRow = null;
            ICell cell = null;
            ICell newCell = null;
            //try
            //{
            //executionId = "22072";
            folderIn = pfolderIn;// args[0];
            fileIn = pfileIn;//args[1];
            folderOut = pfolderOut;//args[2];
            fileOut = pfileOut;//args[3];
            executionId = pexecutionId;//args[4];

            if (folderOut[folderOut.Length - 1] != Path.DirectorySeparatorChar)
                folderOut = folderOut + Path.DirectorySeparatorChar;

            if (folderIn[folderIn.Length - 1] != Path.DirectorySeparatorChar)
                folderIn = folderIn + Path.DirectorySeparatorChar;
            Console.WriteLine("Folder in : " + folderIn + "        Folder out : " + folderOut);

            /* connectionString = "Data Source=ROUD12VTDB02;" +
                                  "Initial Catalog = bind; " +
                                  "trusted_connection = true";
             */

            cloneSheetQuerry = "select xs.sheet_name, xs.sheet_name_bind\r\n" +
                              "  from rep_xls_grp_mapping r, \r\n" +
                              "       rep_executions_v v,\r\n" +
                              "       rep_xls_sheets xs,\r\n" +
                              "       rep_xls_files f\r\n" +
                              " where v.reports_group_id = r.reports_group_id\r\n" +
                              "   and xs.file_id = r.file_id\r\n" +
                              "   and f.file_id = r.file_id\r\n" +
                              "   and v.execution_id = " + executionId + "\r\n" +
                              "   and xs.sheet_is_breakdown = 'Y'\r\n" +
                              "   and f.file_name = '" + fileIn + "'";

            connectionString = pConnection;
            sqlConnection = new SqlConnection(connectionString);
            sqlConnection.Open();

            Console.WriteLine("Connection successful!!");

            command = new SqlCommand();
            command.Connection = sqlConnection;


            SqlCommand commandRsCursor = new SqlCommand();
            commandRsCursor.Connection = sqlConnection;

            commandCursor = new SqlCommand();
            commandCursor.Connection = sqlConnection;


            /*---------------- Start Clone Sheets ----------------*/
            //try
            //{
            //    command.CommandText = cloneSheetQuerry;
            //    rsCloneSheets = command.ExecuteReader();

            //    if (rsCloneSheets.HasRows)
            //        while (rsCloneSheets.Read())
            //        {
            //            sheetNames.Add(rsCloneSheets.GetSqlString(1).ToString(), rsCloneSheets.GetSqlString(2).ToString());
            //        }

            //}
            //catch (Exception ex)
            //{
            //    if (rsCloneSheets != null)
            //        rsCloneSheets.Close();

            //    sheetNames = null;

            //}

            //rsCloneSheets.Close();

            /*---------------- End Clone Sheets ----------------*/

            using (FileStream stream = new FileStream(folderIn + fileIn, FileMode.Open, FileAccess.ReadWrite))
            {
                Console.WriteLine("Open excel done!!");
                workbook = WorkbookFactory.Create(stream);

                int sheetCount = workbook.NumberOfSheets;

                for (sheetNumber = 0; sheetNumber < sheetCount; sheetNumber++)
                {
                    isLastRow = false;
                    sheet = workbook.GetSheetAt(sheetNumber);
                    isCursor = false;
                    firstRow = sheet.FirstRowNum;
                    lastRow = sheet.LastRowNum;
                    sheetName = sheet.SheetName;
                    errorSheetName = sheetName;

                    /*---------------- Start Clone Sheets ----------------*/
                    //if (sheetNames != null && sheetNames.ContainsKey(sheetName)/*   sheetName.equals(corepCloneSheet)*/)
                    //{

                    //    string sheetCountryCodeQuerry = "select distinct xst.table_b272_cell country_cell\r\n" +
                    //                                    "  from rep_xls_grp_mapping r,  \r\n" +
                    //                                    "       rep_executions_v v, \r\n" +
                    //                                    "       rep_xls_sheets xs, \r\n" +
                    //                                    "       rep_xls_files f,\r\n" +
                    //                                    "       rep_xls_sheet_tables xst\r\n" +
                    //                                    " where v.reports_group_id = r.reports_group_id \r\n" +
                    //                                    "   and xs.file_id = r.file_id \r\n" +
                    //                                    "   and f.file_id = r.file_id\r\n" +
                    //                                    "   and xst.sheet_id = xs.sheet_id\r\n" +
                    //                                    "   and xs.sheet_name = '" + sheetName + "'\r\n" +
                    //                                    "   and v.execution_id = " + executionId + "\r\n" +
                    //                                    "   and xs.sheet_is_breakdown = 'Y' \r\n" +
                    //                                    "   and f.file_name = '" + fileIn + "'";


                    //    SqlCommand cmd = new SqlCommand("rep_xl_sheets_cloning", sqlConnection);
                    //    cmd.CommandType = CommandType.StoredProcedure;
                    //    cmd.Parameters.AddWithValue("@p_report_code", sheetNames[sheetName]);
                    //    cmd.Parameters.AddWithValue("@p_execution_id", Int32.Parse(executionId));

                    //    cmd.ExecuteNonQuery();

                    //    string reportNames = cmd.Parameters["@p_reports"].Value.ToString();



                    //    if (reportNames != null && reportNames != "")
                    //        cloneReportArray = reportNames.Split(',');
                    //    else
                    //        cloneReportArray = null;

                    //    if (cloneReportArray != null)
                    //    {
                    //        foreach (string cloneSheetName in cloneReportArray)
                    //        {
                    //            if (cloneSheetName.Contains("_"))
                    //            {

                    //                String countryCode = cloneSheetName.Substring(cloneSheetName.Length - 2);
                    //                String finalSheetName = sheetName + "_" + countryCode;
                    //                workbook.CloneSheet(sheetNumber);
                    //                workbook.SetSheetName(sheetCount, finalSheetName);
                    //                workbook.SetSheetOrder(finalSheetName, sheetNumber + 1);

                    //                try
                    //                {
                    //                    command.CommandText = sheetCountryCodeQuerry;
                    //                    rsCountryCodes = command.ExecuteReader();

                    //                    while (rsCountryCodes.Read())
                    //                    {
                    //                        try
                    //                        {
                    //                            CellReference cellRef = new CellReference(rsCountryCodes.GetString(1));
                    //                            workbook.GetSheetAt(sheetNumber + 1).GetRow(cellRef.Row).GetCell(cellRef.Col).SetCellValue(countryCode);
                    //                        }
                    //                        catch (Exception e)
                    //                        {

                    //                        }
                    //                    }
                    //                }
                    //                catch (Exception e)
                    //                {

                    //                }


                    //                sheetCount++;


                    //            }
                    //        }

                    //    }
                    //    sheetNameClone = sheetName;
                    //    //cloneReportArray = null;
                    //}

                    /*---------------- End Clone Sheets ----------------*/


                    if (sheetName.ToUpper().Equals("DPM") == false && sheetName.ToUpper().Equals("VALIDARI") == false)
                    {

                        for (i = firstRow; i <= sheet.LastRowNum; i++)
                        {
                            errorRowPos = i;
                            row = sheet.GetRow(i);
                            isCursor = false;
                            if (row != null)
                            {
                                lastCell = row.LastCellNum;

                                for (j = 0; j <= lastCell; j++)
                                {
                                    errorCellPos = j;
                                    cell = row.GetCell(j);

                                    if (cell != null)
                                    {
                                        cellContent = null;
                                        cellType = cell.CellType;
                                        cellFormula = "";

                                        // Procesare string in cazul unei celule de tip formula sau cursor
                                        if (cellType == CellType.String)
                                            cellFormula = cell.StringCellValue;
                                        else if (cellType == CellType.Formula)
                                            cellFormula = cell.CellFormula;

                                        startIndex = cellFormula.IndexOf("<%=");
                                        endIndex = cellFormula.IndexOf("%>");
                                        startCursorIndex = cellFormula.ToLower().IndexOf("<%for");
                                        endCursorIndex = cellFormula.ToLower().IndexOf("<%end loop;%>");

                                        // Daca este cursor se adauga celula in lista de celule pentru inserarea randurilor
                                        if (isCursor == true && endCursorIndex == -1)
                                        {
                                            cellList.Add(cell);
                                            cellListType.Add(cell.CellType);
                                            cellListFormula.Add(cellFormula);
                                        }

                                        // Procesare celule
                                        if (cellType == CellType.String || cellType == CellType.Formula)
                                        {
                                            if (startIndex != -1 && endIndex != -1 && endIndex > startIndex)
                                            {
                                                // Prefixul/Sufixul formulei ex: = VALUE( )
                                                cellPrefix = cellFormula.Substring(0, startIndex);
                                                cellSufix = cellFormula.Substring(endIndex + 2);

                                                // Proceseare interiorul formulei
                                                if (cellFormula.Substring(startIndex + 3, cursorName.Length + 1).Equals(cursorName + ".", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    columnName = cellFormula.Substring(startIndex + 3 + cursorName.Length + 1, endIndex - startIndex - cursorName.Length - 4);
                                                    columnIndex = rsCursor.GetOrdinal(columnName);

                                                    for (int cIndex = 0; cIndex < rsCursor.FieldCount; cIndex++)
                                                    {

                                                        // Tipurile celulei in functie de datele aduse de query
                                                        if (rsCursor.GetName(cIndex).ToLower().Equals(columnName.ToLower()))
                                                        {
                                                            columnType = rsCursor.GetDataTypeName(cIndex).ToLower();
                                                            if (columnType.Equals("varchar") || columnType.Equals("nvarchar"))
                                                            {
                                                                sqlResult = Convert.ToString(rsCursor[columnName]);

                                                                if (rsCursor.IsDBNull(columnIndex) == true)
                                                                    sqlResult = "";

                                                                break;
                                                            }
                                                            else if (columnType.Equals("number"))
                                                            {
                                                                /*-----------------------------------------------------------*/
                                                                // if (rsCursor.)
                                                                //    sqlResult = Convert.ToString(rsCursor.GetInt64(columnIndex));
                                                                // else
                                                                // sqlResult = Convert.ToString(rsCursor.GetDouble(columnIndex));
                                                                sqlResult = rsCursor[columnName].ToString();
                                                                if (rsCursor.IsDBNull(columnIndex) == true)
                                                                    sqlResult = "";

                                                                break;
                                                            }
                                                            else if (columnType.Equals("date"))
                                                            {
                                                                data = rsCursor.GetDateTime(columnIndex);
                                                                if (data.IsNull == false)
                                                                    sqlResult = Convert.ToDateTime(data).ToString("dd-MM-yyyy");
                                                                else
                                                                    sqlResult = "";
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                sqlResult = rsCursor[columnName].ToString();
                                                                if (rsCursor.IsDBNull(columnIndex) == true)
                                                                    sqlResult = "";
                                                            }
                                                        }

                                                    }
                                                }
                                                else
                                                {
                                                    // In cazul unei proceduri se inlocuieste :p_execution_id
                                                    procedureCallString = cellFormula.Substring(startIndex + 3, endIndex - startIndex - 3);

                                                    procedureCallString = Regex.Replace(procedureCallString, ":p_execution_id", executionId, RegexOptions.IgnoreCase);
                                                    procedureCallString = Regex.Replace(procedureCallString, "rep_general_pkg.", "rep_general_pkg$", RegexOptions.IgnoreCase);
                                                    procedureCallString = Regex.Replace(procedureCallString, "getreportctrycode", "BIND.dbo.getReportCtryCode", RegexOptions.IgnoreCase);
                                                    procedureCallString = Regex.Replace(procedureCallString, "getreportccycode", "BIND.dbo.getReportCcyCode", RegexOptions.IgnoreCase);
                                                    procedureCallString = Regex.Replace(procedureCallString, "getexecid", "BIND.dbo.getExecId", RegexOptions.IgnoreCase);

                                                    /* ---------------------- Start Clonare Sheets  ----------------------- */
                                                    //if (sheetName.StartsWith(sheetNameClone) && cloneReportArray != null &&
                                                    //   sheetNames[sheetNameClone] != null &&
                                                    //   cloneReportArray.ToList().Contains(sheetNames[sheetNameClone] + "_" + sheetName.Substring(sheetName.Length - 2)))
                                                    //{
                                                    //    procedureCallString = procedureCallString.Replace(sheetNames[sheetNameClone], sheetNames[sheetNameClone] + "_" + sheetName.Substring(sheetName.Length - 2));
                                                    //}
                                                    /* ---------------------- End Clonare Sheets  ----------------------- */


                                                    if (procedureCallString.Substring(0, 9).Equals("BIND.dbo.") == false)
                                                        command.CommandText = "select BIND.dbo." + procedureCallString + " value";

                                                    rs = command.ExecuteReader();
                                                 
                                                    if (rs.HasRows && rs.Read() == true)
                                                        sqlResult = rs["value"].ToString();
                                                    else
                                                        sqlResult = "";

                                                    rs.Close();
                                                }

                                                cellContent = cellPrefix + sqlResult + cellSufix;

                                                if (cellType == CellType.String)
                                                {
                                                    // Daca se poate converti la numar, si stringul nu incepe cu 0
                                                    if (Regex.IsMatch(cellContent, numberRegExp))
                                                    {
                                                        try
                                                        {
                                                            cell.SetCellValue(Double.Parse(cellContent));
                                                        }
                                                        catch (Exception e)
                                                        {
                                                            cell.SetCellValue(cellContent);
                                                        }
                                                    }
                                                    else
                                                        cell.SetCellValue(cellContent);
                                                }
                                                else if (cellType == CellType.Formula)
                                                {
                                                    if (Regex.IsMatch(cellContent, numberRegExp))
                                                        try
                                                        {
                                                            cell.SetCellValue(Double.Parse(cellContent));
                                                        }
                                                        catch (Exception e)
                                                        {
                                                            cell.SetCellFormula(cellContent);
                                                        }
                                                    else
                                                        cell.SetCellFormula(cellContent);
                                                    //evaluator.evaluateFormulaCell(cell);
                                                }
                                                else if (cellType == CellType.Numeric)
                                                {
                                                    try
                                                    {
                                                        cell.SetCellValue(Double.Parse(cellContent));
                                                    }
                                                    catch (Exception e)
                                                    {
                                                        cell.SetCellFormula(cellContent);
                                                    }
                                                }
                                            }
                                            else if (startCursorIndex > -1) //cazul in care am cursor
                                            {

                                                //if (rsCursor.IsClosed == false)
                                                //    rsCursor.Close();

                                                isCursor = true;

                                                cursorSql = cellFormula.Substring(cellFormula.IndexOf("(") + 1, cellFormula.LastIndexOf(")") - cellFormula.IndexOf("(") - 1);
                                                cursorName = cellFormula.Substring(cellFormula.IndexOf("for") + 3, cellFormula.IndexOf("in") - 1 - cellFormula.IndexOf("for") - 3).Trim();

                                                cursorSql = cursorSql.Replace(":p_execution_id", executionId);
                                                cursorSql = cursorSql.Replace(":P_EXECUTION_ID", executionId);
                                                cursorSql = cursorSql.Replace("REP_GENERAL_PKG.", "REP_GENERAL_PKG$");
                                                cursorSql = cursorSql.Replace("rep_general_pkg.", "rep_general_pkg$");

                                                command.CommandText = cursorSql;
                                                //command.Connection = sqlConnection;
                                                rsCursor = command.ExecuteReader();

                                                rsmd = rsCursor.GetSchemaTable();

                                                // Daca cursorul nu returneaza rezultate se elimina randul creat anterior
                                                if (rsCursor.Read() == false)
                                                {
                                                    sheet.RemoveRow(row);
                                                    cellList.Clear();
                                                    cellListFormula.Clear();
                                                    cellListType.Clear();

                                                    if (i < lastRow)
                                                        sheet.ShiftRows(i + 1, lastRow, -1);

                                                    i--;
                                                    //if (rsCursor != null)
                                                    //    rsCursor.Close();
                                                    break;
                                                }

                                                //if (rsCursor != null)
                                                //    rsCursor.Close();


                                                if (cellType == CellType.String)
                                                    cell.SetCellValue("");
                                                else if (cellType == CellType.Formula)
                                                {
                                                    cell.SetCellFormula("");
                                                    //  evaluator.evaluateFormulaCell(cell);
                                                }

                                            }
                                            else if (endCursorIndex > -1)                        // Eliminarea tagului de final cursor
                                            {
                                                isCursor = false;

                                                if (cellType == CellType.String)
                                                    cell.SetCellValue("");
                                                else if (cellType == CellType.Formula)
                                                {
                                                    cell.SetCellFormula("");
                                                    // evaluator.evaluateFormulaCell(cell);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (cellList.Count > 0)
                            {

                                // In procesarea cursorului se shifteaza randurile in jos pentru a adauga altele noi
                                if (i + 1 > lastRow)
                                {
                                    isLastRow = true;
                                }
                                if (!isLastRow)
                                {
                                    sheet.ShiftRows(i + 1, lastRow, rowShiftAmount);
                                }

                                rowCounter = 0;
                               // if(rsCursor.HasRows)
                                while (rsCursor.Read())
                                {

                                    // Daca numarul de randuri adaugate este mai mare sau egal cu valoarea shiftata anterior, atunci se mai executa inca o data
                                    if (i < lastRow && rowCounter >= rowShiftAmount && (!isLastRow)) //TODO
                                    {
                                        sheet.ShiftRows(i + 1, lastRow, rowShiftAmount);
                                        rowCounter = 0;
                                    }
                                    rowCounter++;

                                    newRow = sheet.CreateRow(++i);

                                    // Se creaza noul rand cu stilul celulelor din randul care contine cursorul
                                    for (q = 0; q < cellList.Count; q++)
                                    {
                                        newCell = newRow.CreateCell(cellList.ElementAt(q).ColumnIndex, cellList.ElementAt(q).CellType);
                                        newCell.CellStyle = cellList.ElementAt(q).CellStyle;

                                        //newCell.getCellStyle().cloneStyleFrom(cs);
                                        cellType = cellListType.ElementAt(q);
                                        cellFormula = null;
                                        cellContent = null;

                                        if (cellType == CellType.Boolean)
                                        {
                                            newCell.SetCellValue(cellList.ElementAt(q).BooleanCellValue);
                                        }
                                        else if (cellType == CellType.Numeric)
                                        {
                                            if (DateUtil.IsCellDateFormatted(newCell))
                                            {
                                                newCell.SetCellValue(cellList.ElementAt(q).DateCellValue);
                                            }
                                            else
                                            {
                                                newCell.SetCellValue(cellList.ElementAt(q).NumericCellValue);
                                            }
                                        }
                                        else if (cellType == CellType.Error)
                                        {
                                            newCell.SetCellValue(cellList.ElementAt(q).ErrorCellValue);
                                        }
                                        else if (cellType == CellType.String || cellType == CellType.Formula)
                                        {
                                            cellFormula = cellListFormula.ElementAt(q);

                                            startIndex = cellFormula.IndexOf("<%=");
                                            endIndex = cellFormula.IndexOf("%>");


                                            if (startIndex != -1 && endIndex != -1 && endIndex > startIndex)
                                            {
                                                cellPrefix = cellFormula.Substring(0, startIndex);
                                                cellSufix = cellFormula.Substring(endIndex + 2);

                                                if (cellFormula.Substring(startIndex + 3, cursorName.Length + 1).Equals(cursorName + ".", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    columnName = cellFormula.Substring(startIndex + 3 + cursorName.Length + 1, endIndex - startIndex - 4 - cursorName.Length);
                                                    columnIndex = rsCursor.GetOrdinal(columnName);

                                                    for (int cIndex = 0; cIndex < rsCursor.FieldCount; cIndex++)
                                                    {
                                                        if (rsCursor.GetName(cIndex).ToLower().Equals(columnName.ToLower()))
                                                        {
                                                            columnType = rsCursor.GetDataTypeName(cIndex).ToLower();

                                                            if (columnType.Equals("varchar"))
                                                            {
                                                                sqlResult = rsCursor[columnName].ToString();

                                                                if (rsCursor.HasRows == false)
                                                                    sqlResult = "";

                                                                break;
                                                            }
                                                            else if (columnType.Equals("number"))
                                                            {
                                                                //if (rsmd.getScale(cIndex) <= 0)
                                                                sqlResult = rsCursor[columnName].ToString();
                                                                // else
                                                                //    sqlResult = rsCursor[columnName].ToString();

                                                                if (rsCursor.HasRows == false)
                                                                    sqlResult = "";

                                                                break;
                                                            }
                                                            else if (columnType.Equals("date"))
                                                            {
                                                                data = rsCursor.GetDateTime(columnIndex);
                                                                if (data.IsNull == false)
                                                                    sqlResult = Convert.ToDateTime(data).ToString("dd-MM-yyyy");
                                                                else
                                                                    sqlResult = "";
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                sqlResult = rsCursor[columnName].ToString();
                                                                if (rsCursor.IsDBNull(columnIndex) == true)
                                                                    sqlResult = "";
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    procedureCallString = cellFormula.Substring(startIndex + 3, endIndex - startIndex - 3);
                                                    procedureCallString = procedureCallString.Replace(":p_execution_id", executionId);
                                                    procedureCallString = procedureCallString.Replace(":P_EXECUTION_ID", executionId);
                                                    procedureCallString = procedureCallString.Replace("REP_GENERAL_PKG.", "REP_GENERAL_PKG$");
                                                    procedureCallString = procedureCallString.Replace("rep_general_pkg.", "rep_general_pkg$");

                                                    commandCursor.CommandText = "select BIND.dbo." + procedureCallString + " value";
                                                    commandCursor.Connection = sqlConnection;
                                                    rs = commandCursor.ExecuteReader();

                                                    if (rs.HasRows == true)
                                                    {
                                                        rs.Read();
                                                        sqlResult = rs["value"].ToString();
                                                    }
                                                    else
                                                        sqlResult = "";

                                                    rs.Close();
                                                }

                                                cellContent = cellPrefix + sqlResult + cellSufix;

                                            }
                                            else
                                                cellContent = cellListFormula.ElementAt(q);


                                            if (cellType == CellType.String)
                                            {
                                                // Daca se poate convertii la numar, si stringul nu incepe cu 0
                                                if (Regex.IsMatch(cellContent, numberRegExp))
                                                    try
                                                    {
                                                        newCell.SetCellValue(Double.Parse(cellContent));
                                                    }
                                                    catch (Exception e)
                                                    {
                                                        newCell.SetCellValue(cellContent);

                                                    }
                                                else
                                                    newCell.SetCellValue(cellContent);
                                            }
                                            else if (cellType == CellType.Formula)
                                            {
                                                if (Regex.IsMatch(cellContent, numberRegExp))
                                                    try
                                                    {
                                                        newCell.SetCellValue(Double.Parse(cellContent));
                                                    }
                                                    catch (Exception e)
                                                    {
                                                        newCell.SetCellFormula(cellContent);
                                                    }
                                                else
                                                    newCell.SetCellFormula(cellContent);
                                                // evaluator.evaluateFormulaCell(newCell);
                                            }
                                            else if (cellType == CellType.Numeric)
                                            {
                                                try
                                                {
                                                    newCell.SetCellValue(Double.Parse(cellContent));
                                                }
                                                catch (Exception e)
                                                {
                                                    newCell.SetCellFormula(cellContent);
                                                }
                                            }
                                        }
                                    }

                                    lastRow = sheet.LastRowNum;
                                }
                                lastRow = sheet.LastRowNum;

                                // Aducerea randurilor shiftate la loc in cazul in care numarul de randuri adaugate de cursor < decat cel shiftat

                                if ((rowShiftAmount != rowCounter) && (i + 1 + rowCounter <= lastRow) && (!isLastRow))
                                {
                                    sheet.ShiftRows(i + 1 + (rowShiftAmount - rowCounter), lastRow, -(rowShiftAmount - rowCounter));
                                }

                                rsCursor.Close();
                                rsmd.Clear();
                                cellList.Clear();
                                cellListType.Clear();
                                cellListFormula.Clear();

                            }
                            lastRow = sheet.LastRowNum;
                        }
                    }
                    //sheetCount = wb.GetNumberOfSheets();
                }


                outputFile = new FileStream(folderOut + fileOut, FileMode.Create, FileAccess.Write);
                workbook.Write(outputFile);
                outputFile.Close();


                if (rs != null)
                    rs.Close();
                if (rsCursor != null)
                    rsCursor.Close();
                if (sqlConnection != null)
                    sqlConnection.Close();
                if (workbook != null)
                    workbook.Close();
                //stmt.Close();
                //conn.Close();

            }

            //}
            // catch (Exception e)
            //{
            //  Console.WriteLine(e.Message + "\n");
            // Console.Error.WriteLine(e.StackTrace + "\n --> Exception at Sheet : " + errorSheetName + ", Row : " + errorRowPos + ", Cell : " + errorCellPos);
            //}
            // finally
            // {
            if (outputFile != null)
                outputFile.Close();
            if (rs != null)
                rs.Close();
            if (rsCursor != null)
                rsCursor.Close();
            if (sqlConnection != null)
                sqlConnection.Close();
            if (workbook != null)
                workbook.Close();

            // }

        }
    }
  
}
