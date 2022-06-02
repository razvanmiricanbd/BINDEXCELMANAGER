using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Data.SqlTypes;
using System.Text.RegularExpressions;
using System.Data;


namespace GenerareMachete
{
    class Program
    {
        private static bool isCursor;

        private static int lastRow;
        private static int firstRow;
        private static int errorRowPos;
        private static int lastCell;
        private static int errorCellPos;
        private static int startIndex, endIndex, endCursorIndex, startCursorIndex;
        public static int i, j, q, sheetNumber, columnIndex;
        public static int rowCounter = 0;                                                            // Numarul de randuri noi inserate in sheet in cazul unui cursor
        public static int rowShiftAmount = 10000;

        private static string sheetName;
        private static string errorSheetName;
        private static string cellContent;
        private static string cellFormula;
        private static string cellPrefix;
        private static string cellSufix  = "";
        private static string cursorName = "";
        private static string numberRegExp = "^(((-?(([1-9]+[0-9]*)|0))?\\.[0-9]+)|((-?[1-9]+[0-9]*)|0))$";

        private static string procedureCallString;
        private static string executionId;
        private static string sql;
        private static string cursorSql;
        private static string connectionString;

        private static string fileOut;
        private static string folderOut;
        private static string fileIn;
        private static string folderIn;
        

        private static CellType cellType;
        private static string columnName;
        private static string columnType;
        private static string sqlResult;
        private static SqlDateTime data; 
        private static DataTable rsmd;
      
        static void Main(string[] args)
        {
            List<ICell> cellList = new List<ICell>();                                       // Lista de celule pentru procesarea cursorului
            List<String> cellListFormula = new List<String>();                              // Lista de formule aferente celulelor
            List<CellType> cellListType = new List<CellType>();

            SqlCommand command = null;
            SqlDataReader rs = null;
            SqlDataReader rsCursor = null;
            SqlConnection sqlConnection = null;
            FileStream outputFile = null;
          
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            IRow newRow = null;
            ICell cell = null;
            ICell newCell = null;
            try
            {
                //executionId = "22072";
                folderIn = args[0];
                fileIn = args[1];
                folderOut = args[2];
                fileOut = args[3];
                executionId = args[4];

                if (folderOut[folderOut.Length - 1] != Path.DirectorySeparatorChar)
                    folderOut = folderOut + Path.DirectorySeparatorChar;
                
                if (folderIn[folderIn.Length - 1] != Path.DirectorySeparatorChar)
                    folderIn = folderIn + Path.DirectorySeparatorChar;
                Console.WriteLine("Folder in : " + folderIn + "        Folder out : " + folderOut);

                connectionString = "Data Source=ROUD12VTDB02;" +
                                     "Initial Catalog = bind; " +
                                     "trusted_connection = true";

                sqlConnection = new SqlConnection(connectionString);
                sqlConnection.Open();

                Console.WriteLine("Connection successful!!");

                command = new SqlCommand();
                command.Connection = sqlConnection;              

                using (FileStream stream = new FileStream(folderIn + fileIn, FileMode.Open, FileAccess.ReadWrite))
                {
                    Console.WriteLine("Open excel done!!");
                    workbook = WorkbookFactory.Create(stream);

                    for (sheetNumber = 0; sheetNumber < workbook.NumberOfSheets; sheetNumber++)
                    {
                        sheet = workbook.GetSheetAt(sheetNumber);
                        isCursor = false;
                        firstRow = sheet.FirstRowNum;
                        lastRow = sheet.LastRowNum;
                        sheetName = sheet.SheetName;
                        errorSheetName = sheetName;

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

                                                        procedureCallString = Regex.Replace(procedureCallString, ":p_execution_id", executionId , RegexOptions.IgnoreCase);
                                                        procedureCallString = Regex.Replace(procedureCallString, "rep_general_pkg.", "rep_general_pkg$", RegexOptions.IgnoreCase);
                                                        procedureCallString = Regex.Replace(procedureCallString, "getreportctrycode", "bind.dbo.getReportCtryCode", RegexOptions.IgnoreCase);
                                                        procedureCallString = Regex.Replace(procedureCallString, "getreportccycode", "bind.dbo.getReportCcyCode", RegexOptions.IgnoreCase);
                                                        procedureCallString = Regex.Replace(procedureCallString, "getexecid", "bind.dbo.getExecId", RegexOptions.IgnoreCase);

                                                        if(procedureCallString.Substring(0,9).Equals("bind.dbo.") == false)
                                                            command.CommandText = "select bind.dbo." + procedureCallString + " value";
                                                        
                                                        rs = command.ExecuteReader();

                                                        if (rs.Read() == true)
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
                                                    command.Connection = sqlConnection;
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

                                                        break;
                                                    }

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
                                    sheet.ShiftRows(i + 1, lastRow, rowShiftAmount);
                                    rowCounter = 0;
                                    while (rsCursor.Read())
                                    {

                                        // Daca numarul de randuri adaugate este mai mare sau egal cu valoarea shiftata anterior, atunci se mai executa inca o data
                                        if (i < lastRow && rowCounter >= rowShiftAmount) //TODO
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

                                                        command.CommandText = "select bind.dbo." + procedureCallString + " value";
                                                        command.Connection = sqlConnection;
                                                        rs = command.ExecuteReader();

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
                                    if (rowShiftAmount != rowCounter)
                                        sheet.ShiftRows(i + 1 + (rowShiftAmount - rowCounter), lastRow, - (rowShiftAmount - rowCounter));
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

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message+"\n");
                Console.Error.WriteLine(e.StackTrace + "\n --> Exception at Sheet : " + errorSheetName + ", Row : " + errorRowPos + ", Cell : " + errorCellPos);
            }
            finally
            {
                if(outputFile != null)
                    outputFile.Close();
                if (rs != null)
                    rs.Close();
                if (rsCursor != null)
                    rsCursor.Close();
                if (sqlConnection != null)
                    sqlConnection.Close();
                if (workbook != null)
                    workbook.Close();
                
            }
        }
    }
}
