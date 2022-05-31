using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using NBDCrudWrapper;
using OpenXmlNbdWrapper;
using GenerareMachete;
using System.Globalization;
using System.IO;

namespace BindExcelManager
{
    public partial class BindExcelManager : ServiceBase
    {
        string _connection = null;
        string _schema_name = null;
        string xmlError = "";
        public BindExcelManager()
        {
            InitializeComponent();
            _connection = System.Configuration.ConfigurationManager.
                ConnectionStrings["BIND"].ConnectionString;
            string _schema_name = "bind";
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("BindExcelManager"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "BindExcelManager", "Application");
            }
            eventLog1.Source = "BindExcelManager";
            eventLog1.Log = "Application";

            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 10000; // 10 seconds  
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();

            //OnStart(new string[] { });
            //OnTimer(null, null);
        }

        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("BindExcelManager succesfully started");
            _connection = System.Configuration.ConfigurationManager.
                ConnectionStrings["BIND"].ConnectionString;
            //MSSqlEngine engine = new MSSqlEngine(_connection);
            //MSSqlParameter[] parameters = null;

           // DataTable tb = engine.RunProcedureQuery(@"BINDEXCELSERVICE$GET_FILES", parameters);

            eventLog1.WriteEntry("BindExcelManager succesfully started !");
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("BindExcelManager succesfully stopped");
        }

        private void OnTimer(object sender, EventArgs e)
        {
            //eventLog1.WriteEntry("BindExcelManager OnTimer");
            try
            {
                scanForFile();
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("ERROR: " + ex.Message);
            }
        }

        private void scanForFile() {

            //eventLog1.WriteEntry("BindExcelManager ScanForFile begin");
            _connection = System.Configuration.ConfigurationManager.
                ConnectionStrings["BIND"].ConnectionString;
            MSSqlEngine engine = new MSSqlEngine(_connection);
            MSSqlParameter[] parameters = null;

            //eventLog1.WriteEntry("BindExcelManager ScanForFile dupa connection");

            DataTable tbexcel = engine.RunProcedureQuery(@"BINDEXCELSERVICE$GET_FILES", parameters);
            if (tbexcel.Rows.Count > 0)
            {
                foreach (DataRow row in tbexcel.Rows)
                {
                    //eventLog1.WriteEntry("Process file:" + row["ExportProcedure"]);
                    // Set status as Running
                    try
                    {
                        /*
                        parameters = new MSSqlParameter[2];
                        parameters[0] = new MSSqlParameter
                        {
                            Name = "@fileid",
                            Value = ((int)row["FileId"])
                        };
                        parameters[1] = new MSSqlParameter
                        {
                            Name = "@status",
                            Value = "RUNNING"
                        };
                        engine.RunProcedureStatment("[BINDEXCELSERVICE$SET_FILES_STATUS]", parameters);
                        */
                        // Run Procedure that generates data

                        if (row["screen_id"] == DBNull.Value
                            || row["screen_id"] == null
                            || row["screen_id"].ToString() == ""
                            )
                        {
                            eventLog1.WriteEntry("Process file started:" + row["ExportProcedure"]);
                            parameters = new MSSqlParameter[3];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@ref_date",
                                Value = ((DateTime)row["ref_date"]).ToString("yyyy-MM-dd")
                            };
                            parameters[1] = new MSSqlParameter
                            {
                                Name = "@PageNumber",
                                Value = DBNull.Value
                            };
                            parameters[2] = new MSSqlParameter
                            {
                                Name = "@RowspPage",
                                Value = DBNull.Value
                            };

                            DataTable tbdata = engine.RunProcedureQuery((string)row["ExportProcedure"], parameters);


                            // NbdOpenXmlExcelWriter.WriteExcelOutputForTable((string)row["FilePath"], tb, (string)row["ExportName"]);
                            NbdOpenXmlExcelWriter writer = new NbdOpenXmlExcelWriter();

                            writer.WriteExcelOutputForTableSAX((string)row["FilePath"], tbdata, (string)row["ExportName"]);

                            parameters = new MSSqlParameter[2];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@fileid",
                                Value = ((int)row["FileId"])
                            };
                            parameters[1] = new MSSqlParameter
                            {
                                Name = "@status",
                                Value = "FINISH"
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$SET_FILES_STATUS]", parameters);
                            eventLog1.WriteEntry("Process file succesfully:" + row["ExportProcedure"]);

                        }// New dinamic screen with varios parameters not just date.
                        else {
                            // Get the parameter values
                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@fileid",
                                Value = ((int)row["FileId"])
                            };
                            DataTable tbparam = engine.RunProcedureQuery("BINDAPP$GETExecParameters", parameters);

                            if (tbparam.Rows.Count > 0)
                            {
                                int i = 0;
                                parameters = new MSSqlParameter[tbparam.Rows.Count];
                                foreach (DataRow paramrow in tbparam.Rows)
                                {
                                    parameters[i] = new MSSqlParameter
                                    {
                                        Name = paramrow["ParameterName"].ToString(),
                                        Value = paramrow["ParameterValue"].ToString()
                                    };
                                    i++;
                                }
                            }
                            else parameters = null;

                            eventLog1.WriteEntry("Process file started:" + row["ExportProcedure"]);
                          

                            DataTable tbdata = engine.RunProcedureQuery((string)row["ExportProcedure"], parameters);


                            // NbdOpenXmlExcelWriter.WriteExcelOutputForTable((string)row["FilePath"], tb, (string)row["ExportName"]);
                            NbdOpenXmlExcelWriter writer = new NbdOpenXmlExcelWriter();

                            writer.WriteExcelOutputForTableSAX((string)row["FilePath"], tbdata, (string)row["ExportName"]);

                            parameters = new MSSqlParameter[2];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@fileid",
                                Value = ((int)row["FileId"])
                            };
                            parameters[1] = new MSSqlParameter
                            {
                                Name = "@status",
                                Value = "FINISH"
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$SET_FILES_STATUS]", parameters);
                            eventLog1.WriteEntry("Process file succesfully:" + row["ExportProcedure"]);
                        }
                    }
                    catch (Exception ex) {

                        if (row != null && row["FileId"] != null)
                        {
                            parameters = new MSSqlParameter[2];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@fileid",
                                Value = ((int)row["FileId"])
                            };
                            parameters[1] = new MSSqlParameter
                            {
                                Name = "@status",
                                Value = "CRUSH"
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$SET_FILES_STATUS]", parameters);
                        }
                        eventLog1.WriteEntry("Error processing file" + ex.Message);
                    }
                }
            }//files

            //eventLog1.WriteEntry("BindExcelManager ScanForFile dupa get_files");

            parameters = null;
            DataTable tbexcels = engine.RunProcedureQuery(@"BINDEXCELSERVICE$GET_EXCELFILES", parameters);
            string newExecution = "";
            if (tbexcels.Rows.Count > 0)
            {
                foreach (DataRow row in tbexcels.Rows)
                {
                    //eventLog1.WriteEntry("Process file:" + row["ExportProcedure"]);
                    // Set status as Running
                    try
                    {
                        string xmlPathOnServer = "";
                        DataTable xmlPathTable = engine.RunQuerry("select parameter_value from REP_GENERAL_PARAMETERS where parameter_code = 'XML_OUTPUT_PATH'", null);

                        foreach (DataRow item in xmlPathTable.Rows)
                            xmlPathOnServer = item["parameter_value"].ToString();

                        parameters = new MSSqlParameter[1];
                        parameters[0] = new MSSqlParameter
                        {
                            Name = "@message",
                            Value = "Start generate excel + xml!"
                        };
                        engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                        string fileinpath = row["Excel_Input_Path"].ToString();
                        string filein = row["Excel_Input_Template1"].ToString();
                        string fileoutpath = row["Excel_Output_Path"].ToString();
                        string fileout = row["Excel_file"].ToString();
                        int execution = ((int)row["Execution_id"]);
                        string refDate = row["ref_date"].ToString();
                        string excelForXml = row["excel_input"].ToString();
                        string machetaExcelXml = row["Excel_For_Xml"].ToString();
                        newExecution = row["new_execution"].ToString();
                        
                        #region Parametrii xml

                        string reportsGroupCode = "xyz";
                        string pathToJar = fileinpath;
                        string pathToJRE = fileinpath + @"\jre\bin\java";
                        string pathToUploadExcel = pathToJar;
                        string fileName = excelForXml;
                        string fileNameWithoutExtension = fileName.Substring(0, fileName.LastIndexOf('.'));
                        string xmFileName = fileName.Substring(0, fileName.LastIndexOf('.')) + ".xml";
                        string xmlUniqueId = "";
                        string excelUniqueId = "";
                        string classPath = fileinpath + "GenMacheteBind_lib";
                        string xlsOutName = execution.ToString() + "_" + filein.Replace(" ", "").Replace("MACHETA_", "").Replace(".xlsx", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

                        //Procedura pentru adus consolidat si ref date pe baza executiei
                        string consolidated = "N"/*Request["consolidated"]*/;

                        DateTime refDateDT = DateTime.ParseExact(refDate, "dd-MM-yyyy", CultureInfo.InvariantCulture);
                        string refDateFormat = refDateDT.ToString("yyyy-MM-dd");

                        #endregion


                        parameters = new MSSqlParameter[1];
                        parameters[0] = new MSSqlParameter
                        {
                            Name = "@message",
                            Value = "Start generate excel -> execution_id = " + execution
                        };

                        engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                        /*GeneratorMacheteCorep g = new GeneratorMacheteCorep();*/

                        if (filein != null && filein != "")
                        {

                            /*g.Generate(fileinpath, filein, fileoutpath, fileout, execution.ToString(), _connection);*/
                            //string jarParams = $"-jar \"{pathToJar}\\XlsToXmlConverter.jar\" {reportsGroupCode} {refDateFormat} {consolidated} \"{pathToJar.TrimEnd('\\')}\" \"{fileName}\"";
                            string jarParams = $" -cp {classPath} -jar {pathToJar}GenMacheteBind.jar {fileinpath} {filein} {fileoutpath} {xlsOutName} {execution.ToString()}";

                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "Start generare BNR excel file -> " + "\"" + pathToJRE + "\"" + jarParams
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                           // if (File.Exists(Path.Combine(pathToJar, xmFileName)))
                          //      File.Delete(Path.Combine(pathToJar, xmFileName));

                            Process proc = new Process();
                            proc.StartInfo = new ProcessStartInfo("\"" + pathToJRE + "\"", jarParams);
                            proc.StartInfo.UseShellExecute = false;
                            proc.StartInfo.RedirectStandardError = true;
                            proc.Start();

                            xmlError = proc.StandardError.ReadToEnd();

                            proc.WaitForExit();



                            if (xmlError.Length > 0)
                            {
                                parameters = new MSSqlParameter[1];
                                parameters[0] = new MSSqlParameter
                                {
                                    Name = "@message",
                                    Value = "ERROR -> " + xmlError
                                };
                                engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                                throw new Exception();
                            }
                            else {
                                parameters = new MSSqlParameter[2];
                                parameters[0] = new MSSqlParameter
                                {
                                    Name = "@exId",
                                    Value = execution
                                };
                                parameters[1] = new MSSqlParameter
                                {
                                    Name = "@fileName",
                                    Value = xlsOutName
                                };
                                engine.RunProcedureStatment("[rep_set_xlsx]", parameters);
                            }

                            #region XBRL_Generation

                            String xbrlArgsArr = "";
                            String reportingDate = "";


                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@exId",
                                Value = execution
                            };
                            DataTable xbrlCells = engine.RunProcedureQuery(@"dbo.xbrlArgs", parameters);
                            if (xbrlCells.Rows.Count > 0)
                            {
                                foreach (DataRow rand in xbrlCells.Rows)
                                {
                                    reportingDate = rand["ref_date"].ToString();
                                }
                            }
                            xbrlArgsArr = $" -Xms512m -jar {pathToJar}xbrl.jar {execution.ToString()} {reportingDate} {fileoutpath} {pathToJar} N 0 DUMMY_Module DUMMYLEI";

                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "Start XBRL Process with args:" + xbrlArgsArr
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                            Process xbrlProc = new Process();
                            xbrlProc.StartInfo = new ProcessStartInfo("\"" + pathToJRE + "\"", xbrlArgsArr);
                            xbrlProc.StartInfo.UseShellExecute = false;
                            xbrlProc.StartInfo.RedirectStandardError = true;
                            xbrlProc.Start();

                            xbrlProc.WaitForExit();

                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "END XBRL Process"
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);


                            #endregion


                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "After generate BNR excel file " + Path.Combine(pathToJar, xmFileName)
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);


                            parameters = new MSSqlParameter[3];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@executionid",
                                Value = execution
                            };
                            parameters[1] = new MSSqlParameter
                            {
                                Name = "@status",
                                Value = "2"
                            };
                            parameters[2] = new MSSqlParameter
                            {
                                Name = "@memo",
                                Value = ""
                            };

                            engine.RunProcedureStatment("[BINDEXCELSERVICE$SET_EXCELFILES_STATUS]", parameters);

                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "End generate excel BNR" 
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);
                        }

                        if (machetaExcelXml != null && machetaExcelXml != "")
                        {
                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "Start generate ABACUS excel " + Path.Combine(pathToJar, excelForXml)
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "Start generate ABACUS excel -> " + fileinpath + " " + machetaExcelXml + " " + pathToJar + " " + excelForXml + " " + execution
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                            GeneratorMacheteCorep g = new GeneratorMacheteCorep();

                            g.Generate(fileinpath, machetaExcelXml, pathToJar, excelForXml, execution.ToString(), _connection);

                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "After generate ABACUS excel  " + Path.Combine(fileoutpath, fileout)
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                            //System.IO.File.Copy(Path.Combine(fileoutpath, fileout), Path.Combine(fileinpath, excelForXml), true);

                            //parameters = new MSSqlParameter[1];
                            //parameters[0] = new MSSqlParameter
                            //{
                            //    Name = "@message",
                            //    Value = "After copy excel file: " + Path.Combine(fileinpath, excelForXml)
                            //};
                            //engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                            string jarParams = $"-jar \"{pathToJar}\\XlsToXmlConverter.jar\" {reportsGroupCode} {refDateFormat} {consolidated} \"{pathToJar.TrimEnd('\\')}\" \"{fileName}\"";

                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "Start generare xml file -> " + "\"" + pathToJRE + "\"" + " " + jarParams
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                            if (File.Exists(Path.Combine(pathToJar, xmFileName)))
                                File.Delete(Path.Combine(pathToJar, xmFileName));

                            Process proc = new Process();
                            proc.StartInfo = new ProcessStartInfo("\"" + pathToJRE + "\"", jarParams);
                            proc.StartInfo.UseShellExecute = false;
                            proc.StartInfo.RedirectStandardError = true;
                            proc.Start();

                            xmlError = proc.StandardError.ReadToEnd();

                            proc.WaitForExit();

                            if (xmlError.Length > 0)
                                throw new Exception();


                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "After generate xml file " + Path.Combine(pathToJar, xmFileName)
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@ParamType",
                                Value = 1
                            };

                            DataTable tbNextXmlId = engine.RunProcedureQuery(@"GET_NEXT_ID_ABACUS", parameters);

                            foreach (DataRow item in tbNextXmlId.Rows)
                                xmlUniqueId = item["id_xml"].ToString();

                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@ParamType",
                                Value = 2
                            };

                            DataTable tbNextExcelId = engine.RunProcedureQuery(@"GET_NEXT_ID_ABACUS", parameters);

                            foreach (DataRow item in tbNextExcelId.Rows)
                                excelUniqueId = item["id_excel"].ToString();

                            System.IO.File.Move(Path.Combine(pathToJar, xmFileName), Path.Combine(fileoutpath, xmlUniqueId + "_" + xmFileName));
                            System.IO.File.Move(Path.Combine(pathToJar, excelForXml), Path.Combine(fileoutpath, excelUniqueId + "_" + excelForXml));

                            //Apel procedura de update report_link si report_file cu numele xml-ului , sau modificat BINDEXCELSERVICE$SET_EXCELFILES_STATUS cu inca 2 parametrii

                            if (newExecution != "Y")
                            {
                                parameters = new MSSqlParameter[4];
                                parameters[0] = new MSSqlParameter
                                {
                                    Name = "@executionid",
                                    Value = execution
                                };
                                parameters[1] = new MSSqlParameter
                                {
                                    Name = "@report_file",
                                    Value = xmlUniqueId + "_" + xmFileName
                                };
                                parameters[2] = new MSSqlParameter
                                {
                                    Name = "@report_link",
                                    Value = xmlPathOnServer
                                };
                                parameters[3] = new MSSqlParameter
                                {
                                    Name = "@excel_file",
                                    Value = excelUniqueId + "_" + excelForXml
                                };

                                engine.RunProcedureStatment("[BINDEXCELSERVICE$SET_XML_OUTPUT]", parameters);

                            }                         
                            else
                            {
                                parameters = new MSSqlParameter[5];
                                parameters[0] = new MSSqlParameter
                                {
                                    Name = "@report_link",
                                    Value = xmlPathOnServer
                                };
                                parameters[1] = new MSSqlParameter
                                {
                                    Name = "@report_file",
                                    Value = xmlUniqueId + "_" + xmFileName
                                };
                                parameters[2] = new MSSqlParameter
                                {
                                    Name = "@excel_file",
                                    Value = excelUniqueId + "_" + excelForXml
                                };
                                parameters[3] = new MSSqlParameter
                                {
                                    Name = "@memo",
                                    Value = "Forked from execution: " + execution
                                };
                                parameters[4] = new MSSqlParameter
                                {
                                    Name = "@fileStatus",
                                    Value = 2
                                };

                                engine.RunProcedureStatment("[BINDEXCELSERVICE$SET_EXECUTION_FOR_XML]", parameters);
                            }
                        }


                    }
                    catch (Exception ex)
                    {

                        string errorMessage = "";

                        if (row != null || xmlError.Length > 0/* && row["FileId"] != null*/)
                        {

                            if (xmlError.Length > 3000)
                                errorMessage = xmlError.Substring(0, 3000);
                            else if (xmlError.Length > 0)
                                errorMessage = xmlError;
                            else if (ex.Message.Length > 3000)
                                errorMessage = ex.Message.Substring(0, 3000);
                            else
                                errorMessage = ex.Message;

                            parameters = new MSSqlParameter[1];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@message",
                                Value = "Eroare la generare: " + errorMessage
                            };
                            engine.RunProcedureStatment("[BINDEXCELSERVICE$log]", parameters);

                            parameters = new MSSqlParameter[3];
                            parameters[0] = new MSSqlParameter
                            {
                                Name = "@executionid",
                                Value = ((int)row["Execution_id"])
                            };
                            parameters[1] = new MSSqlParameter
                            {
                                Name = "@status",
                                Value = 0
                            };

                            parameters[2] = new MSSqlParameter
                            {
                                Name = "@memo",
                                Value = errorMessage
                            };

                            if(newExecution == "N")
                                engine.RunProcedureStatment("[BINDEXCELSERVICE$SET_EXCELFILES_STATUS]", parameters);
   
                        }
                        eventLog1.WriteEntry("Error processing file" + ex.Message);
                    }
                }
            }//files

            //eventLog1.WriteEntry("BindExcelManager ScanForFile end");

        }

        private void eventLog1_EntryWritten(object sender, EntryWrittenEventArgs e)
        {

        }
    }
}
