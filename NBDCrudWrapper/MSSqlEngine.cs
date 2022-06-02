
using NBDCrudWrapper.Base.MSql;
using System.Data;
using System.Data.SqlClient;
namespace NBDCrudWrapper
{
    public class MSSqlEngine
    {

        private ConnectionEngine _connectionEngine;
        private string _connectionString;

        public MSSqlEngine(string ConnectionString)
        {
            this._connectionString = ConnectionString;

        }
        public DataTable GetExecutableObjects(string objectType, string catalog)
        {
            _connectionEngine = new ConnectionEngine(@"SELECT ROUTINE_NAME, ROUTINE_TYPE,ROUTINE_SCHEMA
                            FROM INFORMATION_SCHEMA.ROUTINES
                            where ROUTINE_CATALOG = @catalog
                            and  (ROUTINE_TYPE =@routineType 
                                or @getAll ='Y' )");

            _connectionEngine.SetConnection(_connectionString);
            _connectionEngine.StatementType = "Text";

            _connectionEngine.AddParameter("@catalog", catalog);
            _connectionEngine.AddParameter("@routineType", objectType);
            if (objectType == "ALL")
                _connectionEngine.AddParameter("@getAll", "Y");
            else _connectionEngine.AddParameter("@getAll", "N");

            DataTable procedures = _connectionEngine.GetDataTable();

            return procedures;
        }
        public DataTable GetParametersForExecutable(string objectName)
        {
            _connectionEngine = new ConnectionEngine(@" SELECT SCHEMA_NAME(SCHEMA_ID)AS[Schema],
                                        SO.name AS[ObjectName],
                                        P.parameter_id AS[ParameterID],
                                        P.name AS[ParameterName],
                                        TYPE_NAME(P.user_type_id) AS[ParameterDataType],
                                        P.max_length AS[ParameterMaxBytes],
                                        P.is_output AS[IsOutPutParameter]
                                        FROM sys.objects AS SO
                                        INNER JOIN sys.parameters AS P
                                        ON SO.OBJECT_ID = P.OBJECT_ID
                                        WHERE SO.OBJECT_ID IN(SELECT OBJECT_ID
                                        FROM sys.objects
                                        WHERE TYPE IN('P', 'FN'))
                                        and SO.name = @object_name
                                        ORDER BY[Schema], SO.name, P.parameter_id");
            _connectionEngine.SetConnection(_connectionString);
            _connectionEngine.StatementType = "Text";

            _connectionEngine.AddParameter("@object_name", objectName);
            DataTable parameters = _connectionEngine.GetDataTable();

            return parameters;

        }

        public DataTable GetTableObjects(string objectType, string catalog)
        {
            _connectionEngine = new ConnectionEngine(@"SELECT TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,TABLE_TYPE
                                                     FROM INFORMATION_SCHEMA.TABLES 
                                                     where TABLE_CATALOG =@catalog
                                                         and  (TABLE_TYPE =@tableType 
                                                             or @getAll ='Y' )");

            _connectionEngine.SetConnection(_connectionString);
            _connectionEngine.StatementType = "Text";

            _connectionEngine.AddParameter("@catalog", catalog);
            _connectionEngine.AddParameter("@tableType", objectType);
            if (objectType == "ALL")
                _connectionEngine.AddParameter("@getAll", "Y");
            else _connectionEngine.AddParameter("@getAll", "N");

            DataTable tables = _connectionEngine.GetDataTable();

            return tables;
        }

        public DataTable GetColumnsForTables(string objectName)
        {
            _connectionEngine = new ConnectionEngine(@" SELECT column_name, ORDINAL_POSITION ,
                                case
                                when DATA_TYPE in ('varchar','nvarchar') then DATA_TYPE +'(' + cast(CHARACTER_MAXIMUM_LENGTH as varchar) +')'
	                                                       when DATA_TYPE in ('float','real','numeric','decimal')  then DATA_TYPE +
						                                                                        '(' + cast(isnull(NUMERIC_PRECISION,0) as varchar) +','+
																                                       cast(isnull(NUMERIC_SCALE,0) as varchar)+')'
	                                   else DATA_TYPE
	                                   end DATA_TYPE,IS_NULLABLE,COLUMN_DEFAULT,CHARACTER_SET_NAME,COLLATION_NAME
	                                   FROM information_schema.columns 
                                where table_name = @object_name
                                        order by ORDINAL_POSITION");
            _connectionEngine.SetConnection(_connectionString);
            _connectionEngine.StatementType = "Text";

            _connectionEngine.AddParameter("@object_name", objectName);
            DataTable columns = _connectionEngine.GetDataTable();

            return columns;

        }


        public DataTable RunQuerry(string statement, MSSqlParameter[] parameters)
        {
            _connectionEngine = new ConnectionEngine(statement);
            _connectionEngine.SetConnection(_connectionString);
            _connectionEngine.StatementType = "Text";
            if (parameters !=null)
            foreach (MSSqlParameter paramenter in parameters)
            {
                _connectionEngine.AddParameter(paramenter.Name, paramenter.Value);
                
            }
            DataTable data = _connectionEngine.GetDataTable();
            return data;
        }

        public DataTable RunProcedureQuery(string procedure, MSSqlParameter[] parameters)
        {
            _connectionEngine = new ConnectionEngine(procedure);
            _connectionEngine.SetConnection(_connectionString);
            if (parameters != null)
                foreach (MSSqlParameter paramenter in parameters)
                {
                    _connectionEngine.AddParameter(paramenter.Name, paramenter.Value);

                }
            DataTable data = _connectionEngine.GetDataTable();
            return data;
        }

        public void RunProcedureStatment(string procedure, MSSqlParameter[] parameters)
        {
            _connectionEngine = new ConnectionEngine(procedure);
            _connectionEngine.SetConnection(_connectionString);
            if (parameters != null)
                foreach (MSSqlParameter paramenter in parameters)
                {
                    _connectionEngine.AddParameter(paramenter.Name, paramenter.Value);

                }
            _connectionEngine.Execute();

        }
    }
}
