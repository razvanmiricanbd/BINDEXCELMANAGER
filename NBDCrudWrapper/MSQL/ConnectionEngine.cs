using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using NBDCrudWrapper.Base;

namespace NBDCrudWrapper.Base.MSql
{
    /// <summary>
    /// Contains the base actions executed agains the repository
    /// </summary>
    public class ConnectionEngine : IConnectionEngine
    {
        #region Variables

        /// <summary>
        /// The name of the procedure
        /// </summary>
        private readonly string _procedure;

        /// <summary>
        /// The list of Sql Parameters
        /// </summary>
        private List<SqlParameter> _parameters;

        #endregion Variables

        #region Properties

        /// <summary>
        /// The connectyion manager instance
        /// </summary>
        public IConnectionManager Manager { get; private set; }
        public string StatementType { get; set; }
        #endregion

        #region Constructors

        ///// <summary>
        ///// Static constuctor
        ///// </summary>
        //static ConnectionEngine()
        //{
        //    Connections = new ConcurrentDictionary<string, string>();
        //}

        /// <summary>
        /// Creates an instance of the connection
        /// </summary>
        private ConnectionEngine()
        {
            Manager = new ConnectionManager();
            StatementType = "StoredProcedure";
        }

        /// <summary>
        /// Creates an instance of the connection
        /// </summary>
        /// <param name="procedure">The name of the procedure to be executed</param>
        public ConnectionEngine(string procedure)
            : this()
        {
            _procedure = procedure;
            StatementType = "StoredProcedure";
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Adds a new parameter
        /// </summary>
        /// <param name="name">Parameter name</param>
        /// <param name="value">Parameter value</param>
        /// <param name="direction">The direction to transmit data</param>
        /// <param name="type">The type of the parameter</param>
        public SqlParameter AddParameter(string name, object value, ParameterDirection direction = ParameterDirection.Input, DbType type = DbType.String)
        {
            if (_parameters == null)
                _parameters = new List<SqlParameter>();

            //create the parameter instance
            var parameter = new SqlParameter
            {
                ParameterName = name,
                Value = value ?? DBNull.Value,
                Direction = direction,
                DbType = type
            };

            //add the parameter to the list
            _parameters.Add(parameter);

            //return the instance
            return parameter;
        }

        /// <summary>
        /// Executes a procedure with no return value
        /// </summary>
        public void Execute()
        {
            using (var command = GetCommand())
            {
                try
                {
                    //add the parameters to the command
                    AddParameters(command);

                    //execute the command
                    command.ExecuteNonQuery();
                }
                finally
                {
                    command.Connection.Close();
                    command.Connection.Dispose();
                }
            }
        }

        /// <summary>
        /// Executes a procedure and returns the no of rows affected
        /// </summary>
        /// <returns>The first column on the first row</returns>
        public object ExecuteScalar()
        {
            object value;

            using (var command = GetCommand())
            {
                try
                {
                    //add the parameters to the command
                    AddParameters(command);

                    //execute the command
                    value = command.ExecuteScalar();
                }
                finally
                {
                    command.Connection.Close();
                    command.Connection.Dispose();
                }
            }

            return value;
        }

        /// <summary>
        /// Executes a procedure with a dataset return value
        /// </summary>
        /// <returns>The dataset</returns>
        public DataSet GetDataSet()
        {
            var dataset = new DataSet();

            using (var command = GetCommand())
            {
                try
                {
                    //add the parameters to the command
                    AddParameters(command);

                    //execute the command
                    command.ExecuteNonQuery();

                    //creates an addapter for the command
                    var adapter = new SqlDataAdapter(command);

                    //populates the dataset with the adapter data
                    adapter.Fill(dataset);
                }
                finally
                {
                    command.Connection.Close();
                    command.Connection.Dispose();
                }
            }

            return dataset;
        }

        /// <summary>
        /// Executes a procedure with a datatable return value
        /// </summary>
        /// <returns>The first table of the set.</returns>
        public DataTable GetDataTable()
        {
            var table = new DataTable();

            using (var command = GetCommand())
            {
                try
                {
                    //add the parameters to the command
                    AddParameters(command);

                    //execute the command
                   // command.ExecuteNonQuery();

                    //creates an addapter for the command
                    var dataAdapter = new SqlDataAdapter(command);

                    //populates the datatable with the adapter data
                    dataAdapter.Fill(table);
                }
                finally
                {
                    command.Connection.Close();
                    command.Connection.Dispose();
                }
            }

            return table;
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Creates a new instance of the SqlCommand to connect to the database
        /// </summary>
        /// <returns>The SqlCommand instance</returns>
        private SqlCommand GetCommand()
        {
            //check if the procedure name was provided
            if (_procedure == null)
                throw new Exception("ConnectionEngine => GetCommand: The procedure was not provided");

            //create the command instance
            var command = new SqlCommand(_procedure)
            {

                CommandType = (StatementType == "StoredProcedure" ?CommandType.StoredProcedure:CommandType.Text),
                Connection = Manager.GetRepositoryConnection() as SqlConnection
            };

            //open the connection
            if (command.Connection.State != ConnectionState.Open)
            {
                command.Connection.Open();
            }

            return command;
        }

        /// <summary>
        /// Adds parameters to the sql command
        /// </summary>
        /// <param name="command">The Sql Command instance</param>
        private void AddParameters(SqlCommand command)
        {
            //check if the parameters list was initialized
            if (_parameters == null)
                return;

            //add the parameters to the command
            foreach (var parameter in _parameters)
                command.Parameters.Add(parameter);
        }

        #endregion

        public void SetConnection(string connection)
        {
            Manager.SetConnectionString(connection);
        }
    }
}

