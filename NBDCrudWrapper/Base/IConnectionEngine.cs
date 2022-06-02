using System.Data;
using System.Data.SqlClient;


namespace NBDCrudWrapper.Base
{

    public interface IConnectionEngine
    {
        /// <summary>
        /// The connectyion manager instance
        /// </summary>
        IConnectionManager Manager { get; }
        string StatementType { get; set; }

        /// <summary>
        /// Adds a new parameter
        /// </summary>
        /// <param name="name">Parameter name</param>
        /// <param name="value">Parameter value</param>
        /// <param name="direction">The direction to transmit data</param>
        /// <param name="type">The type of the parameter</param>
        SqlParameter AddParameter(string name, object value, ParameterDirection direction = ParameterDirection.Input, DbType type = DbType.String);

        /// <summary>
        /// Executes a procedure with no return value
        /// </summary>
        void Execute();

        /// <summary>
        /// Executes a procedure and returns the no of rows affected
        /// </summary>
        /// <returns>The first column on the first row</returns>
        object ExecuteScalar();

        /// <summary>
        /// Executes a procedure with a dataset return value
        /// </summary>
        /// <returns>The dataset</returns>
        DataSet GetDataSet();

        /// <summary>
        /// Executes a procedure with a datatable return value
        /// </summary>
        /// <returns>The first table of the set.</returns>
        DataTable GetDataTable();

        void SetConnection(string connection);
    }
}



