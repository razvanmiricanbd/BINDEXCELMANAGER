using System;
using System.Data;
using System.Data.SqlClient;
using NBDCrudWrapper.Base;

namespace NBDCrudWrapper.Base.MSql
{
    /// <summary>
    /// Manages the repository connection
    /// </summary>
    public class ConnectionManager : IConnectionManager
    {
        #region Variables

        private string _connectionName = "MSSqlConnectionString";
        private string _connectionString = string.Empty;

        #endregion Variables

        #region Properties

        /// <summary>
        /// Gets the formated connection string for the datasource
        /// </summary>
        public string ConnectionString
        {
            get
            {
                return _connectionString;
            }
        }

        #endregion Properties

        #region Public methods

       

        /// <summary>
        /// Sets the name of the connection string to be read from the confog file
        /// </summary>
        public void SetConnectionName(string name)
        {
            if (name == null) throw new ArgumentNullException("name");
            _connectionName = name;
        }

        /// <summary>
        /// Get the repository connection instance
        /// </summary>
        /// <returns></returns>
        public IDbConnection GetRepositoryConnection()
        {
            return new SqlConnection(ConnectionString);
        }
        public void SetConnectionString(string conn)
        {
            if (conn == null) throw new ArgumentNullException("conn");
            _connectionString = conn;
        }
        #endregion Public methods
    }
}
