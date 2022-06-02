using System.Data;
namespace NBDCrudWrapper.Base
{
    public interface IConnectionManager
    {
        /// <summary>
        /// Gets the formated connection string for the datasource
        /// </summary>
        string ConnectionString { get; }



        /// <summary>
        /// Sets the name of the connection string to be read from the confog file
        /// </summary>
        void SetConnectionName(string name);

        void SetConnectionString(string conn);

        /// <summary>
        /// Get the repository connection instance
        /// </summary>
        /// <returns></returns>
        IDbConnection GetRepositoryConnection();
    }
}
