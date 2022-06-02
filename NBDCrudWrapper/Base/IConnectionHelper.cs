using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace NBDCrudWrapper.Base
{

            /// <summary>
        /// Provides methods that facilitates the connection
        /// </summary>
        public interface IConnectionHelper
        {
            /// <summary>
            /// Gets the procedure name based on the prefix attribute and a given suffix.
            /// If a procedure attribute is not found, then the type is used as sufix
            /// </summary>
            /// <param name="name">The sufix</param>
            /// <returns></returns>
            string GetProcedureName<TEntity>(string name);

            /// <summary>
            /// Adds entity properties as filters to the connection engine
            /// </summary>
            /// <typeparam name="TEntity">The entity type</typeparam>
            /// <param name="connection">The conection instance</param>
            /// <param name="entity">The entity instance</param>
            /// <returns>The identity property info</returns>
            void AddParameters<TEntity>(IConnectionEngine connection, TEntity entity) where TEntity : IBaseEntity, new();

            /// <summary>
            /// Adds entity properties as filters to the connection engine
            /// </summary>
            /// <typeparam name="TEntity">The entity type</typeparam>
            /// <param name="connection">The conection instance</param>
            /// <param name="entity">The entity instance</param>
            /// <param name="ignore">Indicates if the entity id will be ignored</param>
            /// <returns>The identity property info</returns>
            PropertyInfo AddParameters<TEntity>(IConnectionEngine connection, TEntity entity, bool ignore) where TEntity : IBaseEntity, new();

            /// <summary>
            /// Gets the data from the datarow and puts it to a new specific entity instance
            /// </summary>
            /// <typeparam name="TEntity">The entity type</typeparam>
            /// <param name="row">The datarow to be bound</param>
            /// <returns>The specific entity instance</returns>
            TEntity Bind<TEntity>(DataRow row) where TEntity : IBaseEntity, new();

            /// <summary>
            /// Gets the data from the table and puts it to a new specific entity array
            /// </summary>
            /// <typeparam name="TEntity">The entity type</typeparam>
            /// <param name="table">The table instance</param>
            /// <returns>An array of specific entities</returns>
            IEnumerable<TEntity> Bind<TEntity>(DataTable table) where TEntity : IBaseEntity, new();
        }
    }

