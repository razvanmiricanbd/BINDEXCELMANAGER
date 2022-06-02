using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using NBDCrudWrapper.Base;
using NBDCrudWrapper.Base.Attributes;

namespace NBDCrudWrapper.Base.MSql
{
    /// <summary>
    /// Provides methods that facilitates the connection
    /// </summary>
    public class ConnectionHelper : IConnectionHelper
    {
        /// <summary>
        /// Gets the procedure name based on the prefix attribute and a given suffix.
        /// If a procedure attribute is not found, then the type is used as sufix
        /// </summary>
        /// <param name="name">The sufix</param>
        /// <returns></returns>
        public string GetProcedureName<TEntity>(string name)
        {
            var procedure = string.Format("[usp.{0}.{1}]", typeof(TEntity).Name, name);
            var attributes = typeof(TEntity).GetCustomAttributes(typeof(ProcedurePrefixAttribute), true);
            if (attributes.Length == 1)
            {
                var attribute = attributes[0] as ProcedurePrefixAttribute;
                if (attribute != null)
                {
                    procedure = string.Format("[{0}.{1}]", attribute.Prefix, name);
                }
            }

            return procedure;
        }

        /// <summary>
        /// Adds entity properties as filters to the connection engine
        /// </summary>
        /// <typeparam name="TEntity">The entity type</typeparam>
        /// <param name="connection">The conection instance</param>
        /// <param name="entity">The entity instance</param>
        /// <returns>The identity property info</returns>
        public void AddParameters<TEntity>(IConnectionEngine connection, TEntity entity) where TEntity : IBaseEntity, new()
        {
            var properties = entity.GetType().GetProperties();
            foreach (var property in properties)
            {
                //if the property is ignored, skip it
                if (EntityHelper.HasAttribute<IgnoredAttribute>(property))
                    continue;

                //if the property is used only on read, skip it
                if (EntityHelper.HasAttribute<ReadonlyAttribute>(property))
                    continue;

                //the name of the parameter
                var parameterName = property.Name;

                //if there is a bind attibute, get the name of the column from it
                var column = EntityHelper.GetAttribute<ColumnAttribute>(property);
                if (column != null) parameterName = column.Name;

                //get the value of the parameter
                var value = property.GetValue(entity, null) ?? DBNull.Value;

                //add the parameter to the connection
                connection.AddParameter(parameterName, value);
            }
        }

        /// <summary>
        /// Adds entity properties as filters to the connection engine
        /// </summary>
        /// <typeparam name="TEntity">The entity type</typeparam>
        /// <param name="connection">The conection instance</param>
        /// <param name="entity">The entity instance</param>
        /// <param name="ignore">Indicates if the entity id will be ignored</param>
        /// <returns>The identity property info</returns>
        public PropertyInfo AddParameters<TEntity>(IConnectionEngine connection, TEntity entity, bool ignore) where TEntity : IBaseEntity, new()
        {
            var identity = null as PropertyInfo;
            var properties = entity.GetType().GetProperties();
            foreach (var property in properties)
            {
                //if the property is ignored, skip it
                if (EntityHelper.HasAttribute<IgnoredAttribute>(property))
                    continue;

                //if the property is used only on read, skip it
                if (EntityHelper.HasAttribute<ReadonlyAttribute>(property))
                    continue;

                //get the identity attribute
                var identityattribute = EntityHelper.GetAttribute<IdentityAttribute>(property);
                if (identityattribute != null)
                    identity = property;

                //if tha attribute is ignored, skip it
                if (identityattribute != null && ignore)
                    continue;

                //the name of the parameter
                var parameterName = property.Name;

                //if there is a bind attibute, get the name of the column from it
                var bindAttribute = EntityHelper.GetAttribute<ColumnAttribute>(property);
                if (bindAttribute != null)
                    parameterName = bindAttribute.Name;

                //get the value of the parameter
                var value = property.GetValue(entity, null) ?? DBNull.Value;

                //add the parameter to the connection
                connection.AddParameter(parameterName, value);
            }

            return identity;
        }

        /// <summary>
        /// Gets the data from the datarow and puts it to a new specific entity instance
        /// </summary>
        /// <typeparam name="TEntity">The entity type</typeparam>
        /// <param name="row">The datarow to be bound</param>
        /// <returns>The specific entity instance</returns>
        public TEntity Bind<TEntity>(DataRow row) where TEntity : IBaseEntity, new()
        {
            var entity = new TEntity();
            var properties = typeof(TEntity).GetProperties();
            foreach (var property in properties)
            {
                if (!property.CanWrite)
                    continue;

                //if the property is ignored, skip it
                if (EntityHelper.HasAttribute<IgnoredAttribute>(property))
                    continue;

                //the name of the parameter
                var columnName = property.Name;

                //if there is a bind attibute, get the name of the column from it
                var bindAttribute = EntityHelper.GetAttribute<ColumnAttribute>(property);
                if (bindAttribute != null)
                    columnName = bindAttribute.Name;

                //if the property is mandatory and it is not in the table, we throw an error
                if (EntityHelper.HasAttribute<RequiredAttribute>(property) && !row.Table.Columns.Contains(columnName))
                    throw new Exception(string.Format("DataAccess BindFromRow: mandatory property {0} for {1} not found on DataTable columns", columnName, property.DeclaringType));

                //if the column is not in the row, skip it
                if (!row.Table.Columns.Contains(columnName))
                    continue;

                //get the value form the row
                var value = row[columnName];

                //set the value to the object
                if (value != null && value != DBNull.Value)
                    property.SetValue(entity, value, null);
            }

            return entity;
        }

        /// <summary>
        /// Gets the data from the table and puts it to a new specific entity array
        /// </summary>
        /// <typeparam name="TEntity">The entity type</typeparam>
        /// <param name="table">The table instance</param>
        /// <returns>An array of specific entities</returns>
        public IEnumerable<TEntity> Bind<TEntity>(DataTable table) where TEntity : IBaseEntity, new()
        {
            var entities = new List<TEntity>();

            if (table != null)
            {
                entities.AddRange(from DataRow row in table.Rows select Bind<TEntity>(row));
            }

            return entities;
        }
    }
}
