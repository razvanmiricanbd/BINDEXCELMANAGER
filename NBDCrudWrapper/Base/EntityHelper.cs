using System;
using System.Linq;
using System.Reflection;


namespace NBDCrudWrapper.Base
{
    /// <summary>
    /// Proveides a series of methods that facilitate the entity management
    /// </summary>
    public static class EntityHelper
    {
        /// <summary>
        /// Gets the property from the entity that has a specific attribute
        /// </summary>
        /// <typeparam name="TEntity">The entity type</typeparam>
        /// <typeparam name="TAttribute">The attribute type</typeparam>
        /// <returns>The property info</returns>
        public static PropertyInfo GetProperty<TEntity, TAttribute>()
            where TEntity : IBaseEntity, new()
            where TAttribute : Attribute
        {
            var properties = new TEntity().GetType().GetProperties();
            return (from property in properties let attribute = GetAttribute<TAttribute>(property) where attribute != null select property).FirstOrDefault();
        }

        /// <summary>
        /// Gets the property from the entity that has a specific attribute
        /// </summary>
        /// <typeparam name="TEntity">The entity type</typeparam>
        /// <typeparam name="TAttribute">The attribute type</typeparam>
        /// <param name="entity">The entity instance</param>
        /// <returns>The property info</returns>
        public static PropertyInfo GetProperty<TEntity, TAttribute>(TEntity entity)
            where TEntity : IBaseEntity, new()
            where TAttribute : Attribute
        {
            var properties = entity.GetType().GetProperties();
            return (from property in properties let attribute = GetAttribute<TAttribute>(property) where attribute != null select property).FirstOrDefault();
        }

        /// <summary>
        /// Verifies if the property info has a specific attibute
        /// </summary>
        /// <typeparam name="T">The type of the attribute</typeparam>
        /// <param name="info">The property info instance</param>
        /// <returns>True if the attribute is found</returns>
        public static bool HasAttribute<T>(PropertyInfo info) where T : Attribute
        {
            var attributes = info.GetCustomAttributes(typeof(T), true);
            return (attributes.Length >= 1);
        }

        /// <summary>
        /// Gets a specific attribute from the property info
        /// </summary>
        /// <typeparam name="T">The type of the attribute</typeparam>
        /// <param name="info">The property info instance</param>
        /// <returns>The attribute instance</returns>
        public static T GetAttribute<T>(PropertyInfo info) where T : Attribute
        {
            var attributes = info.GetCustomAttributes(typeof(T), true);
            return attributes.FirstOrDefault() as T;
        }
    }
}
