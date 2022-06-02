using System;

namespace NBDCrudWrapper.Base
{
    /// <summary>
    /// Entity used to filter the results
    /// </summary>
    public class Filter
    {
        /// <summary>
        /// The type of the reflected entity
        /// </summary>
        public Type Type { get; private set; }

        /// <summary>
        /// The name of the filter
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The value of the filter
        /// </summary>
        public object Value { get; set; }

        public Filter(Type type)
        {
            if (type == null) throw new ArgumentNullException("type");
            if (!type.IsSubclassOf(typeof(BaseEntity))) throw new Exception(string.Format("{0} must implement class NBD.Base.Entities.Base.BaseEntity", type.Name));

            Type = type;
        }
    }
}
