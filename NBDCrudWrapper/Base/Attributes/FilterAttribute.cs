using System;

namespace NBD.Base.Entities.Attributes.Property
{
    /// <summary>
    /// Marks a property as boundatable to a datacloumn
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class FilterAttribute : Attribute
    {
        #region Variables

        private readonly string _name;
        private readonly Type _type;

        #endregion Variables

        #region Properties

        /// <summary>
        /// The name of the column to bind to
        /// </summary>
        public string Name
        {
            get { return _name; }
        }

        public Type Type
        {
            get { return _type; }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// arks a property as boundatable to a datacloumn
        /// </summary>
        /// <param name="name">The name of the column to bind to</param>
        public FilterAttribute(string name)
        {
            _name = name;
        }

        public FilterAttribute(string name, Type type)
        {
            _name = name;
            _type = type;
        }

        #endregion Constructors
    }
}
