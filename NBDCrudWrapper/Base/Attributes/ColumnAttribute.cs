namespace NBDCrudWrapper.Base.Attributes
{
    /// <summary>
    /// Marks a property as boundatable to a datacloumn
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public class ColumnAttribute : System.Attribute
    {
        #region Variables

        private readonly string _name;

        #endregion Variables

        #region Properties

        /// <summary>
        /// The name of the column to bind to
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// arks a property as boundatable to a datacloumn
        /// </summary>
        /// <param name="name">The name of the column to bind to</param>
        public ColumnAttribute(string name)
        {
            _name = name;
        }

        #endregion Constructors
    }
}
