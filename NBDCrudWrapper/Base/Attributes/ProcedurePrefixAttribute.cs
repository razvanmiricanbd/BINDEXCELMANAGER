namespace NBDCrudWrapper.Base.Attributes
{
    /// <summary>
    /// Provides a prefix for the executed procedure name
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class)]
    public class ProcedurePrefixAttribute : System.Attribute
    {
        #region Variables

        private readonly string _prefix;

        #endregion Variables

        #region Properties

        /// <summary>
        /// The prefix of the procedure
        /// </summary>
        public string Prefix
        {
            get
            {
                return _prefix;
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Provides a prefix for the executed procedure name
        /// </summary>
        /// <param name="prefix">The prefix to be added</param>
        public ProcedurePrefixAttribute(string prefix)
        {
            _prefix = prefix;
        }

        #endregion Constructors
    }
}
