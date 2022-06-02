namespace NBDCrudWrapper.Base.Attributes
{
    /// <summary>
    /// Marks a property to be taken in consideration only when reading from the datasource
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public class ReadonlyAttribute : System.Attribute
    {
    }
}
