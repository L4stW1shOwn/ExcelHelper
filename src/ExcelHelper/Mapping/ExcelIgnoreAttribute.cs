using System;

namespace ExcelHelper.Mapping
{
    /// <summary>
    /// Specifies that a property should be ignored during mapping.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelIgnoreAttribute : Attribute
    {
    }
}
