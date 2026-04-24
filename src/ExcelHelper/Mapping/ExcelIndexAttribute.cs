using System;

namespace ExcelHelper.Mapping
{
    /// <summary>
    /// Specifies the 0-based column index for a property.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelIndexAttribute : Attribute
    {
        /// <summary>
        /// Gets the 0-based column index.
        /// </summary>
        public int Index { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelIndexAttribute"/> class.
        /// </summary>
        /// <param name="index">The 0-based column index.</param>
        public ExcelIndexAttribute(int index)
        {
            Index = index;
        }
    }
}
