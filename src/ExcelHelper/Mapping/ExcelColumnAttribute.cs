using System;

namespace ExcelHelper.Mapping
{
    /// <summary>
    /// Specifies that a property should be mapped to an Excel column.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets the name of the column.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Gets or sets the 0-based index of the column.
        /// </summary>
        public int Index { get; set; } = -1;

        /// <summary>
        /// Gets or sets the default value for the column when the cell is empty.
        /// </summary>
        public object? Default { get; set; }

        /// <summary>
        /// Gets or sets the type of the converter to use for this column.
        /// </summary>
        public Type? Converter { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColumnAttribute"/> class.
        /// </summary>
        public ExcelColumnAttribute()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColumnAttribute"/> class with a column name.
        /// </summary>
        /// <param name="name">The column name.</param>
        public ExcelColumnAttribute(string name)
        {
            Name = name;
        }
    }
}
