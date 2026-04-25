using System;

namespace ExcelHelper.Mapping;

/// <summary>
///     Specifies the column name for a property.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelNameAttribute : Attribute
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelNameAttribute" /> class.
    /// </summary>
    /// <param name="name">The column name.</param>
    public ExcelNameAttribute(string name)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
    }

    /// <summary>
    ///     Gets the column name.
    /// </summary>
    public string Name { get; }
}