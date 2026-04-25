using System;

namespace ExcelHelper.Mapping;

/// <summary>
///     Specifies the default value for a property when the cell is empty.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelDefaultAttribute : Attribute
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelDefaultAttribute" /> class.
    /// </summary>
    /// <param name="value">The default value.</param>
    public ExcelDefaultAttribute(object? value)
    {
        Value = value;
    }

    /// <summary>
    ///     Gets the default value.
    /// </summary>
    public object? Value { get; }
}