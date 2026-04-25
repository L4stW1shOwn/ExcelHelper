using System;

namespace ExcelHelper.Mapping;

/// <summary>
///     Specifies the 0-based column index for a property.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelIndexAttribute : Attribute
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelIndexAttribute" /> class.
    /// </summary>
    /// <param name="index">The 0-based column index.</param>
    public ExcelIndexAttribute(int index)
    {
        Index = index;
    }

    /// <summary>
    ///     Gets the 0-based column index.
    /// </summary>
    public int Index { get; }
}