using System;
using OfficeOpenXml;

namespace ExcelHelper.Internal;

/// <summary>
///     Provides helper methods for working with shared strings in Excel worksheets.
/// </summary>
public static class SharedStringTableHelper
{
    /// <summary>
    ///     Gets the text value of a cell, resolving shared strings if necessary.
    /// </summary>
    /// <param name="cell">The Excel range cell.</param>
    /// <returns>The cell text.</returns>
    public static string GetCellText(ExcelRangeBase cell)
    {
        if (cell == null)
        {
            throw new ArgumentNullException(nameof(cell));
        }

        return cell.Text ?? string.Empty;
    }

    /// <summary>
    ///     Gets the raw value of a cell.
    /// </summary>
    /// <param name="cell">The Excel range cell.</param>
    /// <returns>The cell value.</returns>
    public static object? GetCellValue(ExcelRangeBase cell)
    {
        return cell == null ? throw new ArgumentNullException(nameof(cell)) : cell.Value;
    }
}