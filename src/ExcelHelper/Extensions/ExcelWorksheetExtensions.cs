using OfficeOpenXml;

namespace ExcelHelper.Extensions;

/// <summary>
///     Provides extension methods for <see cref="ExcelWorksheet" />.
/// </summary>
public static class ExcelWorksheetExtensions
{
    /// <summary>
    ///     Determines whether the specified row is blank between the given column range.
    /// </summary>
    /// <param name="worksheet">The worksheet.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="startColumn">The 1-based start column index.</param>
    /// <param name="endColumn">The 1-based end column index.</param>
    /// <returns><c>true</c> if the row is blank; otherwise, <c>false</c>.</returns>
    public static bool IsBlankRow(this ExcelWorksheet? worksheet, int row, int startColumn, int endColumn)
    {
        if (worksheet == null)
        {
            return true;
        }

        for (var col = startColumn; col <= endColumn; col++)
        {
            var value = worksheet.Cells[row, col].Value;
            if (value != null)
            {
                return false;
            }
        }

        return true;
    }
}