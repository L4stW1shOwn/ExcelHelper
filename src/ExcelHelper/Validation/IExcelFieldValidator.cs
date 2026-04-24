using System;

namespace ExcelHelper.Validation
{
    /// <summary>
    /// Validates a field value during reading.
    /// </summary>
    public interface IExcelFieldValidator
    {
        /// <summary>
        /// Validates the specified field value.
        /// </summary>
        /// <param name="value">The field value.</param>
        /// <param name="fieldName">The name of the field.</param>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <returns>A <see cref="ValidationResult"/> indicating whether validation succeeded.</returns>
        ValidationResult Validate(object? value, string? fieldName, int row, int column);
    }
}
