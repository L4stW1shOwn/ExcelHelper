namespace ExcelHelper.Validation;

/// <summary>
///     Validates a record during reading.
/// </summary>
/// <typeparam name="T">The record type.</typeparam>
public interface IExcelRecordValidator<in T>
{
    /// <summary>
    ///     Validates the specified record.
    /// </summary>
    /// <param name="record">The record to validate.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <returns>A <see cref="ValidationResult" /> indicating whether validation succeeded.</returns>
    ValidationResult Validate(T record, int row);
}