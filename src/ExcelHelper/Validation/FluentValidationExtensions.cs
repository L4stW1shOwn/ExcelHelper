using System;
using System.Globalization;

namespace ExcelHelper.Validation
{
    /// <summary>
    /// Validates that a field value is not null.
    /// </summary>
    public sealed class NotNullValidator : IExcelFieldValidator
    {
        /// <summary>
        /// Validates that the value is not null.
        /// </summary>
        public ValidationResult Validate(object? value, string? fieldName, int row, int column)
        {
            if (value == null)
                return ValidationResult.Failed($"Field '{fieldName}' at row {row}, column {column} cannot be null.");

            return ValidationResult.Success();
        }
    }

    /// <summary>
    /// Validates that a string field value is not null or empty.
    /// </summary>
    public sealed class NotEmptyValidator : IExcelFieldValidator
    {
        /// <summary>
        /// Validates that the string value is not null or empty.
        /// </summary>
        public ValidationResult Validate(object? value, string? fieldName, int row, int column)
        {
            if (value == null || (value is string s && string.IsNullOrWhiteSpace(s)))
                return ValidationResult.Failed($"Field '{fieldName}' at row {row}, column {column} cannot be empty.");

            return ValidationResult.Success();
        }
    }

    /// <summary>
    /// Validates that a numeric field value is greater than a specified minimum.
    /// </summary>
    public sealed class GreaterThanValidator : IExcelFieldValidator
    {
        private readonly double _minimum;

        /// <summary>
        /// Initializes a new instance of the <see cref="GreaterThanValidator"/> class.
        /// </summary>
        /// <param name="minimum">The minimum value (exclusive).</param>
        public GreaterThanValidator(double minimum)
        {
            _minimum = minimum;
        }

        /// <summary>
        /// Validates that the numeric value is greater than the minimum.
        /// </summary>
        public ValidationResult Validate(object? value, string? fieldName, int row, int column)
        {
            if (value == null)
                return ValidationResult.Success();

            double doubleValue;
            try
            {
                doubleValue = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return ValidationResult.Failed($"Field '{fieldName}' at row {row}, column {column} must be a number.");
            }

            if (doubleValue <= _minimum)
                return ValidationResult.Failed($"Field '{fieldName}' at row {row}, column {column} must be greater than {_minimum}.");

            return ValidationResult.Success();
        }
    }

    /// <summary>
    /// Validates that a numeric field value is within a specified range.
    /// </summary>
    public sealed class RangeValidator : IExcelFieldValidator
    {
        private readonly double _minimum;
        private readonly double _maximum;

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeValidator"/> class.
        /// </summary>
        /// <param name="minimum">The minimum value (inclusive).</param>
        /// <param name="maximum">The maximum value (inclusive).</param>
        public RangeValidator(double minimum, double maximum)
        {
            _minimum = minimum;
            _maximum = maximum;
        }

        /// <summary>
        /// Validates that the numeric value is within the range.
        /// </summary>
        public ValidationResult Validate(object? value, string? fieldName, int row, int column)
        {
            if (value == null)
                return ValidationResult.Success();

            double doubleValue;
            try
            {
                doubleValue = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return ValidationResult.Failed($"Field '{fieldName}' at row {row}, column {column} must be a number.");
            }

            if (doubleValue < _minimum || doubleValue > _maximum)
                return ValidationResult.Failed($"Field '{fieldName}' at row {row}, column {column} must be between {_minimum} and {_maximum}.");

            return ValidationResult.Success();
        }
    }
}
