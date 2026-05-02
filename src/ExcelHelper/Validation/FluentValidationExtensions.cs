using System;
using System.Globalization;

namespace ExcelHelper.Validation;

/// <summary>
///     Validates that a field value is not null.
/// </summary>
public sealed class NotNullValidator : IExcelFieldValidator
{
    /// <summary>
    ///     Validates that the value is not null.
    /// </summary>
    public ValidationResult Validate(ValidateArgs args)
    {
        return args.Value == null
            ? ValidationResult.Failed($"Field '{args.FieldName}' at row {args.Row}, column {args.Column} cannot be null.")
            : ValidationResult.Success();
    }
}

/// <summary>
///     Validates that a string field value is not null or empty.
/// </summary>
public sealed class NotEmptyValidator : IExcelFieldValidator
{
    /// <summary>
    ///     Validates that the string value is not null or empty.
    /// </summary>
    public ValidationResult Validate(ValidateArgs args)
    {
        if (args.Value == null || (args.Value is string s && string.IsNullOrWhiteSpace(s)))
        {
            return ValidationResult.Failed($"Field '{args.FieldName}' at row {args.Row}, column {args.Column} cannot be empty.");
        }

        return ValidationResult.Success();
    }
}

/// <summary>
///     Validates that a numeric field value is greater than a specified minimum.
/// </summary>
public sealed class GreaterThanValidator : IExcelFieldValidator
{
    private readonly double _minimum;

    /// <summary>
    ///     Initializes a new instance of the <see cref="GreaterThanValidator" /> class.
    /// </summary>
    /// <param name="minimum">The minimum value (exclusive).</param>
    public GreaterThanValidator(double minimum)
    {
        _minimum = minimum;
    }

    /// <summary>
    ///     Validates that the numeric value is greater than the minimum.
    /// </summary>
    public ValidationResult Validate(ValidateArgs args)
    {
        if (args.Value == null)
        {
            return ValidationResult.Success();
        }

        double doubleValue;
        try
        {
            doubleValue = Convert.ToDouble(args.Value, CultureInfo.InvariantCulture);
        }
        catch
        {
            return ValidationResult.Failed($"Field '{args.FieldName}' at row {args.Row}, column {args.Column} must be a number.");
        }

        if (doubleValue <= _minimum)
        {
            return ValidationResult.Failed(
                $"Field '{args.FieldName}' at row {args.Row}, column {args.Column} must be greater than {_minimum}.");
        }

        return ValidationResult.Success();
    }
}

/// <summary>
///     Validates that a numeric field value is within a specified range.
/// </summary>
public sealed class RangeValidator : IExcelFieldValidator
{
    private readonly double _maximum;
    private readonly double _minimum;

    /// <summary>
    ///     Initializes a new instance of the <see cref="RangeValidator" /> class.
    /// </summary>
    /// <param name="minimum">The minimum value (inclusive).</param>
    /// <param name="maximum">The maximum value (inclusive).</param>
    public RangeValidator(double minimum, double maximum)
    {
        _minimum = minimum;
        _maximum = maximum;
    }

    /// <summary>
    ///     Validates that the numeric value is within the range.
    /// </summary>
    public ValidationResult Validate(ValidateArgs args)
    {
        if (args.Value == null)
        {
            return ValidationResult.Success();
        }

        double doubleValue;
        try
        {
            doubleValue = Convert.ToDouble(args.Value, CultureInfo.InvariantCulture);
        }
        catch
        {
            return ValidationResult.Failed($"Field '{args.FieldName}' at row {args.Row}, column {args.Column} must be a number.");
        }

        if (doubleValue < _minimum || doubleValue > _maximum)
        {
            return ValidationResult.Failed(
                $"Field '{args.FieldName}' at row {args.Row}, column {args.Column} must be between {_minimum} and {_maximum}.");
        }

        return ValidationResult.Success();
    }
}
