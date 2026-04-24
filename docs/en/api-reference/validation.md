# API Reference — Validation

Namespace: `ExcelHelper.Validation`

## Interfaces

### `IExcelFieldValidator`

```csharp
public interface IExcelFieldValidator
{
    ValidationResult Validate(object? value, string? fieldName, int row, int column);
}
```

Validates the value of an individual cell.

### `IExcelRecordValidator<T>`

```csharp
public interface IExcelRecordValidator<T>
{
    ValidationResult Validate(T record, int row);
}
```

Validates an entire record after all fields have been assigned.

---

## `ValidationResult`

```csharp
public sealed class ValidationResult
```

### Properties

| Property | Type | Description |
|---|---|---|
| `IsValid` | `bool` | `true` if validation succeeded. |
| `ErrorMessage` | `string?` | Error message if `IsValid` is `false`. |

### Static Methods

```csharp
public static ValidationResult Success()
public static ValidationResult Failed(string errorMessage)
```

---

## Built-in Validators

### `NotNullValidator`

```csharp
public sealed class NotNullValidator : IExcelFieldValidator
```

Checks that the value is not `null`.

---

### `NotEmptyValidator`

```csharp
public sealed class NotEmptyValidator : IExcelFieldValidator
```

Checks that the string is not `null`, empty or whitespace.

---

### `GreaterThanValidator`

```csharp
public sealed class GreaterThanValidator : IExcelFieldValidator
```

Checks that the numeric value is strictly greater than the specified minimum.

#### Constructor

```csharp
public GreaterThanValidator(double minimum)
```

---

### `RangeValidator`

```csharp
public sealed class RangeValidator : IExcelFieldValidator
```

Checks that the numeric value is within the range `[minimum, maximum]`.

#### Constructor

```csharp
public RangeValidator(double minimum, double maximum)
```

---

## Example — Custom Validator

```csharp
public class EmailValidator : IExcelFieldValidator
{
    public ValidationResult Validate(object? value, string? fieldName, int row, int column)
    {
        if (value is not string s || !s.Contains("@"))
            return ValidationResult.Failed($"'{fieldName}' must be a valid email.");

        return ValidationResult.Success();
    }
}
```

---

## Example — Record Validator

```csharp
public class OrderValidator : IExcelRecordValidator<Order>
{
    public ValidationResult Validate(Order record, int row)
    {
        if (record.Quantity > 0 && record.UnitPrice <= 0)
            return ValidationResult.Failed($"Row {row}: Unit price must be positive when quantity is set.");

        return ValidationResult.Success();
    }
}
```
