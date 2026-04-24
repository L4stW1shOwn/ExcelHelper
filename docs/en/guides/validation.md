# Guide — Validation

ExcelHelper allows you to validate data **per field** and **per record** during reading.

## Field Validation (`IExcelFieldValidator`)

Create a validator by implementing `IExcelFieldValidator`:

```csharp
using ExcelHelper.Validation;

public class NotEmptyValidator : IExcelFieldValidator
{
    public ValidationResult Validate(object? value, string? fieldName, int row, int column)
    {
        if (value == null || (value is string s && string.IsNullOrWhiteSpace(s)))
            return ValidationResult.Failed($"Field '{fieldName}' cannot be empty.");

        return ValidationResult.Success();
    }
}
```

### Built-in Validators

ExcelHelper provides several ready-to-use validators:

| Validator | Description |
|---|---|
| `NotNullValidator` | Checks that the value is not `null`. |
| `NotEmptyValidator` | Checks that the string is not empty or whitespace. |
| `GreaterThanValidator(double min)` | Numeric value > `min`. |
| `RangeValidator(double min, double max)` | Numeric value within [`min`, `max`]. |

### Usage in an `ExcelClassMap`

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.Name).Index(0).Validate<NotEmptyValidator>();
        Map(m => m.Age).Index(1).Validate(new RangeValidator(0, 120));
    }
}
```

### Usage with an Instance Validator

Useful when the validator needs initialization parameters:

```csharp
Map(m => m.Email)
    .Index(2)
    .Validate(new RegexValidator(@"^[^@]+@[^@]+$"));
```

## Record Validation (`IExcelRecordValidator<T>`)

To validate the overall consistency of a row (multiple fields together):

```csharp
public class PersonRecordValidator : IExcelRecordValidator<Person>
{
    public ValidationResult Validate(Person record, int row)
    {
        if (record.Age < 18 && record.Country == "US")
            return ValidationResult.Failed($"Row {row}: Must be 18+ in the US.");

        return ValidationResult.Success();
    }
}
```

Register it in the class map:

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.Name).Index(0);
        Map(m => m.Age).Index(1);
        Map(m => m.Country).Index(2);

        Validate<PersonRecordValidator>();       // by type
        Validate(new PersonRecordValidator());   // by instance
    }
}
```

## `ValidationResult`

```csharp
// Success
return ValidationResult.Success();

// Failure
return ValidationResult.Failed("Detailed error message");
```

On failure, an `ExcelValidationException` is thrown with:
- `Message`: the error message
- `FieldName`: the field name (null for record validation)
- `Row` / `Column`: 1-based position in the sheet

## Handling Validation Errors

You can intercept errors via the `ReadingExceptionOccurred` hook:

```csharp
var config = new ExcelConfiguration
{
    ReadingExceptionOccurred = ex =>
    {
        if (ex is ExcelValidationException vex)
        {
            Console.WriteLine($"Validation failed at row {vex.Row}, col {vex.Column}: {vex.Message}");
            return true; // ignore and continue
        }
        return false; // propagate the exception
    }
};
```

## Full Example

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Mapping;
using ExcelHelper.Validation;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

public class Employee
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public decimal Salary { get; set; }
}

public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(m => m.Name).Index(0).Validate<NotEmptyValidator>();
        Map(m => m.Age).Index(1).Validate(new RangeValidator(18, 65));
        Map(m => m.Salary).Index(2).Validate(new GreaterThanValidator(0));

        Validate(new SalaryAgeValidator());
    }
}

public class SalaryAgeValidator : IExcelRecordValidator<Employee>
{
    public ValidationResult Validate(Employee record, int row)
    {
        if (record.Age < 21 && record.Salary > 100000)
            return ValidationResult.Failed($"Row {row}: Salary too high for age {record.Age}.");

        return ValidationResult.Success();
    }
}

// Usage
var config = new ExcelConfiguration();
config.RegisterClassMap(new EmployeeMap());
config.ReadingExceptionOccurred = ex =>
{
    Console.WriteLine($"Error: {ex.Message}");
    return true; // continue despite errors
};

using var stream = File.OpenRead("employees.xlsx");
using var reader = new ExcelReader(stream, config);

var employees = reader.GetRecords<Employee>().ToList();
```
