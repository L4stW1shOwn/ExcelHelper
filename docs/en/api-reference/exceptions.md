# API Reference — Exceptions

Namespace: `ExcelHelper.Exceptions`

## `ExcelHelperException`

Base exception for all ExcelHelper errors.

```csharp
public class ExcelHelperException : Exception
```

### Constructors

```csharp
public ExcelHelperException()
public ExcelHelperException(string message)
public ExcelHelperException(string message, Exception innerException)
```

---

## `ExcelMappingException`

Thrown in case of mapping configuration error or value assignment error.

```csharp
public class ExcelMappingException : ExcelHelperException
```

### Properties

| Property | Type | Description |
|---|---|---|
| `MappedType` | `Type?` | Type currently being mapped. |
| `MemberName` | `string?` | Name of the member concerned. |

### Constructors

```csharp
public ExcelMappingException(string message, Type? mappedType, string? memberName)
public ExcelMappingException(string message, Type? mappedType, string? memberName, Exception innerException)
```

---

## `ExcelTypeConversionException`

Thrown when conversion of a cell value fails.

```csharp
public class ExcelTypeConversionException : ExcelHelperException
```

### Properties

| Property | Type | Description |
|---|---|---|
| `FieldName` | `string?` | Field name. |
| `FieldValue` | `object?` | Raw value of the cell. |
| `TargetType` | `Type` | Target type of the conversion. |
| `Row` | `int` | 1-based row. |
| `Column` | `int` | 1-based column. |

### Constructors

```csharp
public ExcelTypeConversionException(string message, object? fieldValue, Type targetType, string? fieldName, int row, int column)
public ExcelTypeConversionException(string message, object? fieldValue, Type targetType, string? fieldName, int row, int column, Exception innerException)
```

---

## `ExcelValidationException`

Thrown when a validation rule fails.

```csharp
public class ExcelValidationException : ExcelHelperException
```

### Properties

| Property | Type | Description |
|---|---|---|
| `FieldName` | `string?` | Field name (null for record validation). |
| `Row` | `int` | 1-based row. |
| `Column` | `int` | 1-based column. |

### Constructors

```csharp
public ExcelValidationException(string message, string? fieldName, int row, int column)
public ExcelValidationException(string message, string? fieldName, int row, int column, Exception innerException)
```

---

## Example — Targeted catch

```csharp
try
{
    var records = reader.GetRecords<Person>().ToList();
}
catch (ExcelTypeConversionException ex)
{
    Console.WriteLine($"Conversion error at [{ex.Row},{ex.Column}]: '{ex.FieldValue}' → {ex.TargetType.Name}");
}
catch (ExcelValidationException ex)
{
    Console.WriteLine($"Validation failed at [{ex.Row},{ex.Column}]: {ex.Message}");
}
catch (ExcelMappingException ex)
{
    Console.WriteLine($"Mapping error on {ex.MemberName}: {ex.Message}");
}
catch (ExcelHelperException ex)
{
    Console.WriteLine($"ExcelHelper error: {ex.Message}");
}
```
