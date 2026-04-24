# API Reference — Type Converters

Namespace: `ExcelHelper.TypeConversion`

## `IExcelTypeConverter<T>`

Interface to implement to create a custom converter.

```csharp
public interface IExcelTypeConverter<T>
{
    T ConvertFromExcel(object? value, CultureInfo cultureInfo);
    object? ConvertToExcel(object? value, CultureInfo cultureInfo);
}
```

| Method | Description |
|---|---|
| `ConvertFromExcel` | Converts an Excel value to .NET type `T`. |
| `ConvertToExcel` | Converts a .NET `T` value to an Excel-compatible value. |

---

## Built-in Converters

| Converter | .NET Type |
|---|---|
| `StringConverter` | `string` |
| `Int32Converter` | `int` / `int?` |
| `Int64Converter` | `long` / `long?` |
| `DoubleConverter` | `double` / `double?` |
| `DecimalConverter` | `decimal` / `decimal?` |
| `BooleanConverter` | `bool` / `bool?` |
| `DateTimeConverter` | `DateTime` / `DateTime?` |
| `GuidConverter` | `Guid` / `Guid?` |
| `EnumConverter<TEnum>` | `TEnum` / `TEnum?` |
| `NullableConverter<T>` | `T?` (wrapper around an internal converter) |

---

## `ExcelTypeConverterCache`

Global cache of converter instances.

```csharp
public sealed class ExcelTypeConverterCache
```

### Methods

```csharp
public IExcelTypeConverter<T> GetConverter<T>()
public object GetConverter(Type type)
public static void Clear()
```

---

## Example — Custom Converter

```csharp
using ExcelHelper.TypeConversion;
using System.Globalization;

public class YesNoBooleanConverter : IExcelTypeConverter<bool>
{
    public bool ConvertFromExcel(object? value, CultureInfo cultureInfo)
    {
        if (value is string s)
        {
            return s.Trim().Equals("YES", StringComparison.OrdinalIgnoreCase);
        }
        return Convert.ToBoolean(value, cultureInfo);
    }

    public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
    {
        if (value is bool b)
            return b ? "YES" : "NO";
        return value;
    }
}
```

Usage:

```csharp
Map(m => m.IsActive).Index(4).TypeConverter<YesNoBooleanConverter>();
```

---

## Example — Converter with formatting

```csharp
public class FormattedDateConverter : IExcelTypeConverter<DateTime>
{
    private readonly string _format;

    public FormattedDateConverter(string format)
    {
        _format = format;
    }

    public DateTime ConvertFromExcel(object? value, CultureInfo cultureInfo)
    {
        if (value is string s)
            return DateTime.ParseExact(s, _format, cultureInfo);

        return Convert.ToDateTime(value, cultureInfo);
    }

    public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
    {
        if (value is DateTime dt)
            return dt.ToString(_format, cultureInfo);
        return value;
    }
}
```

Usage (via instance):

```csharp
Map(m => m.BirthDate).Index(1).Validate(new FormattedDateConverter("dd/MM/yyyy"));
```

> **Note**: `TypeConverter<T>()` only accepts types without constructor parameters. For converters with parameters, use `Validate()` with an instance or register the converter manually in `ExcelTypeConverterCache`.
