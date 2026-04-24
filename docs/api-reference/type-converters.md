# Référence API — Type Converters

Namespace : `ExcelHelper.TypeConversion`

## `IExcelTypeConverter<T>`

Interface à implémenter pour créer un convertisseur personnalisé.

```csharp
public interface IExcelTypeConverter<T>
{
    T ConvertFromExcel(object? value, CultureInfo cultureInfo);
    object? ConvertToExcel(object? value, CultureInfo cultureInfo);
}
```

| Méthode | Description |
|---|---|
| `ConvertFromExcel` | Convertit une valeur Excel vers le type .NET `T`. |
| `ConvertToExcel` | Convertit une valeur .NET `T` vers une valeur compatible Excel. |

---

## Converters intégrés

| Converter | Type .NET |
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
| `NullableConverter<T>` | `T?` (wrapper autour d'un converter interne) |

---

## `ExcelTypeConverterCache`

Cache global des instances de converters.

```csharp
public sealed class ExcelTypeConverterCache
```

### Méthodes

```csharp
public IExcelTypeConverter<T> GetConverter<T>()
public object GetConverter(Type type)
public static void Clear()
```

---

## Exemple — Converter personnalisé

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

Utilisation :

```csharp
Map(m => m.IsActive).Index(4).TypeConverter<YesNoBooleanConverter>();
```

---

## Exemple — Converter avec formatage

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

Utilisation (via instance) :

```csharp
Map(m => m.BirthDate).Index(1).Validate(new FormattedDateConverter("dd/MM/yyyy"));
```

> **Note** : `TypeConverter<T>()` accepte uniquement des types sans paramètres de constructeur. Pour les converters avec paramètres, utilisez `Validate()` avec une instance ou enregistrez le converter manuellement dans `ExcelTypeConverterCache`.
