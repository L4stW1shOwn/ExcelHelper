# Guide — Type Conversion

ExcelHelper automatically converts Excel cell values to .NET types through the `IExcelTypeConverter<T>` system.

## Natively Supported Types

| .NET Type | Behavior |
|---|---|
| `string` | Direct conversion |
| `int` / `int?` | Integer parse |
| `long` / `long?` | Long integer parse |
| `double` / `double?` | Floating-point parse |
| `decimal` / `decimal?` | Decimal parse |
| `bool` / `bool?` | `TRUE`/`1` → `true`, `FALSE`/`0` → `false` |
| `DateTime` / `DateTime?` | ISO string parse or **OADate** Excel |
| `Guid` / `Guid?` | GUID parse |
| `enum` / `enum?` | Parse by name or value |

## OADate

Excel stores dates as numbers (OADate). By default, `ExcelConfiguration.UseOADate = true`:

```csharp
// Excel cell containing 45292 (represents 2024-01-15)
var config = new ExcelConfiguration { UseOADate = true };
using var reader = new ExcelReader(stream, config);

var date = reader.GetField<DateTime>(0); // 2024-01-15
```

## Culture

Conversion uses `ExcelConfiguration.CultureInfo` (default: `InvariantCulture`):

```csharp
var config = new ExcelConfiguration
{
    CultureInfo = CultureInfo.GetCultureInfo("fr-FR")
};

// "1,5" will be interpreted as 1.5 in fr-FR
```

## Custom Converter

Implement `IExcelTypeConverter<T>`:

```csharp
using ExcelHelper.TypeConversion;
using System.Globalization;

public class EurosConverter : IExcelTypeConverter<decimal>
{
    public decimal ConvertFromExcel(object? value, CultureInfo cultureInfo)
    {
        if (value is string s)
        {
            // "1 234,56 EUR" → 1234.56m
            var cleaned = s.Replace("EUR", "").Replace(" ", "").Trim();
            return decimal.Parse(cleaned, cultureInfo);
        }
        return Convert.ToDecimal(value, cultureInfo);
    }

    public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
    {
        if (value is decimal d)
            return d.ToString("N2", cultureInfo) + " EUR";
        return value;
    }
}
```

Register it in the mapping:

```csharp
Map(m => m.Price).Index(2).TypeConverter<EurosConverter>();
```

Or via attribute:

```csharp
[ExcelColumn(Converter = typeof(EurosConverter))]
public decimal Price { get; set; }
```

## `ExcelTypeConverterCache`

Converters are created once and cached globally:

```csharp
// Clear the cache if needed (tests, dynamic reload)
ExcelTypeConverterCache.Clear();
```

## Full Example

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Mapping;
using ExcelHelper.TypeConversion;
using System.Globalization;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

public class Order
{
    public int Id { get; set; }
    public DateTime OrderDate { get; set; }
    public decimal Total { get; set; }
    public Status Status { get; set; }
}

public enum Status { Pending, Shipped, Delivered }

public class OrderMap : ExcelClassMap<Order>
{
    public OrderMap()
    {
        Map(m => m.Id).Index(0);
        Map(m => m.OrderDate).Index(1).TypeConverter<DateTimeConverter>();
        Map(m => m.Total).Index(2).TypeConverter<DecimalConverter>();
        Map(m => m.Status).Index(3).TypeConverter<EnumConverter<Status>>();
    }
}

var config = new ExcelConfiguration
{
    CultureInfo = CultureInfo.GetCultureInfo("en-US"),
    UseOADate = true
};
config.RegisterClassMap(new OrderMap());

using var stream = File.OpenRead("orders.xlsx");
using var reader = new ExcelReader(stream, config);

var orders = reader.GetRecords<Order>().ToList();
```
