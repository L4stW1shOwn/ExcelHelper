# Guide — Conversion de types

ExcelHelper convertit automatiquement les valeurs de cellules Excel vers les types .NET via le système de `IExcelTypeConverter<T>`.

## Types supportés nativement

| Type .NET | Comportement |
|---|---|
| `string` | Conversion directe |
| `int` / `int?` | Parse entier |
| `long` / `long?` | Parse entier long |
| `double` / `double?` | Parse flottant |
| `decimal` / `decimal?` | Parse décimal |
| `bool` / `bool?` | `TRUE`/`1` → `true`, `FALSE`/`0` → `false` |
| `DateTime` / `DateTime?` | Parse chaîne ISO ou **OADate** Excel |
| `Guid` / `Guid?` | Parse GUID |
| `enum` / `enum?` | Parse par nom ou valeur |

## OADate

Excel stocke les dates sous forme de nombres (OADate). Par défaut, `ExcelConfiguration.UseOADate = true` :

```csharp
// Cellule Excel contenant 45292 (représente le 15/01/2024)
var config = new ExcelConfiguration { UseOADate = true };
using var reader = new ExcelReader(stream, config);

var date = reader.GetField<DateTime>(0); // 2024-01-15
```

## Culture

La conversion utilise `ExcelConfiguration.CultureInfo` (défaut : `InvariantCulture`) :

```csharp
var config = new ExcelConfiguration
{
    CultureInfo = CultureInfo.GetCultureInfo("fr-FR")
};

// "1,5" sera interprété comme 1.5 en fr-FR
```

## Converter personnalisé

Implémentez `IExcelTypeConverter<T>` :

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

Enregistrez-le dans le mapping :

```csharp
Map(m => m.Price).Index(2).TypeConverter<EurosConverter>();
```

Ou via attribut :

```csharp
[ExcelColumn(Converter = typeof(EurosConverter))]
public decimal Price { get; set; }
```

## `ExcelTypeConverterCache`

Les converters sont créés une seule fois et mis en cache globalement :

```csharp
// Vider le cache si nécessaire (tests, rechargement dynamique)
ExcelTypeConverterCache.Clear();
```

## Exemple complet

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
