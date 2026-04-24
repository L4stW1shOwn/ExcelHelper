# Guide — Écriture de fichiers Excel

Ce guide couvre toutes les méthodes d'écriture proposées par `ExcelWriter`.

## Création d'un writer

```csharp
using var stream = File.OpenWrite("output.xlsx");
using var writer = new ExcelWriter(stream);
```

Avec configuration :

```csharp
var config = new ExcelConfiguration { SheetName = "Products" };
using var writer = new ExcelWriter(stream, config);
```

## Écriture d'une collection

```csharp
var people = new List<Person>
{
    new() { Name = "Alice", Age = 30 },
    new() { Name = "Bob", Age = 25 }
};

writer.WriteRecords(people);
```

L'en-tête est écrit automatiquement si `HasHeaderRecord = true` (défaut).

## Écriture d'un seul enregistrement

```csharp
writer.WriteRecord(new Person { Name = "Charlie", Age = 35 });
writer.WriteRecord(new Person { Name = "Diana", Age = 28 });
```

Chaque appel écrit une ligne et avance le curseur interne.

## Écriture async

### .NET Core+ (`IAsyncEnumerable<T>`)

```csharp
async IAsyncEnumerable<Product> GetProductsAsync()
{
    yield return new Product { Name = "Widget", Price = 9.99m };
    yield return new Product { Name = "Gadget", Price = 19.99m };
}

await writer.WriteRecordsAsync(GetProductsAsync());
```

### .NET Framework (`IEnumerable<T>`)

```csharp
await writer.WriteRecordsAsync(products);
```

## En-têtes personnalisés

Les noms de colonnes dans l'en-tête proviennent du mapping :

- Nom de propriété par défaut
- `ExcelNameAttribute` ou `ExcelColumnAttribute`
- `.Name("...")` dans `ExcelClassMap`

```csharp
public class ProductMap : ExcelClassMap<Product>
{
    public ProductMap()
    {
        Map(m => m.Name).Index(0).Name("Product Name");
        Map(m => m.Price).Index(1).Name("Unit Price");
    }
}
```

L'en-tête généré sera :

| Product Name | Unit Price |
|---|---|

## Feuilles multiples

Par défaut, ExcelWriter crée une feuille nommée "Sheet1". Vous pouvez spécifier :

```csharp
var config = new ExcelConfiguration { SheetName = "Sales" };
using var writer = new ExcelWriter(stream, config);
```

## Écriture manuelle d'en-tête

```csharp
writer.WriteHeader<Product>();
```

Cela écrit l'en-tête immédiatement sans attendre le premier `WriteRecord`. Utile si vous devez insérer des lignes avant les données.

## Exemple complet

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Core;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var config = new ExcelConfiguration
{
    SheetName = "Invoices",
    StartRow = 2,
    HasHeaderRecord = true
};

var invoices = new List<Invoice>
{
    new() { Number = "INV-001", Amount = 1500.00m, Date = new DateTime(2024, 1, 15) },
    new() { Number = "INV-002", Amount = 2300.50m, Date = new DateTime(2024, 2, 10) }
};

using var stream = File.OpenWrite("invoices.xlsx");
using var writer = new ExcelWriter(stream, config);

writer.WriteRecords(invoices);

public class Invoice
{
    public string Number { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public DateTime Date { get; set; }
}
```

## Points importants

- Le package Excel est sauvegardé dans le stream lors de l'appel à `Dispose()`.
- Si `leaveOpen = false` (défaut), le stream est fermé automatiquement.
- Les exceptions lors de l'écriture peuvent être interceptées via `ExcelConfiguration.WritingExceptionOccurred`.
