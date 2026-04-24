# Guide — Configuration

`ExcelConfiguration` centralise tous les paramètres de lecture et d'écriture.

## Propriétés principales

### Gestion des feuilles

```csharp
var config = new ExcelConfiguration
{
    SheetName = "Sales",   // Lire/écrire dans la feuille nommée
    SheetIndex = 0         // Ou utiliser l'index 0-based (défaut : 0)
};
```

> `SheetName` a la priorité sur `SheetIndex`.

### En-têtes et lignes de données

```csharp
var config = new ExcelConfiguration
{
    HasHeaderRecord = true,   // La feuille a une ligne d'en-tête (défaut : true)
    HeaderRow = 1,            // Ligne 1-based de l'en-tête (défaut : 1)
    StartRow = 2              // Ligne 1-based où commencent les données (défaut : 2)
};
```

### Nettoyage des données

```csharp
var config = new ExcelConfiguration
{
    TrimCellValues = true,      // Supprime les espaces autour des chaînes (défaut : false)
    IgnoreBlankRows = true,     // Ignore les lignes entièrement vides (défaut : true)
    IgnoreReferences = false    // Ignore les références d'erreur Excel (#REF!, etc.)
};
```

### Culture et dates

```csharp
var config = new ExcelConfiguration
{
    CultureInfo = CultureInfo.GetCultureInfo("fr-FR"), // Culture pour conversions (défaut : Invariant)
    UseOADate = true                                    // Interprète les nombres comme OADate (défaut : true)
};
```

## Hooks (callbacks)

### `ReadingExceptionOccurred`

Intercepte les exceptions lors de la lecture. Retournez `true` pour ignorer et continuer.

```csharp
var config = new ExcelConfiguration
{
    ReadingExceptionOccurred = ex =>
    {
        Console.WriteLine($"Read error: {ex.Message}");
        return true; // ignorer la ligne
    }
};
```

### `WritingExceptionOccurred`

Même principe pour l'écriture :

```csharp
var config = new ExcelConfiguration
{
    WritingExceptionOccurred = ex =>
    {
        Console.WriteLine($"Write error: {ex.Message}");
        return false; // propager l'exception
    }
};
```

### `MissingFieldFound`

Appelé quand un champ obligatoire est manquant :

```csharp
var config = new ExcelConfiguration
{
    MissingFieldFound = (headers, row, context) =>
    {
        Console.WriteLine($"Missing field at row {row}");
    }
};
```

### `BadDataFound`

Appelé quand une conversion échoue avant que l'exception ne soit levée :

```csharp
var config = new ExcelConfiguration
{
    BadDataFound = args =>
    {
        Console.WriteLine($"Bad data in field '{args.Field}' at [{args.Row},{args.Column}]: '{args.RawCellValue}'");
        return true; // ignorer et continuer
    }
};
```

### `PrepareHeaderForMatch`

Transforme les noms d'en-tête avant comparaison :

```csharp
var config = new ExcelConfiguration
{
    PrepareHeaderForMatch = header => header?.Trim().ToUpperInvariant() ?? string.Empty
};
```

## Object Resolver

Personnalisez la création des instances lors du mapping :

```csharp
public class ServiceResolver : IObjectResolver
{
    private readonly IServiceProvider _services;
    public ServiceResolver(IServiceProvider services) => _services = services;

    public object Resolve(Type type)
    {
        return _services.GetService(type) ?? Activator.CreateInstance(type)!;
    }
}

var config = new ExcelConfiguration
{
    ObjectResolver = new ServiceResolver(serviceProvider)
};
```

## Enregistrement des class maps

```csharp
var config = new ExcelConfiguration();
config.RegisterClassMap(new PersonMap());
config.RegisterClassMap(new OrderMap());
```

La méthode retourne `this` pour le chaînage :

```csharp
var config = new ExcelConfiguration()
    .RegisterClassMap(new PersonMap())
    .RegisterClassMap(new OrderMap());
```

## Exemple complet

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Core;
using System.Globalization;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var config = new ExcelConfiguration
{
    SheetName = "Data",
    HasHeaderRecord = true,
    HeaderRow = 1,
    StartRow = 2,
    TrimCellValues = true,
    IgnoreBlankRows = true,
    CultureInfo = CultureInfo.GetCultureInfo("fr-FR"),
    UseOADate = true,
    ReadingExceptionOccurred = ex =>
    {
        Console.WriteLine($"Skipped row: {ex.Message}");
        return true;
    },
    BadDataFound = args =>
    {
        Console.WriteLine($"Bad data at {args.Row},{args.Column}: {args.RawCellValue}");
        return true;
    }
};

using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream, config);

var records = reader.GetRecords<MyRecord>().ToList();
```
