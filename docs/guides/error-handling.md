# Guide — Gestion des erreurs

ExcelHelper fournit un mécanisme structuré pour gérer les erreurs de lecture et d'écriture via des exceptions dédiées et des hooks de callback.

## Hiérarchie des exceptions

```
ExcelHelperException (base)
├── ExcelMappingException
├── ExcelTypeConversionException
└── ExcelValidationException
```

Toutes héritent de `ExcelHelperException`, ce qui permet de les capturer globalement :

```csharp
try
{
    var records = reader.GetRecords<Person>().ToList();
}
catch (ExcelHelperException ex)
{
    Console.WriteLine($"ExcelHelper error: {ex.Message}");
}
```

## `ExcelMappingException`

Levée quand la configuration de mapping est invalide ou quand l'affectation d'une valeur échoue.

| Propriété | Description |
|---|---|
| `MappedType` | Type en cours de mapping. |
| `MemberName` | Nom du membre concerné. |

```csharp
catch (ExcelMappingException ex)
{
    Console.WriteLine($"Mapping failed on {ex.MemberName} of {ex.MappedType?.Name}");
}
```

## `ExcelTypeConversionException`

Levée quand la conversion d'une cellule vers le type .NET échoue.

| Propriété | Description |
|---|---|
| `FieldName` | Nom du champ. |
| `FieldValue` | Valeur brute de la cellule. |
| `TargetType` | Type cible de la conversion. |
| `Row` | Ligne 1-based. |
| `Column` | Colonne 1-based. |

```csharp
catch (ExcelTypeConversionException ex)
{
    Console.WriteLine($"Cannot convert '{ex.FieldValue}' to {ex.TargetType.Name} at [{ex.Row},{ex.Column}]");
}
```

## `ExcelValidationException`

Levée quand une règle de validation échoue.

| Propriété | Description |
|---|---|
| `FieldName` | Nom du champ (null pour validation d'enregistrement). |
| `Row` | Ligne 1-based. |
| `Column` | Colonne 1-based. |

```csharp
catch (ExcelValidationException ex)
{
    Console.WriteLine($"Validation failed at [{ex.Row},{ex.Column}]: {ex.Message}");
}
```

## Hooks de gestion d'erreurs

Au lieu de catcher les exceptions après coup, vous pouvez intercepter les problèmes pendant le traitement.

### `ReadingExceptionOccurred`

Retournez `true` pour ignorer la ligne et continuer ; `false` pour propager l'exception.

```csharp
var config = new ExcelConfiguration
{
    ReadingExceptionOccurred = ex =>
    {
        switch (ex)
        {
            case ExcelTypeConversionException tce:
                Console.WriteLine($"Bad value '{tce.FieldValue}' at row {tce.Row}, col {tce.Column}");
                return true; // ignorer la ligne

            case ExcelValidationException ve:
                Console.WriteLine($"Validation failed at row {ve.Row}: {ve.Message}");
                return true; // ignorer la ligne

            default:
                return false; // propager
        }
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
        return true; // ignorer et continuer
    }
};
```

### `BadDataFound`

Appelé **avant** que `ExcelTypeConversionException` ne soit levée. Retournez `true` pour ignorer le champ et continuer la ligne.

```csharp
var config = new ExcelConfiguration
{
    BadDataFound = args =>
    {
        Console.WriteLine($"Bad data in '{args.Field}' at [{args.Row},{args.Column}]: '{args.RawCellValue}'");
        return true; // ignorer le champ, continuer la ligne
    }
};
```

> Retourner `false` dans `BadDataFound` fait que l'exception `ExcelTypeConversionException` est levée normalement.

### `MissingFieldFound`

Appelé quand un champ obligatoire est vide :

```csharp
var config = new ExcelConfiguration
{
    MissingFieldFound = (headers, row, context) =>
    {
        Console.WriteLine($"Row {row} is missing a required field.");
    }
};
```

## Exemple complet — pipeline de lecture robuste

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var errors = new List<string>();

var config = new ExcelConfiguration
{
    TrimCellValues = true,
    IgnoreBlankRows = true,
    ReadingExceptionOccurred = ex =>
    {
        errors.Add($"[{ex.GetType().Name}] {ex.Message}");
        return true; // continuer malgré tout
    },
    BadDataFound = args =>
    {
        errors.Add($"Bad data at [{args.Row},{args.Column}]: '{args.RawCellValue}'");
        return true; // ignorer le champ et continuer la ligne
    },
    MissingFieldFound = (headers, row, context) =>
    {
        errors.Add($"Missing field at row {row}");
    }
};

using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream, config);

var records = reader.GetRecords<Person>().ToList();

if (errors.Count > 0)
{
    Console.WriteLine($"{errors.Count} error(s) occurred:");
    foreach (var error in errors)
        Console.WriteLine($"  - {error}");
}

Console.WriteLine($"Successfully loaded {records.Count} records.");
```
