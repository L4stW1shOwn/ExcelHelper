# Référence API — ExcelConfiguration

Namespace : `ExcelHelper.Core`

## Description

Centralise tous les paramètres de configuration pour les opérations de lecture et d'écriture.

## Propriétés

| Propriété | Type | Défaut | Description |
|---|---|---|---|
| `Maps` | `ExcelClassMapCollection` | `new` | Collection des class maps enregistrées. |
| `ObjectResolver` | `IObjectResolver` | `DefaultObjectResolver` | Résolveur d'instances pour le mapping. |
| `CultureInfo` | `CultureInfo` | `InvariantCulture` | Culture utilisée pour les conversions. |
| `HasHeaderRecord` | `bool` | `true` | Indique si la feuille a une ligne d'en-tête. |
| `HeaderRow` | `int` | `1` | Ligne 1-based de l'en-tête. |
| `StartRow` | `int` | `2` | Ligne 1-based où commencent les données. |
| `SheetName` | `string?` | `null` | Nom de la feuille à utiliser. |
| `SheetIndex` | `int` | `0` | Index 0-based de la feuille. |
| `UseOADate` | `bool` | `true` | Interprète les nombres comme des OADate. |
| `TrimCellValues` | `bool` | `false` | Supprime les espaces autour des chaînes. |
| `IgnoreBlankRows` | `bool` | `true` | Ignore les lignes entièrement vides. |
| `IgnoreReferences` | `bool` | `false` | Ignore les références d'erreur Excel. |
| `PrepareHeaderForMatch` | `Func<string, string>` | `Trim` | Transforme les en-têtes avant comparaison. |
| `ReadingExceptionOccurred` | `Func<ExcelHelperException, bool>` | `_ => false` | Hook appelé lors d'une erreur de lecture. |
| `WritingExceptionOccurred` | `Func<ExcelHelperException, bool>` | `_ => false` | Hook appelé lors d'une erreur d'écriture. |
| `MissingFieldFound` | `Action<string[], int, ReadingContext>?` | `null` | Hook appelé quand un champ obligatoire est manquant. |
| `BadDataFound` | `Func<BadDataFoundArgs, bool>` | `_ => false` | Hook appelé quand des données invalides sont détectées. |

## Méthodes

### `RegisterClassMap<T>`

```csharp
public ExcelConfiguration RegisterClassMap<T>(ExcelClassMap<T> map)
```

Enregistre un class map pour le type `T`. Retourne `this` pour le chaînage fluent.

## Exemples

### Configuration de base

```csharp
var config = new ExcelConfiguration
{
    SheetName = "Data",
    TrimCellValues = true,
    CultureInfo = CultureInfo.GetCultureInfo("fr-FR")
};
```

### Chaînage fluent

```csharp
var config = new ExcelConfiguration()
    .RegisterClassMap(new PersonMap())
    .RegisterClassMap(new OrderMap());
```

### Hooks

```csharp
var config = new ExcelConfiguration
{
    ReadingExceptionOccurred = ex =>
    {
        Console.WriteLine(ex.Message);
        return true;
    },
    BadDataFound = args =>
    {
        Console.WriteLine($"Bad data: {args.RawCellValue}");
        return true;
    }
};
```
