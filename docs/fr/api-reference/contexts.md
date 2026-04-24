# Référence API — Contexts

Namespace : `ExcelHelper.Core`

## `ExcelContext`

Classe de base abstraite fournissant le contexte commun aux opérations de lecture et d'écriture.

```csharp
public abstract class ExcelContext
```

### Propriétés

| Propriété | Type | Description |
|---|---|---|
| `Configuration` | `ExcelConfiguration` | Configuration de l'opération en cours. |
| `Worksheet` | `ExcelWorksheet?` | Feuille Excel actuellement traitée. |
| `Row` | `int` | Ligne 1-based courante. |
| `Column` | `int` | Colonne 1-based courante. |
| `SheetName` | `string?` | Nom de la feuille (raccourci). |

---

## `ReadingContext`

Contexte spécifique aux opérations de lecture.

```csharp
public sealed class ReadingContext : ExcelContext
```

### Propriétés

| Propriété | Type | Description |
|---|---|---|
| `HeaderRecord` | `string[]?` | Tableau des noms d'en-tête si `HasHeaderRecord` est `true`. |
| `Record` | `string[]?` | Valeurs brutes de l'enregistrement courant. |
| `CurrentIndex` | `int` | Index 0-based de l'enregistrement courant dans les données. |
| `RowCount` | `int` | Nombre total de lignes dans la feuille. |

---

## `WritingContext`

Contexte spécifique aux opérations d'écriture.

```csharp
public sealed class WritingContext : ExcelContext
```

### Propriétés

| Propriété | Type | Description |
|---|---|---|
| `HasHeaderBeenWritten` | `bool` | Indique si l'en-tête a déjà été écrit. |

---

## `BadDataFoundArgs`

Arguments passés au callback `ExcelConfiguration.BadDataFound`.

```csharp
public sealed class BadDataFoundArgs
```

### Propriétés

| Propriété | Type | Description |
|---|---|---|
| `Field` | `string?` | Nom du champ concerné. |
| `RawCellValue` | `string?` | Valeur brute de la cellule. |
| `Row` | `int` | Ligne 1-based. |
| `Column` | `int` | Colonne 1-based. |

### Constructeur

```csharp
public BadDataFoundArgs(string? field, string? rawCellValue, int row, int column)
```

---

## Exemples

### Accès au contexte de lecture

```csharp
using var reader = new ExcelReader(stream);

foreach (var person in reader.GetRecords<Person>())
{
    Console.WriteLine($"Processing row {reader.Context.Row} / {reader.Context.RowCount}");
}
```

### Accès au contexte d'écriture

```csharp
using var writer = new ExcelWriter(stream);

writer.WriteRecords(people);

Console.WriteLine($"Last written row: {writer.Context.Row}");
```

### Utilisation dans `MissingFieldFound`

```csharp
var config = new ExcelConfiguration
{
    MissingFieldFound = (headers, row, context) =>
    {
        Console.WriteLine($"Sheet: {context.SheetName}, Row: {row}");
    }
};
```
