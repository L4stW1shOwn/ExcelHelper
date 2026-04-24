# Documentation ExcelHelper

Bienvenue dans la documentation d'**ExcelHelper**, la bibliothèque .NET pour lire et écrire des fichiers Excel (XLSX) avec une API fluide inspirée de CsvHelper.

## Table des matières

### Guides d'utilisation

- [Démarrage rapide](getting-started.md) — Installation et premier exemple
- [Lecture de fichiers Excel](guides/reading.md) — `ExcelReader`, curseur de lignes, streaming
- [Écriture de fichiers Excel](guides/writing.md) — `ExcelWriter`, en-têtes, écriture par lot
- [Mapping](guides/mapping.md) — `ExcelClassMap`, attributs, auto-mapping
- [Validation](guides/validation.md) — Validateurs de champs et d'enregistrements
- [Conversion de types](guides/type-converters.md) — Converters intégrés et personnalisés
- [Configuration](guides/configuration.md) — `ExcelConfiguration`, culture, feuilles
- [Async](guides/async.md) — `IAsyncEnumerable<T>` et `Task<IReadOnlyList<T>>`
- [Gestion des erreurs](guides/error-handling.md) — Exceptions, hooks, `BadDataFound`

### Référence API

- [ExcelReader](api-reference/ExcelReader.md)
- [ExcelWriter](api-reference/ExcelWriter.md)
- [ExcelConfiguration](api-reference/ExcelConfiguration.md)
- [ExcelClassMap&lt;T&gt;](api-reference/ExcelClassMap.md)
- [ExcelMemberMap&lt;TClass, TMember&gt;](api-reference/ExcelMemberMap.md)
- [Attributs de mapping](api-reference/attributes.md)
- [Validation](api-reference/validation.md)
- [Exceptions](api-reference/exceptions.md)
- [Contexts](api-reference/contexts.md)
- [Type converters](api-reference/type-converters.md)

## Aperçu des fonctionnalités

| Fonctionnalité | Description |
|---|---|
| Lecture / Écriture XLSX | Basé sur EPPlus 7+ |
| Mapping objet | Par code fluent (`ExcelClassMap`) ou par attributs |
| Auto-mapping | Mappage automatique des propriétés publiques |
| Conversion de types | `string`, `int`, `long`, `double`, `decimal`, `bool`, `DateTime`, `Guid`, `enum`, nullables |
| Validation | Par champ et par enregistrement |
| Streaming | `IAsyncEnumerable<T>` (.NET Core+) ou `Task<IReadOnlyList<T>>` (.NET Framework) |
| Multi-cibles | `net10.0`, `net9.0`, `net8.0`, `net48`, `net47`, `netstandard2.1`, `netstandard2.0` |
| Multi-feuilles | Lecture/écriture par nom ou index de feuille |
| OADate | Conversion automatique des dates Excel |

## Installation

```bash
dotnet add package ExcelHelper
```

> **Important** : EPPlus 7+ nécessite un contexte de licence. Définissez-le avant d'utiliser ExcelHelper :
>
> ```csharp
> ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // ou Commercial
> ```
