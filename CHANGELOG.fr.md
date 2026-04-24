# Journal des modifications

Tous les changements notables de ce projet seront documentés dans ce fichier.

## [1.0.0] - 2026-04-23

### Ajouté
- Version initiale d'ExcelHelper
- `ExcelReader` et `ExcelWriter` pour les opérations de lecture/écriture XLSX
- `ExcelClassMap<T>` et `ExcelMemberMap<TClass, TMember>` pour le mappage fluent
- Auto-mapping via réflexion avec support des attributs (`[ExcelColumn]`, `[ExcelIgnore]`, `[ExcelIndex]`, `[ExcelName]`, `[ExcelDefault]`)
- Pipeline de conversion de types avec convertisseurs par défaut (string, int, long, double, decimal, DateTime, bool, enum, Guid, Nullable<T>)
- Conversion OADate pour les dates sérielles Excel
- Validation au niveau du champ (`IExcelFieldValidator`) avec validateurs intégrés (NotNull, NotEmpty, GreaterThan, Range)
- Validation au niveau de l'enregistrement (`IExcelRecordValidator`)
- Support asynchrone : `IAsyncEnumerable<T>` sur .NET Core+ / netstandard2.1, `Task<IReadOnlyList<T>>` sur .NET Framework / netstandard2.0
- Hooks : `ReadingExceptionOccurred`, `WritingExceptionOccurred`, `MissingFieldFound`, `BadDataFound`
- Support multi-feuilles via `SheetName` et `SheetIndex`
- Gestion des en-têtes (`HasHeaderRecord`, `HeaderRow`)
- Ignorance des lignes vides (`IgnoreBlankRows`)
- Conversion sensible à la culture (`CultureInfo`)
- `CompiledExpressionCache` pour un accès performant aux propriétés
- Support multi-cible : `net10.0;net9.0;net8.0;net48;net47;netstandard2.1;netstandard2.0`
- Suite de tests xUnit complète
