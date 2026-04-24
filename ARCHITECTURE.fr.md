# Architecture d'ExcelHelper

## Vue d'ensemble

ExcelHelper est une bibliothèque de lecture et d'écriture Excel (XLSX) pour .NET, architecturalement inspirée par CsvHelper mais adaptée au domaine Excel. Elle est construite sur EPPlus 7+.

La bibliothèque suit les **principes SOLID** avec une séparation stricte des préoccupations entre le mappage, la conversion de types, la validation et les E/S.

## Architecture de haut niveau

```
┌─────────────────────────────────────────────────────────────┐
│                        Code consommateur                    │
├─────────────────────────────────────────────────────────────┤
│  ExcelReader / ExcelWriter                                  │
│  ├─ ExcelConfiguration (culture, en-têtes, rappels)        │
│  ├─ ExcelContext (ReadingContext / WritingContext)          │
│  └─ ExcelClassMap<T> (définitions de mappage)              │
├─────────────────────────────────────────────────────────────┤
│  Pipeline                                                   │
│  ├─ Couche de mappage     (ExcelMemberMap, AutoMap)        │
│  ├─ Couche de conversion  (IExcelTypeConverter, Cache)     │
│  ├─ Couche de validation  (IExcelFieldValidator, etc.)     │
│  └─ Couche E/S            (EPPlus ExcelWorksheet)          │
├─────────────────────────────────────────────────────────────┤
│  Infrastructure                                             │
│  ├─ ReflectionHelper                                        │
│  ├─ CompiledExpressionCache                                 │
│  └─ SharedStringTableHelper                                 │
└─────────────────────────────────────────────────────────────┘
```

## Flux de données

### Pipeline de lecture

```
ExcelWorksheet
    ↓
Énumération des lignes (yield)
    ↓
Extraction de la valeur de cellule (Text / Value / Formula)
    ↓
Conversion OADate (si DateTime et UseOADate=true)
    ↓
TypeConverter (valeur brute → type .NET cible)
    ↓
Validation de champ
    ↓
Affectation de propriété (expression compilée)
    ↓
Validation d'enregistrement
    ↓
yield return T
```

### Pipeline d'écriture

```
Objet T
    ↓
Lecture de propriété (getter d'expression compilée)
    ↓
TypeConverter (type .NET → valeur compatible Excel)
    ↓
Affectation de cellule (Value / Formula / Style)
    ↓
Vidage de ligne optionnel
    ↓
Sauvegarde dans le flux
```

## Modules principaux

### 1. Configuration (`Core/`)

`ExcelConfiguration` est l'objet de configuration central de type immutable. Il contient :
- Les informations de culture pour les conversions
- Les options d'en-tête (ligne, présence)
- Le ciblage de feuille (nom ou index)
- Les rappels (`ReadingExceptionOccurred`, `BadDataFound`, etc.)
- Les indicateurs de comportement (`UseOADate`, `TrimCellValues`, `IgnoreBlankRows`)

### 2. Mappage (`Mapping/`)

`ExcelClassMap<T>` définit comment un type .NET est mappé vers les colonnes Excel.
- `Map(m => m.Property)` crée un `ExcelMemberMap`
- Les mappages de membres prennent en charge : `Index()`, `Name()`, `Default()`, `TypeConverter<>()`, `Ignore()`, `Optional()`
- `AutoMap` génère des mappages via réflexion, en respectant `[ExcelColumn]`, `[ExcelIgnore]`, etc.
- `CompiledExpressionCache` stocke les getters/setters compilés pour la performance

### 3. Conversion de types (`TypeConversion/`)

`IExcelTypeConverter<T>` gère la conversion entre les valeurs de cellule Excel et les types .NET.
- Convertisseurs par défaut pour : `string`, `int`, `long`, `double`, `decimal`, `DateTime`, `bool`, `enum`, `Guid`, `Nullable<T>`
- `ExcelTypeConverterCache` met en cache les instances de convertisseurs
- `OADateConverter` gère le format de date sériel d'Excel

### 4. Validation (`Validation/`)

Validation à deux niveaux :
- **Validation de champ** : par cellule, après conversion
- **Validation d'enregistrement** : par ligne, après matérialisation de l'objet

### 5. Exceptions (`Exceptions/`)

Exceptions hiérarchiques :
- `ExcelHelperException` (base)
- `ExcelValidationException`
- `ExcelTypeConversionException`
- `ExcelMappingException`

## Points d'extension

| Point d'extension | Interface / Classe de base | Enregistrement |
|---|---|---|
| Convertisseur de type personnalisé | `IExcelTypeConverter<T>` | `ExcelConfiguration.TypeConverterCache` |
| Validateur de champ personnalisé | `IExcelFieldValidator` | API Fluent sur `ExcelMemberMap` |
| Validateur d'enregistrement personnalisé | `IExcelRecordValidator` | API Fluent sur `ExcelClassMap` |
| Résolveur d'objet personnalisé | `IObjectResolver` | `ExcelConfiguration.ObjectResolver` |
| Mappage personnalisé | Hériter de `ExcelClassMap<T>` | Enregistrer via `ExcelConfiguration.RegisterClassMap` |

## Stratégie multi-framework

```xml
<TargetFrameworks>net10.0;net9.0;net8.0;net48;net47;netstandard2.1;netstandard2.0</TargetFrameworks>
```

### Matrice de compatibilité

| TFM | `IAsyncEnumerable<T>` | `Span<T>` | Propriétés `init` | Notes |
|---|---|---|---|---|
| `net10.0` | ✅ Natif | ✅ Natif | ✅ Natif | Dernier |
| `net9.0` | ✅ Natif | ✅ Natif | ✅ Natif | |
| `net8.0` | ✅ Natif | ✅ Natif | ✅ Natif | LTS |
| `netstandard2.1` | ✅ Natif | ❌ Non | ❌ Non | |
| `netstandard2.0` | ❌ Fallback `Task<IReadOnlyList<T>>` | ❌ Non | ❌ Non | Nécessite `Microsoft.Bcl.AsyncInterfaces` |
| `net48` / `net47` | ❌ Fallback `Task<IReadOnlyList<T>>` | ❌ Non | ❌ Non | Nécessite `Microsoft.Bcl.AsyncInterfaces` |

### Symboles de compilation

Utilisez `#if NETCOREAPP` pour les API spécifiques à .NET Core+ (Span, IAsyncEnumerable).
Utilisez `#if NETFRAMEWORK` pour les contournements spécifiques à .NET Framework.

## Stratégie de performance

1. **Cache de réflexion** : `ConcurrentDictionary` pour les recherches d'informations de propriété
2. **Expressions compilées** : `Expression.Lambda` pour les getters/setters, mis en cache par type
3. **Cache de convertisseurs** : Convertisseurs singletons stockés dans `ExcelTypeConverterCache`
4. **Traitement par lots** : Lecture/écriture en tailles de lots configurables
5. **Réduction des allocations** : Réutilisation de `StringBuilder`, évitement du boxing dans les convertisseurs

## Décisions techniques clés

1. **EPPlus 7+** : API moderne, maintenance active. L'appelant est responsable de `ExcelPackage.LicenseContext`.
2. **Traitement basé sur DOM** : EPPlus charge l'intégralité du classeur. Le streaming est simulé via `yield return` sur les enregistrements, pas sur le fichier sous-jacent.
3. **API asynchrone double** : Différents types de retour par TFM pour maximiser la compatibilité sans forcer les polyfills sur les runtimes modernes.
4. **API Fluent** : Toute la configuration de mappage et de validation prend en charge le chaînage de méthodes.
5. **Séparation du contexte** : `ReadingContext` et `WritingContext` transportent l'état séparément pour éviter la contamination croisée.

## Structure des fichiers

```
src/ExcelHelper/
├── Core/              # Configuration, contexte, arguments de pipeline
├── Mapping/           # Mappages de classes, mappages de membres, auto-mapping, attributs
├── TypeConversion/    # Convertisseurs, fabrique, cache, OADate
├── Validation/        # Validateurs de champ, validateurs d'enregistrement, extensions fluent
├── Exceptions/        # Hiérarchie d'exceptions
├── Internal/          # Helpers de réflexion, cache d'expressions, chaînes partagées
└── Extensions/        # Méthodes d'extension pour les types EPPlus
```
