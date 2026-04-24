# Référence API — ExcelMemberMap&lt;TClass, TMember&gt;

Namespace : `ExcelHelper.Mapping`

## Description

Représente le mapping fluent d'un membre (propriété ou champ) vers une colonne Excel.

## Propriétés

| Propriété | Type | Description |
|---|---|---|
| `Data` | `MemberMapData` | Données brutes du mapping. |

## Méthodes

### `Index`

```csharp
public ExcelMemberMap<TClass, TMember> Index(int index)
```

Définit l'index de colonne 0-based.

---

### `Name`

```csharp
public ExcelMemberMap<TClass, TMember> Name(string name)
```

Définit le nom de la colonne (utilisé pour l'en-tête).

---

### `Default`

```csharp
public ExcelMemberMap<TClass, TMember> Default(TMember defaultValue)
```

Définit la valeur par défaut si la cellule est vide.

---

### `TypeConverter<TConverter>`

```csharp
public ExcelMemberMap<TClass, TMember> TypeConverter<TConverter>() where TConverter : class
```

Définit le convertisseur de type à utiliser.

---

### `Ignore`

```csharp
public ExcelMemberMap<TClass, TMember> Ignore()
```

Marque le membre comme ignoré.

---

### `Optional`

```csharp
public ExcelMemberMap<TClass, TMember> Optional()
```

Marque le membre comme optionnel (pas d'erreur si manquant).

---

### `Validate<TValidator>`

```csharp
public ExcelMemberMap<TClass, TMember> Validate<TValidator>() where TValidator : IExcelFieldValidator, new()
```

Ajoute un validateur de champ par type.

---

### `Validate`

```csharp
public ExcelMemberMap<TClass, TMember> Validate(IExcelFieldValidator validator)
```

Ajoute un validateur de champ par instance.

---

### `GetGetter`

```csharp
public Func<TClass, TMember> GetGetter()
```

Retourne un getter compilé pour ce membre.

---

### `GetSetter`

```csharp
public Action<TClass, TMember> GetSetter()
```

Retourne un setter compilé pour ce membre.

## Exemple

```csharp
Map(m => m.Price)
    .Index(3)
    .Name("Unit Price (EUR)")
    .Default(0.0m)
    .TypeConverter<DecimalConverter>()
    .Optional()
    .Validate(new GreaterThanValidator(0));
```
