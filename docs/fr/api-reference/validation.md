# Référence API — Validation

Namespace : `ExcelHelper.Validation`

## Interfaces

### `IExcelFieldValidator`

```csharp
public interface IExcelFieldValidator
{
    ValidationResult Validate(object? value, string? fieldName, int row, int column);
}
```

Valide la valeur d'une cellule individuelle.

### `IExcelRecordValidator<T>`

```csharp
public interface IExcelRecordValidator<T>
{
    ValidationResult Validate(T record, int row);
}
```

Valide un enregistrement entier après que tous les champs ont été assignés.

---

## `ValidationResult`

```csharp
public sealed class ValidationResult
```

### Propriétés

| Propriété | Type | Description |
|---|---|---|
| `IsValid` | `bool` | `true` si la validation a réussi. |
| `ErrorMessage` | `string?` | Message d'erreur si `IsValid` est `false`. |

### Méthodes statiques

```csharp
public static ValidationResult Success()
public static ValidationResult Failed(string errorMessage)
```

---

## Validateurs intégrés

### `NotNullValidator`

```csharp
public sealed class NotNullValidator : IExcelFieldValidator
```

Vérifie que la valeur n'est pas `null`.

---

### `NotEmptyValidator`

```csharp
public sealed class NotEmptyValidator : IExcelFieldValidator
```

Vérifie que la chaîne n'est pas `null`, vide ou blanche.

---

### `GreaterThanValidator`

```csharp
public sealed class GreaterThanValidator : IExcelFieldValidator
```

Vérifie que la valeur numérique est strictement supérieure au minimum spécifié.

#### Constructeur

```csharp
public GreaterThanValidator(double minimum)
```

---

### `RangeValidator`

```csharp
public sealed class RangeValidator : IExcelFieldValidator
```

Vérifie que la valeur numérique est comprise dans l'intervalle `[minimum, maximum]`.

#### Constructeur

```csharp
public RangeValidator(double minimum, double maximum)
```

---

## Exemple — Validateur personnalisé

```csharp
public class EmailValidator : IExcelFieldValidator
{
    public ValidationResult Validate(object? value, string? fieldName, int row, int column)
    {
        if (value is not string s || !s.Contains("@"))
            return ValidationResult.Failed($"'{fieldName}' must be a valid email.");

        return ValidationResult.Success();
    }
}
```

---

## Exemple — Validateur d'enregistrement

```csharp
public class OrderValidator : IExcelRecordValidator<Order>
{
    public ValidationResult Validate(Order record, int row)
    {
        if (record.Quantity > 0 && record.UnitPrice <= 0)
            return ValidationResult.Failed($"Row {row}: Unit price must be positive when quantity is set.");

        return ValidationResult.Success();
    }
}
```
