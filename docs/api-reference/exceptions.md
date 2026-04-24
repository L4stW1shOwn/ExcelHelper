# Référence API — Exceptions

Namespace : `ExcelHelper.Exceptions`

## `ExcelHelperException`

Exception de base pour toutes les erreurs ExcelHelper.

```csharp
public class ExcelHelperException : Exception
```

### Constructeurs

```csharp
public ExcelHelperException()
public ExcelHelperException(string message)
public ExcelHelperException(string message, Exception innerException)
```

---

## `ExcelMappingException`

Levée en cas d'erreur de configuration de mapping ou d'affectation de valeur.

```csharp
public class ExcelMappingException : ExcelHelperException
```

### Propriétés

| Propriété | Type | Description |
|---|---|---|
| `MappedType` | `Type?` | Type en cours de mapping. |
| `MemberName` | `string?` | Nom du membre concerné. |

### Constructeurs

```csharp
public ExcelMappingException(string message, Type? mappedType, string? memberName)
public ExcelMappingException(string message, Type? mappedType, string? memberName, Exception innerException)
```

---

## `ExcelTypeConversionException`

Levée quand la conversion d'une valeur de cellule échoue.

```csharp
public class ExcelTypeConversionException : ExcelHelperException
```

### Propriétés

| Propriété | Type | Description |
|---|---|---|
| `FieldName` | `string?` | Nom du champ. |
| `FieldValue` | `object?` | Valeur brute de la cellule. |
| `TargetType` | `Type` | Type cible de la conversion. |
| `Row` | `int` | Ligne 1-based. |
| `Column` | `int` | Colonne 1-based. |

### Constructeurs

```csharp
public ExcelTypeConversionException(string message, object? fieldValue, Type targetType, string? fieldName, int row, int column)
public ExcelTypeConversionException(string message, object? fieldValue, Type targetType, string? fieldName, int row, int column, Exception innerException)
```

---

## `ExcelValidationException`

Levée quand une règle de validation échoue.

```csharp
public class ExcelValidationException : ExcelHelperException
```

### Propriétés

| Propriété | Type | Description |
|---|---|---|
| `FieldName` | `string?` | Nom du champ (null pour validation d'enregistrement). |
| `Row` | `int` | Ligne 1-based. |
| `Column` | `int` | Colonne 1-based. |

### Constructeurs

```csharp
public ExcelValidationException(string message, string? fieldName, int row, int column)
public ExcelValidationException(string message, string? fieldName, int row, int column, Exception innerException)
```

---

## Exemple — Capture ciblée

```csharp
try
{
    var records = reader.GetRecords<Person>().ToList();
}
catch (ExcelTypeConversionException ex)
{
    Console.WriteLine($"Conversion error at [{ex.Row},{ex.Column}]: '{ex.FieldValue}' → {ex.TargetType.Name}");
}
catch (ExcelValidationException ex)
{
    Console.WriteLine($"Validation failed at [{ex.Row},{ex.Column}]: {ex.Message}");
}
catch (ExcelMappingException ex)
{
    Console.WriteLine($"Mapping error on {ex.MemberName}: {ex.Message}");
}
catch (ExcelHelperException ex)
{
    Console.WriteLine($"ExcelHelper error: {ex.Message}");
}
```
