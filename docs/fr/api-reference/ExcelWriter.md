# Référence API — ExcelWriter

Namespace : `ExcelHelper`

## Description

Écrit des enregistrements Excel dans un `Stream`.

## Constructeur

```csharp
public ExcelWriter(Stream stream, ExcelConfiguration? configuration = null, bool leaveOpen = false)
```

| Paramètre | Description |
|---|---|
| `stream` | Stream d'écriture (obligatoire). |
| `configuration` | Configuration optionnelle. Défaut : `new ExcelConfiguration()`. |
| `leaveOpen` | Si `true`, le stream n'est pas fermé lors du `Dispose()`. |

## Propriétés

| Propriété | Type | Description |
|---|---|---|
| `Context` | `WritingContext` | Contexte d'écriture (ligne courante, etc.). |

## Méthodes

### `WriteHeader<T>`

```csharp
public void WriteHeader<T>()
```

Écrit la ligne d'en-tête pour le type `T` si `HasHeaderRecord` est `true` et que l'en-tête n'a pas encore été écrit.

---

### `WriteRecord<T>`

```csharp
public void WriteRecord<T>(T record)
```

Écrit un seul enregistrement. L'en-tête est écrit automatiquement si nécessaire.

---

### `WriteRecords<T>`

```csharp
public void WriteRecords<T>(IEnumerable<T> records)
```

Écrit une collection d'enregistrements.

---

### `WriteRecordsAsync<T>`

**.NET Core+ / `netstandard2.1`** :

```csharp
public Task WriteRecordsAsync<T>(IAsyncEnumerable<T> records, CancellationToken cancellationToken = default)
```

**.NET Framework / `netstandard2.0`** :

```csharp
public Task WriteRecordsAsync<T>(IEnumerable<T> records, CancellationToken cancellationToken = default)
```

Version asynchrone de `WriteRecords`.

---

### `Dispose`

```csharp
public void Dispose()
```

Sauvegarde le package dans le stream et libère les ressources.

## Exemples

### Écriture simple

```csharp
using var stream = File.OpenWrite("output.xlsx");
using var writer = new ExcelWriter(stream);

writer.WriteRecords(people);
```

### Écriture individuelle

```csharp
using var writer = new ExcelWriter(stream);

foreach (var person in people)
{
    writer.WriteRecord(person);
}
```

### Async (.NET Core+)

```csharp
await writer.WriteRecordsAsync(GetProductsAsync());
```
