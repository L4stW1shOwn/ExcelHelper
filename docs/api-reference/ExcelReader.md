# Référence API — ExcelReader

Namespace : `ExcelHelper`

## Description

Lit des enregistrements Excel depuis un `Stream`.

## Constructeur

```csharp
public ExcelReader(Stream stream, ExcelConfiguration? configuration = null, bool leaveOpen = false)
```

| Paramètre | Description |
|---|---|
| `stream` | Stream de lecture (obligatoire). |
| `configuration` | Configuration optionnelle. Défaut : `new ExcelConfiguration()`. |
| `leaveOpen` | Si `true`, le stream n'est pas fermé lors du `Dispose()`. |

## Propriétés

| Propriété | Type | Description |
|---|---|---|
| `Context` | `ReadingContext` | Contexte de lecture (ligne courante, en-têtes, etc.). |

## Méthodes

### `Read`

```csharp
public bool Read()
```

Avance le reader à la ligne de données suivante. Retourne `false` s'il n'y a plus de lignes.

> Ne peut pas être utilisé après `GetRecords<T>()`.

---

### `GetField`

```csharp
public object? GetField(int index)
```

Retourne la valeur brute de la cellule à l'index 0-based spécifié pour la ligne courante.

---

### `GetField<T>`

```csharp
public T? GetField<T>(int index)
```

Retourne la valeur de la cellule convertie en type `T`.

Lève `ExcelTypeConversionException` en cas d'échec de conversion.

---

### `TryGetField<T>`

```csharp
public bool TryGetField<T>(int index, out T? value)
```

Tente de convertir la valeur. Retourne `false` sans lever d'exception si la conversion échoue.

---

### `GetRecord<T>`

```csharp
public T GetRecord<T>()
```

Matérialise la ligne courante en instance de `T` selon le mapping configuré.

> Le curseur n'est pas avancé ; vous pouvez appeler `GetRecord<T>()` plusieurs fois sur la même ligne avec des types différents.

---

### `GetRecords<T>`

```csharp
public IEnumerable<T> GetRecords<T>()
```

Retourne un `IEnumerable<T>` lazy lisant toutes les lignes de données.

> Ne peut pas être utilisé après `Read()`.

---

### `GetRecordsAsync<T>`

**.NET Core+ / `netstandard2.1`** :

```csharp
public IAsyncEnumerable<T> GetRecordsAsync<T>(CancellationToken cancellationToken = default)
```

**.NET Framework / `netstandard2.0`** :

```csharp
public Task<IReadOnlyList<T>> GetRecordsAsync<T>(CancellationToken cancellationToken = default)
```

Version asynchrone de `GetRecords<T>`.

---

### `Dispose`

```csharp
public void Dispose()
```

Libère les ressources (package EPPlus et stream si `leaveOpen` est `false`).

## Exemples

### Lecture simple

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);

var people = reader.GetRecords<Person>().ToList();
```

### Mode curseur

```csharp
using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

while (reader.Read())
{
    var name = reader.GetField<string>(0);
    var age = reader.GetField<int>(1);
}
```

### Async (.NET Core+)

```csharp
await foreach (var person in reader.GetRecordsAsync<Person>())
{
    Console.WriteLine(person.Name);
}
```
