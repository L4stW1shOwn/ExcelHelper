# API Reference — ExcelReader

Namespace: `ExcelHelper`

## Description

Reads Excel records from a `Stream`.

## Constructor

```csharp
public ExcelReader(Stream stream, ExcelConfiguration? configuration = null, bool leaveOpen = false)
```

| Parameter | Description |
|---|---|
| `stream` | Read stream (required). |
| `configuration` | Optional configuration. Default: `new ExcelConfiguration()`. |
| `leaveOpen` | If `true`, the stream is not closed on `Dispose()`. |

## Properties

| Property | Type | Description |
|---|---|---|
| `Context` | `ReadingContext` | Read context (current row, headers, etc.). |

## Methods

### `Read`

```csharp
public bool Read()
```

Advances the reader to the next data row. Returns `false` if there are no more rows.

> Cannot be used after `GetRecords<T>()`.

---

### `GetField`

```csharp
public object? GetField(int index)
```

Returns the raw value of the cell at the specified 0-based index for the current row.

---

### `GetField<T>`

```csharp
public T? GetField<T>(int index)
```

Returns the cell value converted to type `T`.

Throws `ExcelTypeConversionException` if conversion fails.

---

### `TryGetField<T>`

```csharp
public bool TryGetField<T>(int index, out T? value)
```

Attempts to convert the value. Returns `false` without throwing an exception if conversion fails.

---

### `GetRecord<T>`

```csharp
public T GetRecord<T>()
```

Materializes the current row as an instance of `T` according to the configured mapping.

> The cursor is not advanced; you can call `GetRecord<T>()` multiple times on the same row with different types.

---

### `GetRecords<T>`

```csharp
public IEnumerable<T> GetRecords<T>()
```

Returns a lazy `IEnumerable<T>` reading all data rows.

> Cannot be used after `Read()`.

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

Async version of `GetRecords<T>`.

---

### `Dispose`

```csharp
public void Dispose()
```

Releases resources (EPPlus package and stream if `leaveOpen` is `false`).

## Examples

### Simple read

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);

var people = reader.GetRecords<Person>().ToList();
```

### Cursor mode

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
