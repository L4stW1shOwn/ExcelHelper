# API Reference — ExcelWriter

Namespace: `ExcelHelper`

## Description

Writes Excel records into a `Stream`.

## Constructor

```csharp
public ExcelWriter(Stream stream, ExcelConfiguration? configuration = null, bool leaveOpen = false)
```

| Parameter | Description |
|---|---|
| `stream` | Write stream (required). |
| `configuration` | Optional configuration. Default: `new ExcelConfiguration()`. |
| `leaveOpen` | If `true`, the stream is not closed on `Dispose()`. |

## Properties

| Property | Type | Description |
|---|---|---|
| `Context` | `WritingContext` | Write context (current row, etc.). |

## Methods

### `WriteHeader<T>`

```csharp
public void WriteHeader<T>()
```

Writes the header row for type `T` if `HasHeaderRecord` is `true` and the header has not yet been written.

---

### `WriteRecord<T>`

```csharp
public void WriteRecord<T>(T record)
```

Writes a single record. The header is written automatically if necessary.

---

### `WriteRecords<T>`

```csharp
public void WriteRecords<T>(IEnumerable<T> records)
```

Writes a collection of records.

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

Async version of `WriteRecords`.

---

### `Dispose`

```csharp
public void Dispose()
```

Saves the package to the stream and releases resources.

## Examples

### Simple write

```csharp
using var stream = File.OpenWrite("output.xlsx");
using var writer = new ExcelWriter(stream);

writer.WriteRecords(people);
```

### Individual write

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
