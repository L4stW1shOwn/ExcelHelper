using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;
using ExcelHelper.Extensions;
using ExcelHelper.Internal;
using ExcelHelper.Mapping;
using ExcelHelper.TypeConversion;
using ExcelHelper.Validation;
using OfficeOpenXml;

namespace ExcelHelper;

/// <summary>
///     Reads Excel records from a stream.
/// </summary>
public sealed class ExcelReader : IDisposable
{
    private readonly ExcelPackage _package;
    private readonly ExcelConfiguration _configuration;
    private readonly ExcelWorksheet _worksheet;
    private readonly Stream? _stream;
    private readonly bool _leaveOpen;
    private bool _disposed;
    private string[]? _headerRecord;
    private readonly ExcelTypeConverterCache _converterCache;

    private enum ReadingMode
    {
        None,
        Enumerable,
        Cursor
    }

    private ReadingMode _readingMode;
    private int _currentRow;
    private bool _hasRead;
    private bool _currentRowValid;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelReader" /> class.
    /// </summary>
    /// <param name="stream">The stream to read from.</param>
    /// <param name="configuration">The configuration.</param>
    /// <param name="leaveOpen">Whether to leave the stream open when disposing.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="stream" /> is null.</exception>
    public ExcelReader(Stream stream, ExcelConfiguration? configuration, bool leaveOpen = false)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _leaveOpen = leaveOpen;
        _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        Context = new ReadingContext(_configuration);
        _converterCache = new ExcelTypeConverterCache();
        _package = new ExcelPackage(stream);
        _worksheet = GetWorksheet();
        Context.Worksheet = _worksheet;
        Context.RowCount = _worksheet.Dimension?.End.Row ?? 0;

        if (_configuration.HasHeaderRecord)
        {
            ReadHeader();
        }
    }

    /// <summary>
    ///     Gets the reading context.
    /// </summary>
    public ReadingContext Context { get; }

    /// <summary>
    ///     Advances the reader to the next data row.
    /// </summary>
    /// <returns><c>true</c> if there are more rows; otherwise <c>false</c>.</returns>
    /// <exception cref="InvalidOperationException">Thrown when mixed with <see cref="GetRecords{T}" />.</exception>
    public bool Read()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(ExcelReader));
        }

        if (_readingMode == ReadingMode.Enumerable)
        {
            throw new InvalidOperationException("Cannot use Read() after GetRecords() has been called.");
        }

        _readingMode = ReadingMode.Cursor;

        if (!_hasRead)
        {
            _currentRow = _configuration.StartRow;
            _hasRead = true;
        }
        else
        {
            _currentRow++;
        }

        var endRow = _worksheet.Dimension?.End.Row ?? 0;

        while (_currentRow <= endRow)
        {
            Context.Row = _currentRow;
            Context.CurrentIndex = _currentRow - _configuration.StartRow;

            if (_configuration.IgnoreBlankRows && IsBlankRow(_currentRow))
            {
                _currentRow++;
                continue;
            }

            _currentRowValid = true;
            return true;
        }

        _currentRowValid = false;
        return false;
    }

    /// <summary>
    ///     Gets the raw cell value at the specified 0-based column index for the current row.
    /// </summary>
    /// <param name="index">The 0-based column index.</param>
    /// <returns>The raw cell value, or <c>null</c> if the cell is empty.</returns>
    /// <exception cref="InvalidOperationException">Thrown when Read() has not been called or when mixed with GetRecords().</exception>
    /// <exception cref="ObjectDisposedException">Thrown when the reader has been disposed.</exception>
    public object? GetField(int index)
    {
        EnsureCursorMode();
        var col = index + 1;
        var cell = _worksheet.Cells[_currentRow, col];
        var rawValue = SharedStringTableHelper.GetCellValue(cell);
        if (_configuration.TrimCellValues && rawValue is string s)
        {
            rawValue = s.Trim();
        }

        return rawValue;
    }

    /// <summary>
    ///     Gets the cell value at the specified 0-based column index, converted to type <typeparamref name="T" />.
    /// </summary>
    /// <typeparam name="T">The target type.</typeparam>
    /// <param name="index">The 0-based column index.</param>
    /// <returns>The converted value, or <c>default</c> if the cell is empty.</returns>
    /// <exception cref="InvalidOperationException">Thrown when Read() has not been called or when mixed with GetRecords().</exception>
    /// <exception cref="ObjectDisposedException">Thrown when the reader has been disposed.</exception>
    /// <exception cref="ExcelTypeConversionException">Thrown when conversion fails.</exception>
    public T? GetField<T>(int index)
    {
        var rawValue = GetField(index);
        if (rawValue == null)
        {
            return default;
        }

        var converter = _converterCache.GetConverter<T>();
        try
        {
            return converter.ConvertFromExcel(rawValue, _configuration.CultureInfo);
        }
        catch (Exception ex)
        {
            throw new ExcelTypeConversionException(
                $"Failed to convert value '{rawValue}' to type '{typeof(T).Name}'.",
                rawValue,
                typeof(T),
                null,
                _currentRow,
                index + 1,
                ex);
        }
    }

    /// <summary>
    ///     Attempts to get the cell value at the specified 0-based column index, converted to type <typeparamref name="T" />.
    /// </summary>
    /// <typeparam name="T">The target type.</typeparam>
    /// <param name="index">The 0-based column index.</param>
    /// <param name="value">The converted value if successful; otherwise <c>default</c>.</param>
    /// <returns><c>true</c> if conversion succeeded; otherwise <c>false</c>.</returns>
    public bool TryGetField<T>(int index, out T? value)
    {
        try
        {
            value = GetField<T>(index);
            return true;
        }
        catch (ExcelTypeConversionException)
        {
            value = default;
            return false;
        }
    }

    /// <summary>
    ///     Reads all records of type <typeparamref name="T" /> from the Excel file.
    /// </summary>
    /// <typeparam name="T">The record type.</typeparam>
    /// <returns>An enumerable of records.</returns>
    /// <exception cref="InvalidOperationException">Thrown when mixed with <see cref="Read()" />.</exception>
    public IEnumerable<T> GetRecords<T>()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(ExcelReader));
        }

        if (_readingMode == ReadingMode.Cursor)
        {
            throw new InvalidOperationException("Cannot use GetRecords() after Read() has been called.");
        }

        _readingMode = ReadingMode.Enumerable;

        var map = GetOrCreateMap<T>();
        var resolver = _configuration.ObjectResolver;
        var members = map.MemberMaps.ToList();
        var recordValidators = map.RecordValidators;

        for (var row = _configuration.StartRow; row <= (_worksheet.Dimension?.End.Row ?? 0); row++)
        {
            Context.Row = row;
            Context.CurrentIndex = row - _configuration.StartRow;

            if (_configuration.IgnoreBlankRows && IsBlankRow(row, members))
            {
                continue;
            }

            T record;
            try
            {
                record = ReadRecordAtRow<T>(row, members, resolver, recordValidators);
            }
            catch (ExcelHelperException ex)
            {
                var eventArgs = new ReadingExceptionEventArgs(Context, ex);
                if (_configuration.ReadingExceptionOccurred(eventArgs))
                {
                    continue;
                }

                throw;
            }

            yield return record;
        }
    }

    private void EnsureCursorMode()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(ExcelReader));
        }

        if (_readingMode == ReadingMode.Enumerable)
        {
            throw new InvalidOperationException("Cannot use cursor methods after GetRecords() has been called.");
        }

        if (!_currentRowValid)
        {
            throw new InvalidOperationException("Read() must be called before GetField.");
        }
    }

    private T ReadRecordAtRow<T>(int row, List<MemberMapData> members, IObjectResolver resolver,
        List<object> recordValidators)
    {
        Context.Row = row;
        var instance = (T)resolver.Resolve(typeof(T));

        foreach (var memberMap in members)
        {
            Context.Column = memberMap.Index + 1; // EPPlus is 1-based

            if (memberMap.Ignore)
            {
                continue;
            }

            var cell = _worksheet.Cells[Context.Row, Context.Column];
            var rawValue = SharedStringTableHelper.GetCellValue(cell);

            if (_configuration.TrimCellValues && rawValue is string s)
            {
                rawValue = s.Trim();
            }

            // Apply default if empty
            if (rawValue == null || (rawValue is string str && string.IsNullOrEmpty(str)))
            {
                if (memberMap.Default != null)
                {
                    rawValue = memberMap.Default;
                }
                else if (!memberMap.IsOptional)
                {
                    // Missing required field
                    _configuration.MissingFieldFound(new MissingFieldEventArgs(Context, memberMap.Name, Context.Row));
                }
            }

            // Type conversion
            object? convertedValue;
            try
            {
                convertedValue = ConvertValue(rawValue, memberMap);
            }
            catch (Exception ex)
            {
                var badDataArgs =
                    new BadDataFoundEventArgs(Context, memberMap.Name, rawValue, Context.Row, Context.Column);
                if (_configuration.BadDataFound(badDataArgs))
                {
                    continue;
                }

                throw new ExcelTypeConversionException(
                    $"Failed to convert value '{rawValue}' for field '{memberMap.Name}'.",
                    rawValue,
                    GetMemberType(memberMap),
                    memberMap.Name,
                    Context.Row,
                    Context.Column,
                    ex);
            }

            // Field validation
            foreach (var validator in memberMap.Validators)
            {
                if (validator is IExcelFieldValidator fieldValidator)
                {
                    var result = fieldValidator.Validate(convertedValue, memberMap.Name, Context.Row, Context.Column);
                    if (!result.IsValid)
                    {
                        throw new ExcelValidationException(
                            result.ErrorMessage ?? $"Validation failed for field '{memberMap.Name}'.",
                            memberMap.Name,
                            Context.Row,
                            Context.Column);
                    }
                }
            }

            // Property assignment
            if (convertedValue != null || GetMemberType(memberMap).IsClass ||
                ReflectionHelper.IsNullable(GetMemberType(memberMap)))
            {
                try
                {
                    SetMemberValue(instance, memberMap, convertedValue);
                }
                catch (Exception ex)
                {
                    throw new ExcelMappingException(
                        $"Failed to set value '{convertedValue}' to member '{memberMap.Member.Name}'.",
                        typeof(T),
                        memberMap.Member.Name,
                        ex);
                }
            }
        }

        // Record validation
        foreach (var validator in recordValidators)
        {
            if (validator is IExcelRecordValidator<T> recordValidator)
            {
                var result = recordValidator.Validate(instance, Context.Row);
                if (!result.IsValid)
                {
                    throw new ExcelValidationException(
                        result.ErrorMessage ?? "Record validation failed.",
                        null,
                        Context.Row,
                        0);
                }
            }
        }

        return instance;
    }

    private object? ConvertValue(object? rawValue, MemberMapData memberMap)
    {
        if (rawValue == null)
        {
            return null;
        }

        var targetType = GetMemberType(memberMap);
        var converter = _converterCache.GetConverter(targetType);
        var convertMethod = converter.GetType().GetMethod("ConvertFromExcel");
        return convertMethod?.Invoke(converter, new[] { rawValue, _configuration.CultureInfo });
    }

    private static void SetMemberValue<T>(T instance, MemberMapData memberMap, object? value)
    {
        if (memberMap.Member is PropertyInfo property)
        {
            property.SetValue(instance, value);
        }
        else if (memberMap.Member is FieldInfo field)
        {
            field.SetValue(instance, value);
        }
    }

    private static Type GetMemberType(MemberMapData memberMap)
    {
        if (memberMap.Member is PropertyInfo property)
        {
            return property.PropertyType;
        }

        if (memberMap.Member is FieldInfo field)
        {
            return field.FieldType;
        }

        throw new InvalidOperationException($"Member '{memberMap.Member.Name}' is not a property or field.");
    }

    private bool IsBlankRow(int row, List<MemberMapData> members)
    {
        if (_worksheet.Dimension == null)
        {
            return true;
        }

        var startCol = members.Count > 0 ? members.Min(m => m.Index) + 1 : 1;
        var endCol = members.Count > 0 ? members.Max(m => m.Index) + 1 : _worksheet.Dimension.End.Column;

        return _worksheet.IsBlankRow(row, startCol, endCol);
    }

    private bool IsBlankRow(int row)
    {
        if (_worksheet.Dimension == null)
        {
            return true;
        }

        return _worksheet.IsBlankRow(row, 1, _worksheet.Dimension.End.Column);
    }

    private void ReadHeader()
    {
        if (_worksheet.Dimension == null)
        {
            return;
        }

        var headerRow = _configuration.HeaderRow;
        var colCount = _worksheet.Dimension.End.Column;
        var headers = new List<string>();

        for (var col = 1; col <= colCount; col++)
        {
            var text = SharedStringTableHelper.GetCellText(_worksheet.Cells[headerRow, col]);
            headers.Add(_configuration.PrepareHeaderForMatch(text));
        }

        _headerRecord = headers.ToArray();
        Context.HeaderRecord = _headerRecord;
    }

    private ExcelWorksheet GetWorksheet()
    {
        if (!string.IsNullOrEmpty(_configuration.SheetName))
        {
            var worksheet = _package.Workbook.Worksheets[_configuration.SheetName];
            if (worksheet == null)
            {
                throw new ExcelMappingException($"Worksheet '{_configuration.SheetName}' not found.", null, null);
            }

            return worksheet;
        }

        if (_configuration.SheetIndex >= 0 && _configuration.SheetIndex < _package.Workbook.Worksheets.Count)
        {
            return _package.Workbook.Worksheets[_configuration.SheetIndex];
        }

        throw new ExcelMappingException($"Worksheet at index {_configuration.SheetIndex} not found.", null, null);
    }

    private ExcelClassMap<T> GetOrCreateMap<T>()
    {
        var map = _configuration.Maps.Get<T>();
        if (map != null)
        {
            return map;
        }

        map = new ExcelClassMap<T>();
        map.AutoMap();
        return map;
    }

#if NETSTANDARD2_1 || NETCOREAPP
    /// <summary>
    ///     Asynchronously reads all records of type <typeparamref name="T" /> from the Excel file.
    /// </summary>
    /// <typeparam name="T">The record type.</typeparam>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <returns>An async enumerable of records.</returns>
    public async IAsyncEnumerable<T> GetRecordsAsync<T>(
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(ExcelReader));
        }

        foreach (var record in GetRecords<T>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return record;
            await Task.Yield();
        }
    }
#else
        /// <summary>
        /// Asynchronously reads all records of type <typeparamref name="T"/> from the Excel file.
        /// </summary>
        /// <typeparam name="T">The record type.</typeparam>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A task containing the list of records.</returns>
        public async Task<IReadOnlyList<T>> GetRecordsAsync<T>(CancellationToken cancellationToken = default)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ExcelReader));

            return await Task.Run(() => GetRecords<T>().ToList(), cancellationToken);
        }
#endif

    /// <summary>
    ///     Gets the current row as a record of type <typeparamref name="T" />.
    ///     The cursor is not advanced — multiple calls for different types are allowed on the same row.
    /// </summary>
    /// <typeparam name="T">The record type.</typeparam>
    /// <returns>The materialised record.</returns>
    /// <exception cref="InvalidOperationException">Thrown when Read() has not been called or when mixed with GetRecords().</exception>
    /// <exception cref="ObjectDisposedException">Thrown when the reader has been disposed.</exception>
    public T GetRecord<T>()
    {
        EnsureCursorMode();
        var map = GetOrCreateMap<T>();
        var resolver = _configuration.ObjectResolver;
        var members = map.MemberMaps.ToList();
        var recordValidators = map.RecordValidators;

        Context.Column = 0;

        return ReadRecordAtRow<T>(_currentRow, members, resolver, recordValidators);
    }

    /// <summary>
    ///     Releases all resources used by the <see cref="ExcelReader" />.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _package.Dispose();
            if (!_leaveOpen)
            {
                _stream?.Dispose();
            }

            _disposed = true;
        }
    }
}