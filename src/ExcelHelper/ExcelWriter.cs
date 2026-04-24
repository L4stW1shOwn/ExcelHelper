using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;
using ExcelHelper.Internal;
using ExcelHelper.Mapping;
using ExcelHelper.TypeConversion;
using OfficeOpenXml;

namespace ExcelHelper
{
    /// <summary>
    /// Writes Excel records to a stream.
    /// </summary>
    public sealed class ExcelWriter : IDisposable
    {
        private readonly ExcelPackage _package;
        private readonly ExcelConfiguration _configuration;
        private readonly WritingContext _context;
        private readonly ExcelWorksheet _worksheet;
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private bool _disposed;
        private readonly ExcelTypeConverterCache _converterCache;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <param name="configuration">The configuration.</param>
        /// <param name="leaveOpen">Whether to leave the stream open when disposing.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="stream"/> is null.</exception>
        public ExcelWriter(Stream stream, ExcelConfiguration? configuration = null, bool leaveOpen = false)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            _stream = stream;
            _leaveOpen = leaveOpen;
            _configuration = configuration ?? new ExcelConfiguration();
            _context = new WritingContext(_configuration);
            _converterCache = new ExcelTypeConverterCache();
            _package = new ExcelPackage();
            _worksheet = GetOrCreateWorksheet();
            _context.Worksheet = _worksheet;
            _context.Row = _configuration.StartRow;
        }

        /// <summary>
        /// Gets the writing context.
        /// </summary>
        public WritingContext Context => _context;

        /// <summary>
        /// Writes the header record if configured.
        /// </summary>
        /// <typeparam name="T">The record type.</typeparam>
        public void WriteHeader<T>()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ExcelWriter));

            if (!_configuration.HasHeaderRecord || _context.HasHeaderBeenWritten)
                return;

            var map = GetOrCreateMap<T>();
            var members = map.MemberMaps.ToList();

            int headerRow = _configuration.HeaderRow;
            foreach (var memberMap in members)
            {
                int col = memberMap.Index + 1;
                _worksheet.Cells[headerRow, col].Value = memberMap.Name;
            }

            _context.HasHeaderBeenWritten = true;
        }

        /// <summary>
        /// Writes a collection of records to the Excel file.
        /// </summary>
        /// <typeparam name="T">The record type.</typeparam>
        /// <param name="records">The records to write.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="records"/> is null.</exception>
        public void WriteRecords<T>(IEnumerable<T> records)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ExcelWriter));

            if (records == null)
                throw new ArgumentNullException(nameof(records));

            var map = GetOrCreateMap<T>();
            var members = map.MemberMaps.ToList();

            WriteHeader<T>();

            foreach (var record in records)
            {
                _context.Row = Math.Max(_context.Row, _configuration.StartRow);

                try
                {
                    WriteRecord(record, members);
                }
                catch (ExcelHelperException ex)
                {
                    if (_configuration.WritingExceptionOccurred(ex))
                        continue;
                    throw;
                }

                _context.Row++;
            }
        }

        /// <summary>
        /// Writes a single record to the Excel file.
        /// </summary>
        /// <typeparam name="T">The record type.</typeparam>
        /// <param name="record">The record to write.</param>
        public void WriteRecord<T>(T record)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ExcelWriter));

            var map = GetOrCreateMap<T>();
            var members = map.MemberMaps.ToList();

            WriteHeader<T>();
            WriteRecord(record, members);
            _context.Row++;
        }

        private void WriteRecord<T>(T record, List<MemberMapData> members)
        {
            foreach (var memberMap in members)
            {
                _context.Column = memberMap.Index + 1;

                if (memberMap.Ignore)
                    continue;

                var rawValue = GetMemberValue(record, memberMap);
                object? convertedValue;

                try
                {
                    convertedValue = ConvertToExcelValue(rawValue, memberMap);
                }
                catch (Exception ex)
                {
                    throw new ExcelTypeConversionException(
                        $"Failed to convert value '{rawValue}' for field '{memberMap.Name}'.",
                        rawValue,
                        GetMemberType(memberMap),
                        memberMap.Name,
                        _context.Row,
                        _context.Column,
                        ex);
                }

                try
                {
                    _worksheet.Cells[_context.Row, _context.Column].Value = convertedValue;
                }
                catch (Exception ex)
                {
                    throw new ExcelMappingException(
                        $"Failed to write value '{convertedValue}' to cell [{_context.Row},{_context.Column}].",
                        typeof(T),
                        memberMap.Member.Name,
                        ex);
                }
            }
        }

        private static object? GetMemberValue<T>(T record, MemberMapData memberMap)
        {
            if (memberMap.Member is PropertyInfo property)
                return property.GetValue(record);

            if (memberMap.Member is FieldInfo field)
                return field.GetValue(record);

            throw new InvalidOperationException($"Member '{memberMap.Member.Name}' is not a property or field.");
        }

        private object? ConvertToExcelValue(object? rawValue, MemberMapData memberMap)
        {
            if (rawValue == null)
                return null;

            var sourceType = GetMemberType(memberMap);
            var converter = _converterCache.GetConverter(sourceType);
            var convertMethod = converter.GetType().GetMethod("ConvertToExcel");
            return convertMethod?.Invoke(converter, new object?[] { rawValue, _configuration.CultureInfo });
        }

        private static Type GetMemberType(MemberMapData memberMap)
        {
            if (memberMap.Member is PropertyInfo property)
                return property.PropertyType;

            if (memberMap.Member is FieldInfo field)
                return field.FieldType;

            throw new InvalidOperationException($"Member '{memberMap.Member.Name}' is not a property or field.");
        }

        private ExcelWorksheet GetOrCreateWorksheet()
        {
            if (!string.IsNullOrEmpty(_configuration.SheetName))
            {
                var worksheet = _package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == _configuration.SheetName);
                if (worksheet == null)
                    worksheet = _package.Workbook.Worksheets.Add(_configuration.SheetName);
                return worksheet;
            }

            if (_configuration.SheetIndex >= 0 && _configuration.SheetIndex < _package.Workbook.Worksheets.Count)
                return _package.Workbook.Worksheets[_configuration.SheetIndex];

            return _package.Workbook.Worksheets.Add("Sheet1");
        }

        private ExcelClassMap<T> GetOrCreateMap<T>()
        {
            var map = _configuration.Maps.Get<T>();
            if (map != null)
                return map;

            map = new ExcelClassMap<T>();
            map.AutoMap();
            return map;
        }

#if NETSTANDARD2_1 || NETCOREAPP
        /// <summary>
        /// Asynchronously writes a collection of records to the Excel file.
        /// </summary>
        /// <typeparam name="T">The record type.</typeparam>
        /// <param name="records">The records to write.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public async Task WriteRecordsAsync<T>(IAsyncEnumerable<T> records, CancellationToken cancellationToken = default)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ExcelWriter));

            if (records == null)
                throw new ArgumentNullException(nameof(records));

            var map = GetOrCreateMap<T>();
            var members = map.MemberMaps.ToList();

            WriteHeader<T>();

            await foreach (var record in records.WithCancellation(cancellationToken))
            {
                _context.Row = Math.Max(_context.Row, _configuration.StartRow);

                try
                {
                    WriteRecord(record, members);
                }
                catch (ExcelHelperException ex)
                {
                    if (_configuration.WritingExceptionOccurred(ex))
                        continue;
                    throw;
                }

                _context.Row++;
                await Task.Yield();
            }
        }
#else
        /// <summary>
        /// Asynchronously writes a collection of records to the Excel file.
        /// </summary>
        /// <typeparam name="T">The record type.</typeparam>
        /// <param name="records">The records to write.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public async Task WriteRecordsAsync<T>(IEnumerable<T> records, CancellationToken cancellationToken = default)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ExcelWriter));

            if (records == null)
                throw new ArgumentNullException(nameof(records));

            await Task.Run(() => WriteRecords(records), cancellationToken);
        }
#endif

        /// <summary>
        /// Saves the Excel package to the stream and releases all resources.
        /// </summary>
        public void Dispose()
        {
            if (!_disposed)
            {
                _package.SaveAs(_stream);
                _package.Dispose();
                if (!_leaveOpen)
                    _stream.Dispose();
                _disposed = true;
            }
        }
    }
}
