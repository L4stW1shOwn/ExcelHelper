using System;
using System.Globalization;
using ExcelHelper.Internal;
using ExcelHelper.TypeConversion;
using Xunit;

namespace ExcelHelper.Tests;

public class ExcelTypeConverterTests
{
    private static readonly CultureInfo Culture = CultureInfo.InvariantCulture;

    [Fact]
    public void StringConverter_Should_Convert_From_Excel()
    {
        var converter = new StringConverter();
        Assert.Equal("hello", converter.ConvertFromExcel("hello", Culture));
        Assert.Equal("123", converter.ConvertFromExcel(123, Culture));
        Assert.Equal(string.Empty, converter.ConvertFromExcel(null, Culture));
    }

    [Fact]
    public void StringConverter_Should_Convert_To_Excel()
    {
        var converter = new StringConverter();
        Assert.Equal("hello", converter.ConvertToExcel("hello", Culture));
        Assert.Null(converter.ConvertToExcel(null, Culture));
    }

    [Fact]
    public void Int32Converter_Should_Convert_From_Excel()
    {
        var converter = new Int32Converter();
        Assert.Equal(42, converter.ConvertFromExcel(42, Culture));
        Assert.Equal(42, converter.ConvertFromExcel(42.0, Culture));
        Assert.Equal(42, converter.ConvertFromExcel(42m, Culture));
        Assert.Equal(42, converter.ConvertFromExcel(42L, Culture));
        Assert.Equal(42, converter.ConvertFromExcel("42", Culture));
        Assert.Equal(0, converter.ConvertFromExcel(null, Culture));
    }

    [Fact]
    public void Int32Converter_Should_Convert_To_Excel()
    {
        var converter = new Int32Converter();
        Assert.Equal(42, converter.ConvertToExcel(42, Culture));
        Assert.Null(converter.ConvertToExcel(null, Culture));
    }

    [Fact]
    public void Int64Converter_Should_Convert_From_Excel()
    {
        var converter = new Int64Converter();
        Assert.Equal(42L, converter.ConvertFromExcel(42, Culture));
        Assert.Equal(42L, converter.ConvertFromExcel(42L, Culture));
        Assert.Equal(42L, converter.ConvertFromExcel("42", Culture));
        Assert.Equal(0L, converter.ConvertFromExcel(null, Culture));
    }

    [Fact]
    public void DoubleConverter_Should_Convert_From_Excel()
    {
        var converter = new DoubleConverter();
        Assert.Equal(3.14, converter.ConvertFromExcel(3.14, Culture));
        Assert.Equal(3.14, converter.ConvertFromExcel("3.14", Culture));
        Assert.Equal(42.0, converter.ConvertFromExcel(42, Culture));
        Assert.Equal(0.0, converter.ConvertFromExcel(null, Culture));
    }

    [Fact]
    public void DecimalConverter_Should_Convert_From_Excel()
    {
        var converter = new DecimalConverter();
        Assert.Equal(3.14m, converter.ConvertFromExcel(3.14m, Culture));
        Assert.Equal(3.14m, converter.ConvertFromExcel(3.14, Culture));
        Assert.Equal(42m, converter.ConvertFromExcel(42, Culture));
        Assert.Equal(0m, converter.ConvertFromExcel(null, Culture));
    }

    [Fact]
    public void BooleanConverter_Should_Convert_From_Excel()
    {
        var converter = new BooleanConverter();
        Assert.True(converter.ConvertFromExcel(true, Culture));
        Assert.True(converter.ConvertFromExcel("true", Culture));
        Assert.True(converter.ConvertFromExcel("1", Culture));
        Assert.True(converter.ConvertFromExcel("yes", Culture));
        Assert.True(converter.ConvertFromExcel(1, Culture));
        Assert.False(converter.ConvertFromExcel(false, Culture));
        Assert.False(converter.ConvertFromExcel("false", Culture));
        Assert.False(converter.ConvertFromExcel("0", Culture));
        Assert.False(converter.ConvertFromExcel("no", Culture));
        Assert.False(converter.ConvertFromExcel(0, Culture));
        Assert.False(converter.ConvertFromExcel(null, Culture));
    }

    [Fact]
    public void DateTimeConverter_Should_Convert_From_Excel()
    {
        var converter = new DateTimeConverter();
        var date = new DateTime(2024, 1, 15);
        Assert.Equal(date, converter.ConvertFromExcel(date, Culture));
        Assert.Equal(date, converter.ConvertFromExcel(date.ToString("O"), Culture));
        Assert.Equal(date, converter.ConvertFromExcel(45306.0, Culture)); // OADate for 2024-01-15
        Assert.Equal(default, converter.ConvertFromExcel(null, Culture));
    }

    [Fact]
    public void DateTimeConverter_Should_Convert_To_Excel()
    {
        var converter = new DateTimeConverter();
        var date = new DateTime(2024, 1, 15);
        var result = converter.ConvertToExcel(date, Culture);
        Assert.Equal(45306.0, result);
    }

    [Fact]
    public void GuidConverter_Should_Convert_From_Excel()
    {
        var converter = new GuidConverter();
        var guid = Guid.NewGuid();
        Assert.Equal(guid, converter.ConvertFromExcel(guid, Culture));
        Assert.Equal(guid, converter.ConvertFromExcel(guid.ToString(), Culture));
        Assert.Equal(default, converter.ConvertFromExcel(null, Culture));
    }

    [Fact]
    public void GuidConverter_Should_Throw_For_Invalid_String()
    {
        var converter = new GuidConverter();
        Assert.Throws<InvalidOperationException>(() => converter.ConvertFromExcel("not-a-guid", Culture));
    }

    [Fact]
    public void EnumConverter_Should_Convert_From_Excel()
    {
        var converter = new EnumConverter<DayOfWeek>();
        Assert.Equal(DayOfWeek.Monday, converter.ConvertFromExcel(DayOfWeek.Monday, Culture));
        Assert.Equal(DayOfWeek.Monday, converter.ConvertFromExcel("Monday", Culture));
        Assert.Equal(DayOfWeek.Monday, converter.ConvertFromExcel("monday", Culture));
        Assert.Equal(DayOfWeek.Monday, converter.ConvertFromExcel("1", Culture));
        Assert.Equal(DayOfWeek.Monday, converter.ConvertFromExcel(1, Culture));
        Assert.Equal(default, converter.ConvertFromExcel(null, Culture));
    }

    [Fact]
    public void EnumConverter_Should_Throw_For_Invalid_Value()
    {
        var converter = new EnumConverter<DayOfWeek>();
        Assert.Throws<InvalidOperationException>(() => converter.ConvertFromExcel("InvalidDay", Culture));
    }

    [Fact]
    public void NullableConverter_Should_Convert_Null()
    {
        var inner = new Int32Converter();
        var converter = new NullableConverter<int>(inner);
        Assert.Null(converter.ConvertFromExcel(null, Culture));
        Assert.Null(converter.ConvertFromExcel(string.Empty, Culture));
        Assert.Null(converter.ConvertFromExcel("   ", Culture));
    }

    [Fact]
    public void NullableConverter_Should_Convert_Value()
    {
        var inner = new Int32Converter();
        var converter = new NullableConverter<int>(inner);
        Assert.Equal(42, converter.ConvertFromExcel(42, Culture));
        Assert.Equal(42, converter.ConvertFromExcel("42", Culture));
    }

    [Fact]
    public void NullableConverter_Should_Convert_To_Excel()
    {
        var inner = new Int32Converter();
        var converter = new NullableConverter<int>(inner);
        Assert.Equal(42, converter.ConvertToExcel(42, Culture));
        Assert.Null(converter.ConvertToExcel(null, Culture));
    }

    [Fact]
    public void ExcelTypeConverterCache_Should_Return_Cached_Converters()
    {
        var cache = new ExcelTypeConverterCache();

        var stringConverter = cache.GetConverter<string>();
        Assert.IsType<StringConverter>(stringConverter);

        var intConverter = cache.GetConverter<int>();
        Assert.IsType<Int32Converter>(intConverter);

        var doubleConverter = cache.GetConverter<double>();
        Assert.IsType<DoubleConverter>(doubleConverter);

        var decimalConverter = cache.GetConverter<decimal>();
        Assert.IsType<DecimalConverter>(decimalConverter);

        var boolConverter = cache.GetConverter<bool>();
        Assert.IsType<BooleanConverter>(boolConverter);

        var dateTimeConverter = cache.GetConverter<DateTime>();
        Assert.IsType<DateTimeConverter>(dateTimeConverter);

        var guidConverter = cache.GetConverter<Guid>();
        Assert.IsType<GuidConverter>(guidConverter);

        var enumConverter = cache.GetConverter<DayOfWeek>();
        Assert.IsType<EnumConverter<DayOfWeek>>(enumConverter);
    }

    [Fact]
    public void ExcelTypeConverterCache_Should_Return_Same_Instance()
    {
        var cache = new ExcelTypeConverterCache();
        var c1 = cache.GetConverter<int>();
        var c2 = cache.GetConverter<int>();
        Assert.Same(c1, c2);
    }

    [Fact]
    public void ExcelTypeConverterCache_Should_Handle_Nullable_Types()
    {
        var cache = new ExcelTypeConverterCache();
        var converter = cache.GetConverter(typeof(int?));
        Assert.NotNull(converter);
    }

    [Fact]
    public void OADateConverter_Should_Convert_Correctly()
    {
        var date = new DateTime(2024, 1, 15);
        var oaDate = OADateConverter.ToOADate(date);
        Assert.Equal(45306.0, oaDate);

        var back = OADateConverter.FromOADate(oaDate);
        Assert.Equal(date, back);
    }

    [Fact]
    public void ExcelTypeConverterCache_Should_Throw_For_Unsupported_Type()
    {
        var cache = new ExcelTypeConverterCache();
        Assert.Throws<NotSupportedException>(() => cache.GetConverter(typeof(ExcelTypeConverterTests)));
    }
}