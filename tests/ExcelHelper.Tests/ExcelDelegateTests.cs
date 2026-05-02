using System;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;
using Xunit;

namespace ExcelHelper.Tests;

public class ExcelDelegateTests
{
    [Fact]
    public void ReadingExceptionEventArgs_Stores_Context_And_Exception()
    {
        var config = new ExcelConfiguration();
        var context = new ReadingContext(config);
        var inner = new ExcelHelperException("test");
        var args = new ReadingExceptionEventArgs(context, inner);

        Assert.Same(context, args.ReadingContext);
        Assert.Same(inner, args.Exception);
        Assert.Same(context, args.Context);
    }

    [Fact]
    public void WritingExceptionEventArgs_Stores_Context_And_Exception()
    {
        var config = new ExcelConfiguration();
        var context = new WritingContext(config);
        var inner = new ExcelHelperException("test");
        var args = new WritingExceptionEventArgs(context, inner);

        Assert.Same(context, args.WritingContext);
        Assert.Same(inner, args.Exception);
        Assert.Same(context, args.Context);
    }

    [Fact]
    public void MissingFieldEventArgs_Stores_Context_FieldName_Row()
    {
        var config = new ExcelConfiguration();
        var context = new ReadingContext(config);
        var args = new MissingFieldEventArgs(context, "Name", 5);

        Assert.Same(context, args.ReadingContext);
        Assert.Equal("Name", args.FieldName);
        Assert.Equal(5, args.Row);
    }

    [Fact]
    public void BadDataFoundEventArgs_Stores_Context_FieldName_RawValue_Row_Column()
    {
        var config = new ExcelConfiguration();
        var context = new ReadingContext(config);
        var args = new BadDataFoundEventArgs(context, "Age", "abc", 3, 2);

        Assert.Same(context, args.ReadingContext);
        Assert.Equal("Age", args.FieldName);
        Assert.Equal("abc", args.RawValue);
        Assert.Equal(3, args.Row);
        Assert.Equal(2, args.Column);
    }

    [Fact]
    public void ValidationFailedEventArgs_Stores_Context_FieldName_FieldValue_ErrorMessage()
    {
        var config = new ExcelConfiguration();
        var context = new ReadingContext(config);
        var args = new ValidationFailedEventArgs(context, "Price", -5m, "Price must be positive");

        Assert.Same(context, args.Context);
        Assert.Equal("Price", args.FieldName);
        Assert.Equal(-5m, args.FieldValue);
        Assert.Equal("Price must be positive", args.ErrorMessage);
    }
}
