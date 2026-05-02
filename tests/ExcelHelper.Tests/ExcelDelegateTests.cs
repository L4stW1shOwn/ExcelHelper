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
}
