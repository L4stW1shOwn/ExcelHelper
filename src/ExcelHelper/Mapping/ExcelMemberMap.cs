using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelHelper.Internal;
using ExcelHelper.Validation;

namespace ExcelHelper.Mapping;

/// <summary>
///     Represents the mapping data for a member.
/// </summary>
public sealed class MemberMapData
{
    /// <summary>
    ///     Gets or sets the <see cref="MemberInfo" /> being mapped.
    /// </summary>
    public MemberInfo Member { get; set; } = null!;

    /// <summary>
    ///     Gets or sets the 1-based column index.
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    ///     Gets or sets the column name.
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    ///     Gets or sets the default value when the cell is empty.
    /// </summary>
    public object? Default { get; set; }

    /// <summary>
    ///     Gets or sets the type of the converter to use.
    /// </summary>
    public Type? TypeConverter { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether this member is ignored.
    /// </summary>
    public bool Ignore { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether this member is optional.
    /// </summary>
    public bool IsOptional { get; set; }

    /// <summary>
    ///     Gets the validators for this member.
    /// </summary>
    public List<object> Validators { get; } = [];
}

/// <summary>
///     Represents the mapping for a class member.
/// </summary>
/// <typeparam name="TClass">The class type.</typeparam>
/// <typeparam name="TMember">The member type.</typeparam>
public class ExcelMemberMap<TClass, TMember>
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelMemberMap{TClass, TMember}" /> class.
    /// </summary>
    /// <param name="member">The member information.</param>
    public ExcelMemberMap(MemberInfo member)
    {
        Data.Member = member ?? throw new ArgumentNullException(nameof(member));
        Data.Name = member.Name;
    }

    /// <summary>
    ///     Gets the mapping data for this member.
    /// </summary>
    public MemberMapData Data { get; } = new();

    /// <summary>
    ///     Sets the 0-based column index for this member.
    /// </summary>
    /// <param name="index">The column index.</param>
    /// <returns>This <see cref="ExcelMemberMap{TClass, TMember}" /> instance.</returns>
    public ExcelMemberMap<TClass, TMember> Index(int index)
    {
        Data.Index = index;
        return this;
    }

    /// <summary>
    ///     Sets the column name for this member.
    /// </summary>
    /// <param name="name">The column name.</param>
    /// <returns>This <see cref="ExcelMemberMap{TClass, TMember}" /> instance.</returns>
    public ExcelMemberMap<TClass, TMember> Name(string name)
    {
        Data.Name = name ?? throw new ArgumentNullException(nameof(name));
        return this;
    }

    /// <summary>
    ///     Sets the default value for this member when the cell is empty.
    /// </summary>
    /// <param name="defaultValue">The default value.</param>
    /// <returns>This <see cref="ExcelMemberMap{TClass, TMember}" /> instance.</returns>
    public ExcelMemberMap<TClass, TMember> Default(TMember defaultValue)
    {
        Data.Default = defaultValue;
        return this;
    }

    /// <summary>
    ///     Sets the type converter for this member.
    /// </summary>
    /// <typeparam name="TConverter">The type of the converter.</typeparam>
    /// <returns>This <see cref="ExcelMemberMap{TClass, TMember}" /> instance.</returns>
    public ExcelMemberMap<TClass, TMember> TypeConverter<TConverter>()
        where TConverter : class
    {
        Data.TypeConverter = typeof(TConverter);
        return this;
    }

    /// <summary>
    ///     Sets a value indicating whether this member is ignored.
    /// </summary>
    /// <returns>This <see cref="ExcelMemberMap{TClass, TMember}" /> instance.</returns>
    public ExcelMemberMap<TClass, TMember> Ignore()
    {
        Data.Ignore = true;
        return this;
    }

    /// <summary>
    ///     Sets a value indicating whether this member is optional.
    /// </summary>
    /// <returns>This <see cref="ExcelMemberMap{TClass, TMember}" /> instance.</returns>
    public ExcelMemberMap<TClass, TMember> Optional()
    {
        Data.IsOptional = true;
        return this;
    }

    /// <summary>
    ///     Gets the compiled getter for this member.
    /// </summary>
    /// <returns>A compiled getter delegate.</returns>
    public Func<TClass, TMember> GetGetter()
    {
        return CompiledExpressionCache.GetGetter<TClass, TMember>(Data.Member.Name);
    }

    /// <summary>
    ///     Adds a validator for this member.
    /// </summary>
    /// <typeparam name="TValidator">The type of the validator.</typeparam>
    /// <returns>This <see cref="ExcelMemberMap{TClass, TMember}" /> instance.</returns>
    public ExcelMemberMap<TClass, TMember> Validate<TValidator>()
        where TValidator : IExcelFieldValidator, new()
    {
        Data.Validators.Add(new TValidator());
        return this;
    }

    /// <summary>
    ///     Adds a validator for this member.
    /// </summary>
    /// <param name="validator">The validator instance.</param>
    /// <returns>This <see cref="ExcelMemberMap{TClass, TMember}" /> instance.</returns>
    public ExcelMemberMap<TClass, TMember> Validate(IExcelFieldValidator validator)
    {
        if (validator == null)
        {
            throw new ArgumentNullException(nameof(validator));
        }

        Data.Validators.Add(validator);
        return this;
    }

    /// <summary>
    ///     Adds a validator for this member using a boolean predicate.
    /// </summary>
    /// <param name="predicate">The predicate that returns <c>true</c> if valid.</param>
    /// <returns>This <see cref="ExcelMemberMap{TClass, TMember}" /> instance.</returns>
    public ExcelMemberMap<TClass, TMember> Validate(Func<ValidateArgs, bool> predicate)
    {
        if (predicate == null)
        {
            throw new ArgumentNullException(nameof(predicate));
        }

        Data.Validators.Add(new FuncValidator(predicate));
        return this;
    }

    /// <summary>
    ///     Gets the compiled setter for this member.
    /// </summary>
    /// <returns>A compiled setter delegate.</returns>
    public Action<TClass, TMember> GetSetter()
    {
        return CompiledExpressionCache.GetSetter<TClass, TMember>(Data.Member.Name);
    }
}