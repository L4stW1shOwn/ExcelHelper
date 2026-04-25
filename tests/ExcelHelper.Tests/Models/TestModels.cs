using System;
using ExcelHelper.Mapping;
using ExcelHelper.Validation;

namespace ExcelHelper.Tests.Models;

public class SimplePerson
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public DateTime BirthDate { get; set; }
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public bool InStock { get; set; }
}

public class PersonWithAttributes
{
    [ExcelName("Full Name")] public string Name { get; set; } = string.Empty;

    [ExcelIndex(5)] public int Age { get; set; }

    [ExcelIgnore] public string Secret { get; set; } = string.Empty;

    [ExcelDefault(42)] public int DefaultValue { get; set; }
}

public class PersonWithIgnore
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }

    [ExcelIgnore] public string Secret { get; set; } = string.Empty;
}

public class DateRecord
{
    public DateTime EventDate { get; set; }
    public string EventName { get; set; } = string.Empty;
}

public class ValidatedPerson
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public decimal Salary { get; set; }
}

public class AgeRangeRecord
{
    public int Value { get; set; }
}

public class MinimalPerson
{
    public string Name { get; set; } = string.Empty;
}

public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.Name).Index(0);
        Map(m => m.Age).Index(1);
        Map(m => m.BirthDate).Index(2);
    }
}

public class ValidatedPersonMap : ExcelClassMap<ValidatedPerson>
{
    public ValidatedPersonMap()
    {
        Map(m => m.Name).Index(0).Validate<NotNullValidator>().Validate<NotEmptyValidator>();
        Map(m => m.Age).Index(1).Validate(new GreaterThanValidator(0));
        Map(m => m.Salary).Index(2).Optional();
    }
}

public class AgeRangeMap : ExcelClassMap<AgeRangeRecord>
{
    public AgeRangeMap()
    {
        Map(m => m.Value).Index(0).Validate(new RangeValidator(0, 100));
    }
}

public class RecordValidatedPersonMap : ExcelClassMap<ValidatedPerson>
{
    public RecordValidatedPersonMap()
    {
        Map(m => m.Name).Index(0);
        Map(m => m.Age).Index(1);
        Map(m => m.Salary).Index(2);
        Validate(new SalaryRecordValidator());
    }
}

public class SalaryRecordValidator : IExcelRecordValidator<ValidatedPerson>
{
    public ValidationResult Validate(ValidatedPerson record, int row)
    {
        if (record.Salary <= 0)
        {
            return ValidationResult.Failed("Salary must be greater than 0.");
        }

        return ValidationResult.Success();
    }
}