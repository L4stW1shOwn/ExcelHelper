using ExcelHelper.Core;
using ExcelHelper.Mapping;
using OfficeOpenXml;

namespace ExcelHelper.Demo;

internal class Program
{
    private static void Main(string[] args)
    {
        Read();
    }

    private static void Write()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // ou LicenseContext.Commercial
        var file = @"/mnt/datas/dev/ExcelHelper/tests/ExcelHelper.Demo/Test_1.xlsx";

        using var stream = new FileStream(file, FileMode.OpenOrCreate);
        using var writer = new ExcelWriter(stream, new ExcelConfiguration());
        writer.Context.RegisterClassMap<PersonMap>();
        writer.Context.Configuration.SheetName = "People";

        writer.WriteRecords(new List<Person>
        {
            new() { FirstName = "John", LastName = "Doe", Age = 30 },
            new() { FirstName = "Jane", LastName = "Smith", Age = 25 }
        });
    }

    private static void Read()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // ou LicenseContext.Commercial
        var file = @"/mnt/datas/dev/ExcelHelper/tests/ExcelHelper.Demo/Test_1.xlsx";

        using var stream = new FileStream(file, FileMode.OpenOrCreate);
        using var reader = new ExcelReader(stream, new ExcelConfiguration());

        var records = reader.GetRecords<Person>().ToList();
    }
}

public class Person
{
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
    public int Age { get; set; }
}

public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.FirstName).Name("First Name");
        Map(m => m.LastName).Name("Last Name");
        Map(m => m.Age).Name("Age");
    }
}