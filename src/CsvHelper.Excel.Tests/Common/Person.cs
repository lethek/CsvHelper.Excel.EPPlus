namespace CsvHelper.Excel.Tests
{
    public record Person
    {
        public int? Id { get; init; }
        public string Name { get; init; }
        public int Age { get; init; }
        public string Empty { get; init; }
    }
}
