namespace CsvHelper.Excel.Tests
{
    public class Person
    {
        public int? Id { get; set; }

        public string Name { get; set; }
        
        public int Age { get; set; }

        public string Empty { get; set; }


        public override int GetHashCode()
        {
            int hash = 17;
            hash = hash * 23 + Id?.GetHashCode() ?? 0;
            hash = hash * 23 + Name?.GetHashCode() ?? 0;
            hash = hash * 23 + Age.GetHashCode();
            hash = hash * 23 + Empty?.GetHashCode() ?? 0;
            return hash;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Person)) return false;
            var other = (Person)obj;
            if (Id != other.Id) return false;
            if (Name != other.Name) return false;
            if (Age != other.Age) return false;
            if ((Empty ?? "") != (other.Empty ?? "")) return false;
            return true;
        }
    }
}