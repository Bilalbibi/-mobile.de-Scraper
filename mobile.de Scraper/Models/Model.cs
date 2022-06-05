namespace mobile.de_Scraper.Models
{
    public class Model
    {
        public string Name { get; set; }
        public string Id { get; set; }
        public override string ToString()
        {
            return Name;
        }

        public override bool Equals(object obj)
        {
            return Id.Equals(((Model)obj).Id);
        }
    }
}
