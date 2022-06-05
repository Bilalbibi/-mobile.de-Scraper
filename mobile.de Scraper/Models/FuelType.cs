namespace mobile.de_Scraper.Models
{
    public class FuelType
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public override bool Equals(object obj)
        {
            return Code.Equals(((FuelType)obj).Code);
        }
    }
}