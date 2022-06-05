namespace mobile.de_Scraper.Models
{
    public class Battery
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public override bool Equals(object obj)
        {
            return Code.Equals(((Battery)obj).Code);
        }
    }
}