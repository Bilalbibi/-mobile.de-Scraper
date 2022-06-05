using System.Collections.Generic;

namespace mobile.de_Scraper.Models
{
    public class Config
    {
        public List<Make> Makes { get; set; } = new List<Make>();
        public List<string> Prices { get; set; }
        public List<string> Dates { get; set; }
        public List<string> Millage { get; set; }
        public List<FuelType> Fuels { get; set; }
        public List<Battery> Batteries { get; set; }
    }
}