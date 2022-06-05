namespace mobile.de_Scraper.Models
{
    public class InputModel
    {
        public Make Make { get; set; }
        public Model Model { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string MinPrice { get; set; }
        public string MaxPrice { get; set; }
        public string MinKilometers { get; set; }
        public string MaxKilometers { get; set; }
        public bool Vat { get; set; }
        public bool CompanySeller { get; set; }
        public bool AccidentCar { get;  set; }
        public string BatteryCapacityFrom { get;  set; }
        public string BatteryCapacityTo { get;  set; }
        public FuelType FuelType { get; set; }
        public Battery Battery { get; set; }
    }
}