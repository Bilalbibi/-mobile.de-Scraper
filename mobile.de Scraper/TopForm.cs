using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using mobile.de_Scraper.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OpenQA.Selenium.Chrome;
using scrapingTemplateV51.Models;
using DevExpress.XtraGrid.Views.Grid;
using HtmlAgilityPack;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace mobile.de_Scraper
{
    public partial class TopForm : DevExpress.XtraEditors.XtraForm
    {
        private ChromeDriver _driver;
        private string cookies;
        public HttpCaller HttpCaller = new HttpCaller();
        private List<Car> cars = new List<Car>();
        private Config _config;
        public List<InputModel> inputModels = new List<InputModel>();
        public TopForm()
        {
            InitializeComponent();
        }
        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString(), @"Unhandled Thread Exception");
        }
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show((e.ExceptionObject as Exception)?.ToString(), @"Unhandled UI Exception");
        }
        public delegate void SetProgressD(int x);
        public void SetProgress(int x)
        {
            if (InvokeRequired)
            {
                Invoke(new SetProgressD(SetProgress), x);
                return;
            }
            if ((x <= 100))
            {
                progressB.EditValue = x;
            }
        }
        private async void TopForm_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists("outcomes"))
            {
                Directory.CreateDirectory("outcomes");
            }
            ServicePointManager.DefaultConnectionLimit = 65000;
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //await ScrapeConfigurations();
            //return;
            _config = JsonConvert.DeserializeObject<Config>(File.ReadAllText("conf"));
            MakesRepositoryItemLookUpEdit.DataSource = _config.Makes;
            DateFromRepositoryItemLookUpEdit.DataSource = _config.Dates;
            DateToRepositoryItemLookUpEdit.DataSource = _config.Dates;
            MinKmsRepositoryItemLookUpEdit.DataSource = _config.Millage;
            MaxKmsRepositoryItemLookUpEdit.DataSource = _config.Millage;
            MinPriceRepositoryItemLookUpEdit.DataSource = _config.Prices;
            MaxPriceRepositoryItemLookUpEdit.DataSource = _config.Prices;
            FuelTypeRepositoryItemLookUpEdit.DataSource = _config.Fuels;
            BatteryRepositoryItemLookUpEdit.DataSource = _config.Batteries;

            var d = JsonConvert.DeserializeObject<List<InputModel>>(File.ReadAllText("vv"));
            GridControle.DataSource = new BindingList<InputModel>(d);

        }
        private async Task ScrapeConfigurations()
        {
            _config = new Config();
            Invoke(new Action(() => Display("Loading www.mobile.de filters...")));

            var jsonMakes = await HttpCaller.GetHtml("https://www.mobile.de/svc/r/makes/Car?_lang=de", "www.mobile.de", null);
            var obj = JObject.Parse(jsonMakes);
            _config.Makes.Add(new Make { Name = "Any", Id = "1", Models = new List<Model>() });
            foreach (var makeNode in obj["makes"])
            {
                var id = (string)makeNode.SelectToken("i");
                if (id != "")
                {
                    var brand = (string)makeNode.SelectToken("n");
                    if (brand == "Any")
                    {
                        continue;
                    }
                    var models = await GetModels(id);
                    var make = new Make { Name = brand, Id = id, Models = models };

                    _config.Makes.Add(make);
                }
            }
            var doc = await HttpCaller.GetDoc("https://www.mobile.de/?vc=Car&dam=0&mk=&sfmr=false&cn=&ml=&p=", "www.mobile.de", null);
            var pricesJson = doc.DocumentNode.SelectSingleNode("//script[contains(text(),'{price: \"\",prices:')]").InnerText;
            var x = pricesJson.IndexOf("{price:", StringComparison.Ordinal);
            var xx = pricesJson.IndexOf(",openOnPageload: ", StringComparison.Ordinal);

            pricesJson = pricesJson.Substring(x, xx - x).Replace("value", "\"value\"").Replace("label", "\"label\"")
                .Replace("prices", "quotes").Replace("quotes", "\"quotes\"").Replace("price", "\"price\"") + "}";

            obj = JObject.Parse(pricesJson);
            var prices = new List<string>();
            foreach (var element in obj["quotes"])
            {
                prices.Add((string)element.SelectToken("label") ?? "");
            }

            var datesElements = doc.DocumentNode.SelectNodes("//input[@id='qsfrg']/following-sibling::select/option").ToList();
            var dates = new List<string>();
            foreach (var datesElement in datesElements)
            {
                dates.Add(datesElement.InnerText);
            }


            var KilometersElements = doc.DocumentNode.SelectNodes("//input[@id='qsmil']/following-sibling::select/option").ToList();
            var millage = new List<string>();
            foreach (var KilometersElement in KilometersElements)
            {
                millage.Add(KilometersElement.InnerText);
            }

            var txtFuelTypes = File.ReadAllLines("Fuel types.txt").ToList();
            var fuels = new List<FuelType>();
            foreach (var s in txtFuelTypes)
            {
                var array = s.Split(',');
                fuels.Add(new FuelType() { Code = array[1], Name = array[0] });
            }
            var txtBatteries = File.ReadAllLines("Batteries.txt").ToList();
            var batteries = new List<Battery>();
            foreach (var s in txtBatteries)
            {
                var array = s.Split(',');
                batteries.Add(new Battery { Code = array[1], Name = array[0] });
            }

            _config.Prices = prices;
            _config.Millage = millage;
            _config.Fuels = fuels;
            _config.Dates = dates;
            _config.Batteries = batteries;
            File.WriteAllText("conf", JsonConvert.SerializeObject(_config, Formatting.Indented));

            Invoke(new Action(() => Display("")));
            Invoke(new Action(() => MessageBox.Show(@"filters loaded, you can select the required filters")));
        }
        private async Task<List<Model>> GetModels(string id)
        {
            var models = new List<Model>();
            models.Add(new Model { Name = "Any", Id = "" });
            var resJson = await HttpCaller.GetHtml($@"https://m.mobile.de/svc/r/models/{id}?_jsonp=_loadModels&_lang=en", "m.mobile.de", null);
            var json = resJson.Replace("_loadModels(", "").Replace(");", "");
            var obj = JObject.Parse(json);
            foreach (var o in obj["models"])
            {
                var model = new Model();
                model.Name = (string)o.SelectToken("n");
                model.Id = (string)o.SelectToken("i");
                models.Add(model);
            }

            return models;
        }
        public delegate void DisplayD(string s);
        public void Display(string s)
        {
            if (InvokeRequired)
            {
                Invoke(new DisplayD(Display), s);
                return;
            }
            DispalyArea.Text = s;
        }
        private async Task MainWork()
        {
            foreach (var inputModel in inputModels)
            {
                await StartScraping(inputModel);
            }
            await SaveData();
        }
        private async Task StartScraping(InputModel inputModel)
        {
            var damagedCar = "";
            if (!inputModel.AccidentCar)
            {
                damagedCar = "NO_DAMAGE_UNREPAIRED";
            }

            var commercial = "";
            if (inputModel.CompanySeller)
            {
                commercial = "ONLY_COMMERCIAL_FSBO_ADS";
            }

            var makeID = "";
            var modelID = "";
            if (!inputModel.Make.Name.Equals("Any"))
            {
                makeID = inputModel.Make.Id;
                if (inputModel.Model != null)
                {
                    modelID = inputModel.Model.Id;
                }
            }
            var fromDate = "";
            if (inputModel.FromDate != null)
            {
                if (!inputModel.FromDate.Equals("Any"))
                {
                    fromDate = inputModel.FromDate + "-01-01";
                }
            }
            var toDate = "";
            if (inputModel.ToDate != null)
            {
                if (!inputModel.ToDate.Equals("Any"))
                {
                    toDate = inputModel.ToDate + "-12-31";
                }
            }
            var url = "";
            var doc = new HtmlDocument();

            do
            {
                url = $@"https://suchen.mobile.de/fahrzeuge/search.html?adLimitation={commercial}&damageUnrepaired={damagedCar}&grossPrice=true&isSearchRequest=true&makeModelVariant1.makeId={makeID}&maxFirstRegistrationDate={toDate}&maxPrice={inputModel.MaxPrice}&minFirstRegistrationDate={fromDate}&minPrice={inputModel.MinPrice}&pageNumber=1&scopeId=C&sfmr=false&vatable={inputModel.Vat.ToString().ToLower()}&lang=en&maxMileage={inputModel.MaxKilometers}&minMileage={inputModel.MinKilometers}&fuels={inputModel.FuelType.Code}&makeModelVariant1.modelId={modelID}&bat={inputModel.Battery.Code}&minBatteryCapacity={inputModel.BatteryCapacityFrom}&maxBatteryCapacity={inputModel.BatteryCapacityTo}";
                doc = await HttpCaller.GetDoc(url, "suchen.mobile.de", cookies);
                doc.Save("test.html");
                if (!doc.DocumentNode.OuterHtml.Contains("matching your search criteria"))
                {
                    await Task.Run((() => GetCookies(url)));
                    continue;
                }
                break;
            } while (true);

            //doc.Save("look.html");
            if (doc.DocumentNode.OuterHtml.Contains("No ads found"))
            {
                Display("no result");
                return;
            }
            int pages = 1;
            var pagesText = doc.DocumentNode
                .SelectSingleNode("//li[@class='pref-next u-valign-sub']/following-sibling::li[last()]/span")?.InnerText
                .Trim() ?? doc.DocumentNode
                .SelectSingleNode("//ul[@class='pagination']/li[last()-1]/span")?.InnerText
                .Trim();
            if (pagesText != null)
            {

            }

            try
            {
                pages = int.Parse(pagesText ?? string.Empty);
            }
            catch (Exception)
            {
                // ignored
            }

            var result = doc.DocumentNode.SelectSingleNode("//button[@id='minisearch-search-btn']//span[@class='hit-counter']")?.InnerText.Trim().Replace(",", "");
            if (result == "0")
            {
                return;
            }
            var carNbr = int.Parse(result);
            int counter = 1;
            for (int i = 1; i <= pages; i++)
            {
                url = $@"https://suchen.mobile.de/fahrzeuge/search.html?adLimitation={commercial}&damageUnrepaired={damagedCar}&grossPrice=true&isSearchRequest=true&makeModelVariant1.makeId={makeID}&maxFirstRegistrationDate={toDate}&maxPrice={inputModel.MaxPrice}&minFirstRegistrationDate={fromDate}&minPrice={inputModel.MinPrice}&pageNumber={i}&scopeId=C&sfmr=false&vatable={inputModel.Vat.ToString().ToLower()}&lang=en&maxMileage={inputModel.MaxKilometers}&minMileage={inputModel.MinKilometers}&fuels={inputModel.FuelType.Code}&makeModelVariant1.modelId={modelID}&bat={inputModel.Battery.Code}&minBatteryCapacity={inputModel.BatteryCapacityFrom}&maxBatteryCapacity={inputModel.BatteryCapacityTo}";
                doc = await HttpCaller.GetDoc(url, "suchen.mobile.de", cookies);
                //doc.Save("ggg.html");
                var products = doc.DocumentNode.SelectNodes("//div[contains(@class,'cBox-body cBox-body')]/a")?.ToList();
                if (products == null)
                {
                    await Task.Run((() => GetCookies(url)));
                    doc = await HttpCaller.GetDoc(url, "suchen.mobile.de", cookies);
                    products = doc.DocumentNode.SelectNodes("//div[contains(@class,'cBox-body cBox-body')]/a")?.ToList();
                }
                foreach (var product in products)
                {
                    if (product.InnerText.Contains("Sponsored"))
                        continue;
                    var link = product.GetAttributeValue("href", "");
                    var car = await ScrapeCarDetails(link);
                    cars.Add(car);
                    Display($@"{counter} car scraped/{carNbr}");
                    SetProgress(((counter * 100) / (carNbr)));
                    counter++;
                }

            }
        }
        private async Task<Car> ScrapeCarDetails(string url)
        {
            var car = new Car();
            var doc = await HttpCaller.GetDoc(url + "&lang=en", "suchen.mobile.de", cookies);
            doc.Save("example car.html");
            var imagesContainer = doc.DocumentNode.SelectSingleNode("//script[contains(text(),'if (\"undefined\" !')]").InnerText.Trim();
            car.Url = WebUtility.HtmlDecode(url);

            var title = doc.DocumentNode.SelectSingleNode("//h1").InnerText.Trim();
            var makes = _config.Makes.Select(x => x.Name).ToList();

            var words = title.Split(' ');
            car.Make = words[0];
            car.Model = title.Replace(car.Make, "").Replace("-", "");
            var img = doc.DocumentNode?.SelectSingleNode("//div[@class='image-carousel']//img")?.GetAttributeValue("src", "") ?? "N/A";
            if (img == "N/A")
            {
                car.Image = img;
            }
            else
            {
                car.Image = "https:" + img;
            }
            car.Year = doc.DocumentNode.SelectSingleNode("//div[@id='rbt-firstRegistration-v']")?.InnerText.Trim() ?? "N/A";
            car.Kilometre = WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//div[@id='rbt-mileage-v']")?.InnerText.Trim() ?? "N/A");
            car.Price = WebUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//span[@class='h3 rbt-prime-price']")?.InnerText.Trim().Replace("(Gross)", "") ?? "N/A");
            var phone = doc.DocumentNode.SelectSingleNode("//p[@id='rbt-db-phone']")?.InnerText.Trim();
            if (phone != null)
            {
                var array = phone.Split(':');
                car.Phone = array[1].Replace("Phone", "");
            }
            else
                car.Phone = "N/A";
            //cars.Add(car);
            return car;
        }
        async Task GetCookies(string url)
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            ChromeOptions options = new ChromeOptions();
            _driver = new ChromeDriver(service, options);
            _driver.Manage().Window.Position = new Point(-32000, -32000);
            _driver.Navigate().GoToUrl("https://www.mobile.de/?lang=en");
            await Task.Delay(2000);
            _driver.FindElementByXPath("//button[text()='Accept']").Click();
            await Task.Delay(2000);
            _driver.Navigate().GoToUrl(url);
            await Task.Delay(2000);
            _driver.Navigate().Refresh();

            try
            {
                _driver.FindElementById("agree-all-top").Click();
            }
            catch (Exception)
            {
                //
            }
            await Task.Delay(3000);
            var cookieBiulde = new StringBuilder();

            foreach (var cookie in _driver.Manage().Cookies.AllCookies)
            {
                cookieBiulde.Append(cookie.Name + "=" + cookie.Value + ";");
            }
            _driver.Quit();
            cookies = cookieBiulde.ToString()
                .Substring(0, cookieBiulde.ToString().LastIndexOf(";", StringComparison.Ordinal));
        }
        private async void simpleButton1_Click(object sender, EventArgs e)
        {


            //File.WriteAllText("vv", JsonConvert.SerializeObject(inputModels, Formatting.Indented));
            //return;
            //var car = await ScrapeCarDetails("https://suchen.mobile.de/fahrzeuge/details.html?id=312785408&amp;damageUnrepaired=ALSO_DAMAGE_UNREPAIRED&amp;fuels=ELECTRICITY&amp;grossPrice=true&amp;isSearchRequest=true&amp;makeModelVariant1.makeId=135&amp;makeModelVariant1.modelId=5&amp;maxFirstRegistrationDate=2019-12-31&amp;maxPrice=43000&amp;minFirstRegistrationDate=1900-01-01&amp;page
            //ber=1&amp;scopeId=C&amp;sfmr=false&amp;vatable=false&amp;searchId=0f23ac3c-256a-067f-7a2f-3cb8a7a894be");
            //await SaveData();
            //return;
            //var doc1 = await HttpCaller.GetDoc("https://suchen.mobile.de/fahrzeuge/search.html?adLimitation=ONLY_COMMERCIAL_FSBO_ADS&damageUnrepaired=NO_DAMAGE_UNREPAIRED&grossPrice=true&isSearchRequest=true&makeModelVariant1.makeId=20700&maxFirstRegistrationDate=2021-12-31&maxPrice=22500&minFirstRegistrationDate=1900-01-01&minPrice=1000&pageNumber=2&scopeId=C&sfmr=false&vatable=true&lang=en", "suchen.mobile.de", "_fbp=fb.1.1614964268766.2137666824; bm_sz=B3D81E4EA5EFFB96DFDE9909E3B3E1E4~YAAQtv4SAuyCU8l3AQAA3h1fAwuunFCGdmfFeQzFrV8wxYZM6jPXBxsZd+O9jhY/kbguqzQ6Yof3ynoQSxQRJDwzRXINFwprW8VYewfP7o7IvEgXTnltHJMYcqUKAExKcZ8e5KZDiPRbmbR2MuRPOkxjFS4tYw3amrin/2g/a/D0E04kNQF7HYV8+yUTQ4M=; _ga=GA1.2.761541693.1614964269; _gid=GA1.2.1551951768.1614964269; optimizelyEndUserId=oeu1614964269652r0.5111431893185285; _gat=1; visited=1; ak_bmsc=EBF9E88D620B55BD8467ADFFF3ADD23A~000000000000000000000000000000~YAAQtv4SAgODU8l3AQAA7ipfAwsyk9Yp92rmmqep6IZTVVUYE+16zEfl90yRbN0Gxgo3NXh6Mf8lAsTfUYuy4HHbQ+0t4F/FWIjiGEb1d8LXTbhpEHfsvjtiIiO6cbMB2QG66WMPvZjt0WoTre+IoKaMfQixETgQLYi19etQgZWCeAVQDbjsymROWgEw7yK4spM/0gjS3wfrFOBRGZjI600pjBluJ9R0wbVoe1RUbuCEM3eby/GPUB4av/gKpQ+ebhyz3FFV2WEmtxavqZAtGfIv0SEuJlz8pZ6oOHhnEc/UQDbsueTRq1V6cFLmBZd4T6+zOtqwwV4epqooHUUkDGFssW3h80jtomJROPzwuAGnSCcAwA35UW8Vs4ZyjKKsjklA9WFuV0Gz0gwaiKrtLO4N1KRYHUapvcNxLiVL+aP5s8AxCH3/NhJsE50tl8xW7wmCXT/EQYISXg+SsN/HWanU/TWi0cFpkDbemow=; vi=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjaWQiOiJmYjI2OGFlNS1jNDg5LTRkYTYtYjc1NS00YjQ3MjllYmY1NjgiLCJpYXQiOjE2MTQ5NjQyNzN9.1f3Utn6eryADWIRiH-gukqBcMZGijnEGi5hbqIiYGJg; mdeConsentDataGoogle=1; mdeConsentData=CPCl_4APCl_4AEyADADEBOCgAP_AAH_AAAYgHbNf_X_fb39j-_59__t0eY1f9_7_v-0zjhfds-8Nyf_X_L8X_2M7vF36pq4KuR4ku3bBIQFtHOnUTUmx6olVrTPsak2Mr7NKJ7Lkmnsbe2dYGHtfn91T-ZKZr_7v___7________79______3_v_____-_____9_8Dtmv_r_vt7-x_f8-__26PMav-_9_3_aZxwvu2feG5P_r_l-L_7Gd3i79U1cFXI8SXbtgkIC2jnTqJqTY9USq1pn2NSbGV9mlE9lyTT2NvbOsDD2vz-6p_MlM1_93___9________9-______7_3______f____-_-AA.YAAAAAAADwAAACYAAAAA; __gads=ID=116c0ca508d2a1ae:T=1614964275:S=ALNI_Ma6quGLZ1uym3LGP38tQd0tn4Bdmg; mobile.LOCALE=en; _abck=82E2666B572AA7C1AFA9AFDB6D163C2D~0~YAAQdfASAmWBg+13AQAAbYhfAwUysAEU9aqAI6xi35UiDRiEex2gjAPZ0H1C1ZriwpZcZS2/JjjKbASy2hVkh3R9x5hMxr2XfAW3rEZhqBleU5zvHPy51kxivnmLfJCcYBiP4cHaQ/CoYVM31Q54o5JqKPhW+4SJ6Lng+1S8RuXe5kSmm0nV2Aun13Fm96fJo1EE23GQLQhhipDtE3q46s85hw5R+1rvRw+XGbMO7sJaHsdrtqqfJeZySrJ2tyOwdLAQQMcu3CkteBWxZEjbWlsIy5ayF/42dDUoBpALWWvtU31F/zNhKJ0VVE/Fo1H1KhYGk2ZRATI8GSEyYuk4ci8kU5JHy6ONMJBLvgzsBgbhaCK/1Twdt8bwRa/VBlDVjY7ZkaSgVeFDsUqEBQ7j6tGWGKn/Cfc=~-1~-1~-1; lux_uid=161496429708120063; ioam2018=001e6a5aefa90243560424f79:1644513097130:1614964297130:.mobile.de:2:mobile:DE/EN/OB/S/P/S/G:noevent:1614964297130:i9pk7l; iom_consent=0103ff03ff&1614964297453; _uetsid=9ff29e207dd011eba1993b69c317be11; _uetvid=9ff2e6207dd011eb8041ad12394d2798; axd=4253787041664625994; tis=EP117%3A2735; bm_sv=4F34E96C22B8C90C1394C7EFAA1BB9F8~XDskopcsii59Ts8IPukAkufzcn0sSxTBWfxLxWmU7L+bzwXOvHT5b9pzXP27XMy8Z7md0ylaUhqA8naojb2gedi7fIhwvLQ7R26VPLTDyNj0J/4V2IhoALh5Z4PS1AzzwhMVdQEVf6AaugF+kmmQVXSVyqeo1jNF3cxgLz+KnZc=; RT=\"z=1&dm=mobile.de&si=8re7ngfezhb&ss=klwk0jto&sl=h&tt=z07&obo=b&rl=1&ld=1llq&r=2f9cf29d2bd97c2622192103d66ee255&ul=1llv\"");
            //doc1.Save("try.html");
            if (!Daily.Checked && !ThreeDays.Checked)
            {
                MessageBox.Show(@"Please select time base scraping ""Daily"" or ""3 days"" option ");
                return;
            }

            do
            {
                inputModels = new List<InputModel>();
                cars = new List<Car>();
                for (int i = 0; i < FiltersDGV.RowCount; i++)
                {
                    var input = FiltersDGV.GetRow(i) as InputModel;
                    inputModels.Add(input);
                }

                var d2 = new DateTime();
                var days = 1;
                if (Daily.Checked)
                {
                    d2 = DateTime.Now.AddDays(1);
                }
                if (ThreeDays.Checked)
                {
                    days = 3;
                    d2 = DateTime.Now.AddDays(3);
                }
                await MainWork();
                var d1 = DateTime.Now;
                Display($@"work done for today next run will be {DateTime.Now.AddDays(days):dd/MM/yyyy} ");
                await Task.Delay(d2 - d1);
            } while (true);
        }
        private async Task SaveData()
        {
            var date = DateTime.Now.ToString("dd_MM_yyyy hh_mm_ss");
            var path = $@"outcomes\mobile.de{date}.xlsx";
            var excelPkg = new ExcelPackage(new FileInfo(path));

            var sheet = excelPkg.Workbook.Worksheets.Add("Cars");
            sheet.Protection.IsProtected = false;
            sheet.Protection.AllowSelectLockedCells = false;
            sheet.Row(1).Height = 20;
            sheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Row(1).Style.Font.Bold = true;
            sheet.Row(1).Style.Font.Size = 8;
            sheet.Cells[1, 1].Value = "Make";
            sheet.Cells[1, 2].Value = "Model";
            sheet.Cells[1, 3].Value = "Price";
            sheet.Cells[1, 4].Value = "Phone";
            sheet.Cells[1, 5].Value = "Year";
            sheet.Cells[1, 6].Value = "Kilometers";
            sheet.Cells[1, 7].Value = "Weblink";

            var range = sheet.Cells[$"A1:G{cars.Count + 1}"];
            var tab = sheet.Tables.Add(range, "");

            tab.TableStyle = TableStyles.Medium2;
            sheet.Cells.Style.Font.Size = 12;

            var row = 2;
            foreach (var car in cars)
            {

                sheet.Cells[row, 1].Value = car.Make;
                sheet.Cells[row, 2].Value = car.Model;
                sheet.Cells[row, 3].Value = car.Price;
                sheet.Cells[row, 4].Value = car.Phone;
                sheet.Cells[row, 5].Value = car.Year;
                sheet.Cells[row, 6].Value = car.Kilometre;
                sheet.Cells[row, 7].Value = car.Url;
                row++;
            }

            for (int i = 2; i <= sheet.Dimension.End.Column; i++)
                sheet.Column(i).AutoFit();

            sheet.View.FreezePanes(2, 1);
            await excelPkg.SaveAsync();

        }
        private void FiltersDGV_CustomRowCellEdit(object sender, CustomRowCellEditEventArgs e)
        {

            if (e.Column.FieldName.Equals("Model"))
            {

                var inputModel = FiltersDGV.GetRow(e.RowHandle) as InputModel;
                if (inputModel == null) return;
                if (inputModel.Make == null)
                {
                    var r2 = new RepositoryItemLookUpEdit();
                    e.RepositoryItem = r2;
                    e.Column.ColumnEdit = r2;
                    return;
                }
                var r = new RepositoryItemLookUpEdit();
                r.DataSource = inputModel.Make.Models;
                r.Columns.Add(new LookUpColumnInfo { FieldName = "Name" });
                r.DisplayMember = "Name";
                GridControle.RepositoryItems.Add(r);
                e.RepositoryItem = r;
                e.Column.ColumnEdit = r;
            }
        }
        private void MakesrepositoryItemLookUpEdit_EditValueChanged(object sender, EventArgs e)
        {
            var inputModel = FiltersDGV.GetRow(FiltersDGV.FocusedRowHandle) as InputModel;

        }

        private void TopForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _driver?.Quit();
            inputModels = new List<InputModel>();
            for (int i = 0; i < FiltersDGV.RowCount; i++)
            {
                var input = FiltersDGV.GetRow(i) as InputModel;
                inputModels.Add(input);
            }
            File.WriteAllText("vv", JsonConvert.SerializeObject(inputModels, Formatting.Indented));
        }
    }
}