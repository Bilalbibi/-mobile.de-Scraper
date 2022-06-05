using MetroFramework.Controls;
using MetroFramework.Forms;
using Newtonsoft.Json;
using scrapingTemplateV51.Models;
using System;
using System.Collections.Generic;
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
using DevExpress.XtraGrid.Views.Grid;
using mobile.de_Scraper.Models;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OpenQA.Selenium.Chrome;

namespace scrapingTemplateV51
{
    public partial class MainForm : MetroForm
    {
        public bool LogToUi = true;
        public bool LogToFile = true;

        private readonly string _path = Application.StartupPath;
        private Dictionary<string, string> _uiConfig;
        public HttpCaller HttpCaller = new HttpCaller();
        private ChromeDriver _driver;
        private Config _config;
        private List<Car> cars = new List<Car>();
        private string cookies;
        public MainForm()
        {
            InitializeComponent();
        }


        private async Task MainWork()
        {
            await StartScraping();
            await SaveData();
        }

        private async Task StartScraping()
        {

            await GetCookies();

            //var make = MarkBoxI.GetItemText(MarkBoxI.SelectedItem);

            //var selectedMake = makesAndModels.Find(x => x.Make.Name == make);

            //make = selectedMake.Make.Id;

            //var model = ModelBoxI.GetItemText(ModelBoxI.SelectedItem);
            //var modelObj = selectedMake.Models.First(mo => mo.Name == model);
            //model = modelObj.Id;



            //var minDate = DatesMinBoxI.GetItemText(DatesMinBoxI.SelectedItem);
            //var maxDate = DateMaxBoxI.GetItemText(DateMaxBoxI.SelectedItem);
            //if (minDate == "Any")
            //{
            //    minDate = "";
            //}
            //if (maxDate == "Any")
            //{
            //    maxDate = "";
            //}
            //var kmsMin = KilometersMinBoxI.GetItemText(KilometersMinBoxI.SelectedItem).Replace("km", "").Replace(",", "").Trim().Replace(".", ""); ;
            //var kmsMax = KilometersMaxBoxI.GetItemText(KilometersMaxBoxI.SelectedItem).Replace("km", "").Replace(",", "").Trim().Replace(".", ""); ;
            //if (kmsMin == "Any")
            //{
            //    kmsMin = "";
            //}
            //if (kmsMax == "Any")
            //{
            //    kmsMax = "";
            //}
            //var maxPrice = MaxPricesBoxI.GetItemText(MaxPricesBoxI.SelectedItem).Replace("€", "").Replace(",", "").Trim().Replace(".","");
            //var minPrice = MinPriceBoxI.GetItemText(MinPriceBoxI.SelectedItem).Replace("€", "").Replace(",", "").Trim().Replace(".", ""); ;
            //if (maxPrice == "Any")
            //{
            //    maxPrice = "";
            //}

            //var damagedCar = "";
            //if (AccidentNo.Checked)
            //{
            //    damagedCar = "NO_DAMAGE_UNREPAIRED";
            //}

            //var commercial = "";
            //if (CompanyYes.Checked)
            //{
            //    commercial = "ONLY_COMMERCIAL_FSBO_ADS";
            //}
            //string vat;
            //if (VATYES.Checked)
            //    vat = "true";
            //else
            //    vat = "false";
            //var fuelType = fuelTypes[FuelTypesBoxI.GetItemText(FuelTypesBoxI.SelectedItem)];
            //var url =
            //    $@"https://suchen.mobile.de/fahrzeuge/search.html?adLimitation={commercial}&damageUnrepaired={damagedCar}&grossPrice=true&isSearchRequest=true&makeModelVariant1.makeId={make}&maxFirstRegistrationDate={maxDate}-12-31&maxPrice={maxPrice}&minFirstRegistrationDate={minDate}-01-01&minPrice={minPrice}&pageNumber=1&scopeId=C&sfmr=false&vatable={vat}&lang=en&maxMileage={kmsMax}&minMileage={kmsMin}&fuels={fuelType}&makeModelVariant1.modelId={model
            //    }";
            //var doc = await HttpCaller.GetDoc(url, "suchen.mobile.de", cookies);

            //doc.Save("look.html");
            //if (doc.DocumentNode.OuterHtml.Contains("No ads found"))
            //{
            //    Display("no result");
            //    return;
            //}
            //int pages = 1;
            //var pagesText = doc.DocumentNode
            //    .SelectSingleNode("//li[@class='pref-next u-valign-sub']/following-sibling::li[last()]/span")?.InnerText
            //    .Trim();
            //if (pagesText!=null)
            //{

            //}

            //try
            //{
            //    pages = int.Parse(pagesText ?? string.Empty);
            //}
            //catch (Exception)
            //{
            //    // ignored
            //}

            //var result = doc.DocumentNode.SelectSingleNode("//button[@id='minisearch-search-btn']//span[@class='hit-counter']").InnerText.Trim().Replace(",", "");
            //var carNbr = int.Parse(result);
            //int counter = 1;
            //for (int i = 1; i <= pages; i++)
            //{
            //    doc = await HttpCaller.GetDoc(url, "suchen.mobile.de", cookies);

            //    var products = doc.DocumentNode.SelectNodes("//div[contains(@class,'cBox-body cBox-body')]/a").ToList();

            //    foreach (var product in products)
            //    {
            //        var link = product.GetAttributeValue("href","");
            //        var car = await ScrapeCarDetails(link);
            //        cars.Add(car);
            //        Display($@"{counter} car scraped/{carNbr}");
            //        SetProgress(((counter * 100) / carNbr));
            //        counter++;
            //    }

            //}
        }

        private async Task GetCookies()
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            ChromeOptions options = new ChromeOptions();
            //options.AddArgument("headless");
            _driver = new ChromeDriver(service, options);
            _driver.Navigate().GoToUrl("https://www.mobile.de/?lang=en");
            await Task.Delay(3000);
            _driver.FindElementById("gdpr-consent-accept-button").Click();
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

        private async Task<Car> ScrapeCarDetails(string url)
        {
            var car = new Car();
            var doc = await HttpCaller.GetDoc(url + "&lang=en", "suchen.mobile.de", null);
            car.Url = url;
            var Vat = doc.DocumentNode.SelectSingleNode("//span[@class='rbt-sec-price']")?.InnerText.Trim();
            if (Vat != null)
                car.VAT = "Yes";
            else
                car.VAT = "No";
            var title = doc.DocumentNode.SelectSingleNode("//h1").InnerText.Trim();
            var makes = _config.Makes.Select(x => x.Name).ToList();
            foreach (var make in makes)
            {
                if (title.Contains(make))
                {
                    car.Make = make;
                    car.Model = title.Replace(make, "").Replace("-", "");
                    break;
                }
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

            return car;
        }
        private async void Form1_Load(object sender, EventArgs e)
        {
            ServicePointManager.DefaultConnectionLimit = 65000;
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            LoadConfig();
             await ScrapeConfigurations();
            _config = JsonConvert.DeserializeObject<Config>(File.ReadAllText("conf"));
            MakesRepositoryItemLookUpEdit.DataSource = _config.Makes;
            DateFromRepositoryItemLookUpEdit.DataSource = _config.Dates;
            DateToRepositoryItemLookUpEdit.DataSource = _config.Dates;
            MinKmsRepositoryItemLookUpEdit.DataSource = _config.Millage;
            MaxKmsRepositoryItemLookUpEdit.DataSource = _config.Millage;
            MinPriceRepositoryItemLookUpEdit.DataSource = _config.Prices;
            MaxPriceRepositoryItemLookUpEdit.DataSource = _config.Prices;
            FuelTypeRepositoryItemLookUpEdit.DataSource = _config.Fuels;
            
            var d = JsonConvert.DeserializeObject<List<InputModel>>(File.ReadAllText("vv"));
            inputModelBindingSource.DataSource = d;
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

            _config.Prices = prices;
            _config.Millage = millage;
            _config.Fuels = fuels;
            _config.Dates = dates;
            File.WriteAllText("conf", JsonConvert.SerializeObject(_config, Formatting.Indented));

            Invoke(new Action(() => Display("")));
            Invoke(new Action(() => MessageBox.Show(@"filters loaded, you can select the required filters")));
        }

        private async Task<List<Model>> GetModels(string id)
        {
            var models = new List<Model>();

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

        void InitControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    try
                    {
                        if (x.Name.EndsWith("I"))
                        {
                            switch (x)
                            {
                                case MetroCheckBox _:
                                case CheckBox _:
                                    ((CheckBox)x).Checked = bool.Parse(_uiConfig[((CheckBox)x).Name]);
                                    break;
                                case RadioButton radioButton:
                                    radioButton.Checked = bool.Parse(_uiConfig[radioButton.Name]);
                                    break;
                                case TextBox _:
                                case RichTextBox _:
                                case MetroTextBox _:
                                    x.Text = _uiConfig[x.Name];
                                    break;
                                case NumericUpDown numericUpDown:
                                    numericUpDown.Value = int.Parse(_uiConfig[numericUpDown.Name]);
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    InitControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        public void SaveControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    #region Add key value to disctionarry

                    if (x.Name.EndsWith("I"))
                    {
                        switch (x)
                        {
                            case MetroCheckBox _:
                            case RadioButton _:
                            case CheckBox _:
                                _uiConfig.Add(x.Name, ((CheckBox)x).Checked + "");
                                break;
                            case TextBox _:
                            case RichTextBox _:
                            case MetroTextBox _:
                                _uiConfig.Add(x.Name, x.Text);
                                break;
                            case NumericUpDown _:
                                _uiConfig.Add(x.Name, ((NumericUpDown)x).Value + "");
                                break;
                            default:
                                Console.WriteLine(@"could not find a type for " + x.Name);
                                break;
                        }
                    }
                    #endregion
                    SaveControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        private void SaveConfig()
        {
            _uiConfig = new Dictionary<string, string>();
            SaveControls(this);
            try
            {
                File.WriteAllText("config.txt", JsonConvert.SerializeObject(_uiConfig, Formatting.Indented));
            }
            catch (Exception e)
            {
                ErrorLog(e.ToString());
            }
        }
        private void LoadConfig()
        {
            try
            {
                _uiConfig = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText("config.txt"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }
            InitControls(this);
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString(), @"Unhandled Thread Exception");
        }
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show((e.ExceptionObject as Exception)?.ToString(), @"Unhandled UI Exception");
        }
        #region UIFunctions
        public delegate void WriteToLogD(string s, Color c);
        public void WriteToLog(string s, Color c)
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new WriteToLogD(WriteToLog), s, c);
                    return;
                }
                if (LogToUi)
                {
                    if (DebugT.Lines.Length > 5000)
                    {
                        DebugT.Text = "";
                    }
                    DebugT.SelectionStart = DebugT.Text.Length;
                    DebugT.SelectionColor = c;
                    DebugT.AppendText(DateTime.Now.ToString(Utility.SimpleDateFormat) + " : " + s + Environment.NewLine);
                }
                Console.WriteLine(DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s);
                if (LogToFile)
                {
                    File.AppendAllText(_path + "/data/log.txt", DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        public void NormalLog(string s)
        {
            WriteToLog(s, Color.Black);
        }
        public void ErrorLog(string s)
        {
            WriteToLog(s, Color.Red);
        }
        public void SuccessLog(string s)
        {
            WriteToLog(s, Color.Green);
        }
        public void CommandLog(string s)
        {
            WriteToLog(s, Color.Blue);
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
                ProgressB.Value = x;
            }
        }
        public delegate void DisplayD(string s);
        public void Display(string s)
        {
            if (InvokeRequired)
            {
                Invoke(new DisplayD(Display), s);
                return;
            }
            displayT.Text = s;
        }

        #endregion
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _driver.Quit();
            SaveConfig();
        }
        private void loadInputB_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog { Filter = @"TXT|*.txt", InitialDirectory = _path };
            if (o.ShowDialog() == DialogResult.OK)
            {
            }
        }
        private void openInputB_Click_1(object sender, EventArgs e)
        {
            try
            {
            }
            catch (Exception ex)
            {
                ErrorLog(ex.ToString());
            }
        }
        private void openOutputB_Click_1(object sender, EventArgs e)
        {
            try
            {
            }
            catch (Exception ex)
            {
                ErrorLog(ex.ToString());
            }
        }
        private void loadOutputB_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Filter = @"csv file|*.csv",
                Title = @"Select the output location"
            };
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
            }
        }

        private async void startB_Click_1(object sender, EventArgs e)
        {
            List<InputModel> inputModels = new List<InputModel>();
            for (int i = 0; i < FiltersDGV.RowCount; i++)
            {
                var input = FiltersDGV.GetRow(i) as InputModel;
                inputModels.Add(input);
            }
            File.WriteAllText("vv", JsonConvert.SerializeObject(inputModels, Formatting.Indented));
            return;
            //var doc1 = await HttpCaller.GetDoc("https://suchen.mobile.de/fahrzeuge/search.html?adLimitation=ONLY_COMMERCIAL_FSBO_ADS&damageUnrepaired=NO_DAMAGE_UNREPAIRED&grossPrice=true&isSearchRequest=true&makeModelVariant1.makeId=20700&maxFirstRegistrationDate=2021-12-31&maxPrice=22500&minFirstRegistrationDate=1900-01-01&minPrice=1000&pageNumber=2&scopeId=C&sfmr=false&vatable=true&lang=en", "suchen.mobile.de", "_fbp=fb.1.1614964268766.2137666824; bm_sz=B3D81E4EA5EFFB96DFDE9909E3B3E1E4~YAAQtv4SAuyCU8l3AQAA3h1fAwuunFCGdmfFeQzFrV8wxYZM6jPXBxsZd+O9jhY/kbguqzQ6Yof3ynoQSxQRJDwzRXINFwprW8VYewfP7o7IvEgXTnltHJMYcqUKAExKcZ8e5KZDiPRbmbR2MuRPOkxjFS4tYw3amrin/2g/a/D0E04kNQF7HYV8+yUTQ4M=; _ga=GA1.2.761541693.1614964269; _gid=GA1.2.1551951768.1614964269; optimizelyEndUserId=oeu1614964269652r0.5111431893185285; _gat=1; visited=1; ak_bmsc=EBF9E88D620B55BD8467ADFFF3ADD23A~000000000000000000000000000000~YAAQtv4SAgODU8l3AQAA7ipfAwsyk9Yp92rmmqep6IZTVVUYE+16zEfl90yRbN0Gxgo3NXh6Mf8lAsTfUYuy4HHbQ+0t4F/FWIjiGEb1d8LXTbhpEHfsvjtiIiO6cbMB2QG66WMPvZjt0WoTre+IoKaMfQixETgQLYi19etQgZWCeAVQDbjsymROWgEw7yK4spM/0gjS3wfrFOBRGZjI600pjBluJ9R0wbVoe1RUbuCEM3eby/GPUB4av/gKpQ+ebhyz3FFV2WEmtxavqZAtGfIv0SEuJlz8pZ6oOHhnEc/UQDbsueTRq1V6cFLmBZd4T6+zOtqwwV4epqooHUUkDGFssW3h80jtomJROPzwuAGnSCcAwA35UW8Vs4ZyjKKsjklA9WFuV0Gz0gwaiKrtLO4N1KRYHUapvcNxLiVL+aP5s8AxCH3/NhJsE50tl8xW7wmCXT/EQYISXg+SsN/HWanU/TWi0cFpkDbemow=; vi=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjaWQiOiJmYjI2OGFlNS1jNDg5LTRkYTYtYjc1NS00YjQ3MjllYmY1NjgiLCJpYXQiOjE2MTQ5NjQyNzN9.1f3Utn6eryADWIRiH-gukqBcMZGijnEGi5hbqIiYGJg; mdeConsentDataGoogle=1; mdeConsentData=CPCl_4APCl_4AEyADADEBOCgAP_AAH_AAAYgHbNf_X_fb39j-_59__t0eY1f9_7_v-0zjhfds-8Nyf_X_L8X_2M7vF36pq4KuR4ku3bBIQFtHOnUTUmx6olVrTPsak2Mr7NKJ7Lkmnsbe2dYGHtfn91T-ZKZr_7v___7________79______3_v_____-_____9_8Dtmv_r_vt7-x_f8-__26PMav-_9_3_aZxwvu2feG5P_r_l-L_7Gd3i79U1cFXI8SXbtgkIC2jnTqJqTY9USq1pn2NSbGV9mlE9lyTT2NvbOsDD2vz-6p_MlM1_93___9________9-______7_3______f____-_-AA.YAAAAAAADwAAACYAAAAA; __gads=ID=116c0ca508d2a1ae:T=1614964275:S=ALNI_Ma6quGLZ1uym3LGP38tQd0tn4Bdmg; mobile.LOCALE=en; _abck=82E2666B572AA7C1AFA9AFDB6D163C2D~0~YAAQdfASAmWBg+13AQAAbYhfAwUysAEU9aqAI6xi35UiDRiEex2gjAPZ0H1C1ZriwpZcZS2/JjjKbASy2hVkh3R9x5hMxr2XfAW3rEZhqBleU5zvHPy51kxivnmLfJCcYBiP4cHaQ/CoYVM31Q54o5JqKPhW+4SJ6Lng+1S8RuXe5kSmm0nV2Aun13Fm96fJo1EE23GQLQhhipDtE3q46s85hw5R+1rvRw+XGbMO7sJaHsdrtqqfJeZySrJ2tyOwdLAQQMcu3CkteBWxZEjbWlsIy5ayF/42dDUoBpALWWvtU31F/zNhKJ0VVE/Fo1H1KhYGk2ZRATI8GSEyYuk4ci8kU5JHy6ONMJBLvgzsBgbhaCK/1Twdt8bwRa/VBlDVjY7ZkaSgVeFDsUqEBQ7j6tGWGKn/Cfc=~-1~-1~-1; lux_uid=161496429708120063; ioam2018=001e6a5aefa90243560424f79:1644513097130:1614964297130:.mobile.de:2:mobile:DE/EN/OB/S/P/S/G:noevent:1614964297130:i9pk7l; iom_consent=0103ff03ff&1614964297453; _uetsid=9ff29e207dd011eba1993b69c317be11; _uetvid=9ff2e6207dd011eb8041ad12394d2798; axd=4253787041664625994; tis=EP117%3A2735; bm_sv=4F34E96C22B8C90C1394C7EFAA1BB9F8~XDskopcsii59Ts8IPukAkufzcn0sSxTBWfxLxWmU7L+bzwXOvHT5b9pzXP27XMy8Z7md0ylaUhqA8naojb2gedi7fIhwvLQ7R26VPLTDyNj0J/4V2IhoALh5Z4PS1AzzwhMVdQEVf6AaugF+kmmQVXSVyqeo1jNF3cxgLz+KnZc=; RT=\"z=1&dm=mobile.de&si=8re7ngfezhb&ss=klwk0jto&sl=h&tt=z07&obo=b&rl=1&ld=1llq&r=2f9cf29d2bd97c2622192103d66ee255&ul=1llv\"");
            //doc1.Save("try.html");
            //return;
            if (!Daily.Checked && !ThreeDays.Checked)
            {
                MessageBox.Show(@"Please select time base scraping ""Daily"" or ""3 days"" option ");
                return;
            }


            do
            {
                var time = 0;
                var days = 0;
                if (Daily.Checked)
                {
                    days = 1;
                    time = (1000 * 60 * 60) * days;
                }
                if (ThreeDays.Checked)
                {
                    days = 3;
                    time = (1000 * 60 * 60) * days;
                }
                await MainWork();
                Display($@"work done for today next run will be {DateTime.Now.AddDays(days):dd/MM/yyyy} ");
                await Task.Delay(time);
            } while (true);

        }
        private async Task SaveData()
        {
            var date = DateTime.Now.ToString("dd_MM_yyyy");
            var path = $@"mobile.de{date}.xlsx";
            var excelPkg = new ExcelPackage(new FileInfo(path));

            var sheet = excelPkg.Workbook.Worksheets.Add("Cars");
            sheet.Protection.IsProtected = false;
            sheet.Protection.AllowSelectLockedCells = false;
            sheet.Row(1).Height = 20;
            sheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Row(1).Style.Font.Bold = true;
            sheet.Row(1).Style.Font.Size = 8;
            sheet.Cells[1, 1].Value = "Weblink";
            sheet.Cells[1, 2].Value = "Phone";
            sheet.Cells[1, 3].Value = "Make";
            sheet.Cells[1, 4].Value = "Model";
            sheet.Cells[1, 5].Value = "Vat";
            sheet.Cells[1, 6].Value = "Mileage";
            sheet.Cells[1, 7].Value = "Year";
            sheet.Cells[1, 8].Value = "Price";

            var range = sheet.Cells[$"A1:H{cars.Count + 1}"];
            var tab = sheet.Tables.Add(range, "");

            tab.TableStyle = TableStyles.Medium2;
            sheet.Cells.Style.Font.Size = 12;

            var row = 2;
            foreach (var car in cars)
            {

                sheet.Cells[row, 1].Value = car.Url;
                sheet.Cells[row, 2].Value = car.Phone;
                sheet.Cells[row, 3].Value = car.Make;
                sheet.Cells[row, 4].Value = car.Model;
                sheet.Cells[row, 5].Value = car.VAT;
                sheet.Cells[row, 6].Value = car.Kilometre;
                sheet.Cells[row, 7].Value = car.Year;
                sheet.Cells[row, 8].Value = car.Price;
                row++;
            }

            for (int i = 2; i <= sheet.Dimension.End.Column; i++)
                sheet.Column(i).AutoFit();

            sheet.View.FreezePanes(2, 1);
            await excelPkg.SaveAsync();

        }

        private void MarkBox_SelectedValueChanged(object sender, EventArgs e)
        {
            //ModelBoxI.Enabled = true;
            //ModelBoxI.Items.Clear();
            //var makeName = MarkBoxI.GetItemText(MarkBoxI.SelectedItem);
            //var makeModels = makesAndModels.Find(x => x.Make.Name == makeName);

            //foreach (var model in makeModels.Models)
            //{
            //    ModelBoxI.Items.Add(model.Name);
            //}
        }

        private void MakesrepositoryItemLookUpEdit_EditValueChanged(object sender, EventArgs e)
        {
            var inputModel = FiltersDGV.GetRow(FiltersDGV.FocusedRowHandle) as InputModel;

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
                //r.Properties.DataBindings.Add("EditValue",FiltersDGV.DataSource, "Model", false, DataSourceUpdateMode.OnPropertyChanged);
                e.RepositoryItem = r;
                e.Column.ColumnEdit = r;
            }
        }

        private void FiltersDGV_InitNewRow(object sender, DevExpress.XtraGrid.Views.Grid.InitNewRowEventArgs e)
        {
            //GridView view = sender as GridView;
            //view.SetRowCellValue(e.RowHandle, view.Columns[2], _config.Dates.First());
            //view.SetRowCellValue(e.RowHandle, view.Columns[3], _config.Dates.First());
            //view.SetRowCellValue(e.RowHandle, view.Columns[4], _config.Prices.First());
            //view.SetRowCellValue(e.RowHandle, view.Columns[5], _config.Prices.First());
            //view.SetRowCellValue(e.RowHandle, view.Columns[6], _config.Millage.First());
            //view.SetRowCellValue(e.RowHandle, view.Columns[7], _config.Millage.First());
            //view.SetRowCellValue(e.RowHandle, view.Columns[8], _config.Fuels.First());
        }
    }
}
