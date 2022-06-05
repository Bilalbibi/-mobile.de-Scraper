using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace scrapingTemplateV51.Models
{
    public class HttpCaller
    {
        HttpClient _httpClient;
        HttpClient _httpClient2;
        readonly HttpClientHandler _httpClientHandler = new HttpClientHandler()
        {
            //CookieContainer = new CookieContainer(),
            UseCookies = false,
            AllowAutoRedirect = true,
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        };
        public HttpCaller()
        {
            _httpClient = new HttpClient(_httpClientHandler);
            _httpClient.DefaultRequestHeaders.Add("User-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36");
            //_httpClient.DefaultRequestHeaders.Add("upgrade-insecure-requests", "1");
            _httpClient.DefaultRequestHeaders.Add("Accept-language", "en-US,en;q=0.9,fr;q=0.8,de;q=0.7");
            //_httpClient.DefaultRequestHeaders.Add("Host", "suchen.mobile.de");
            _httpClient.DefaultRequestHeaders.Add("Connection", "keep-alive");
            _httpClient.DefaultRequestHeaders.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9");


        }
        public async Task<HtmlDocument> GetDoc(string url, string host, string cookies, int maxAttempts = 1)
        {
            var html = await GetHtml(url, host, cookies, maxAttempts);
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            return doc;
        }
        public async Task<string> GetHtml(string url, string host, string cookies, int maxAttempts = 1)
        {
            int tries = 0;
            do
            {
                try
                {
                    var request = new HttpRequestMessage();
                    request.Headers.Add("Host", host);
                    if (cookies != null)
                    {
                        request.Headers.Add("Cookie", cookies);
                    }
                    request.Method = HttpMethod.Get;
                    request.RequestUri = new Uri(url);
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    var response = await _httpClient.SendAsync(request);
                    string html = await response.Content.ReadAsStringAsync();
                    return html;
                }
                catch (WebException ex)
                {
                    var errorMessage = "";
                    try
                    {
                        errorMessage = new StreamReader(ex.Response.GetResponseStream()).ReadToEnd();
                    }
                    catch (Exception)
                    {
                    }
                    tries++;
                    if (tries == maxAttempts)
                    {
                        throw new Exception(ex.Status + " " + ex.Message + " " + errorMessage);
                    }
                    await Task.Delay(2000);
                }
            } while (true);
        }
        public async Task<string> GetJsonModels(string url, int maxAttempts = 1)
        {
            int tries = 0;
            do
            {
                try
                {
                    var response = await _httpClient2.GetAsync(url);
                    CookieContainer cookies = new CookieContainer();
                    Uri uri = new Uri(url);
                    IEnumerable<Cookie> responseCookies = cookies.GetCookies(uri).Cast<Cookie>();
                    foreach (Cookie cookie in responseCookies)
                        Console.WriteLine(cookie.Name + "= " + cookie.Value + ";");
                    string html = await response.Content.ReadAsStringAsync();
                    return html;
                }
                catch (WebException ex)
                {
                    var errorMessage = "";
                    try
                    {
                        errorMessage = new StreamReader(ex.Response.GetResponseStream()).ReadToEnd();
                    }
                    catch (Exception)
                    {
                    }
                    tries++;
                    if (tries == maxAttempts)
                    {
                        throw new Exception(ex.Status + " " + ex.Message + " " + errorMessage);
                    }
                    await Task.Delay(2000);
                }
            } while (true);
        }

        public async Task<string> PostJson(string url, string json, int maxAttempts = 1)
        {
            int tries = 0;
            do
            {
                try
                {
                    var content = new StringContent(json, Encoding.UTF8, "application/json");
                    // content.Headers.Add("x-appeagle-authentication", Token);
                    var r = await _httpClient.PostAsync(url, content);
                    var s = await r.Content.ReadAsStringAsync();
                    return (s);
                }
                catch (WebException ex)
                {
                    var errorMessage = "";
                    try
                    {
                        errorMessage = new StreamReader(ex.Response.GetResponseStream()).ReadToEnd();
                    }
                    catch (Exception)
                    {
                    }
                    tries++;
                    if (tries == maxAttempts)
                    {
                        throw new Exception(ex.Status + " " + ex.Message + " " + errorMessage);
                    }
                    await Task.Delay(2000);
                }
            } while (true);

        }
        public async Task<string> PostFormData(string url, List<KeyValuePair<string, string>> formData, int maxAttempts = 1)
        {
            var formContent = new FormUrlEncodedContent(formData);
            int tries = 0;
            do
            {
                try
                {
                    var response = await _httpClient.PostAsync(url, formContent);
                    string html = await response.Content.ReadAsStringAsync();
                    return html;
                }
                catch (WebException ex)
                {
                    var errorMessage = "";
                    try
                    {
                        errorMessage = new StreamReader(ex.Response.GetResponseStream()).ReadToEnd();
                    }
                    catch (Exception)
                    {
                    }
                    tries++;
                    if (tries == maxAttempts)
                    {
                        throw new Exception(ex.Status + " " + ex.Message + " " + errorMessage);
                    }
                    await Task.Delay(2000);
                }
            } while (true);
        }
    }
}
