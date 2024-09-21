using System.CodeDom.Compiler;
using System.Diagnostics;
using System.Globalization;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools;
using OpenQA.Selenium.Interactions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Magic.BrowserAutomationNET
{
    public class Chrome
    {
        public string InstanceId { get; set; }
        public bool Headless { get; set; } = false;
        public bool MaximizeWindow { get; set; } = true;
        public bool ShowImages { get; set; } = true;
        public string DefaultDownloadDirectory { get; set; } = string.Empty;
        public string WindowSize { get; set; } = "0,0";
        public string UserAgent { get; set; } = string.Empty;
        public bool DisableExtension { get; set; } = false;
        public bool DisableApplicationCache { get; set; } = false;
        public bool Kiosk { get; set; } = false;
        public string Profile { get; set; } = string.Empty;
        public string DriverDirectory { get; set; }
        public string BinaryLocation { get; set; }
        public bool SaveResources { get; set; } = true;

        public IWebDriver? Driver { get; set; }

        public ChromeDriverService? ChromeDriverService { get; set; }
        public ChromeOptions ChromeOptions { get; set; } = new ChromeOptions();

        public Chrome(string binaryLocation = "", string driverDirectory = "")
        {

            this.BinaryLocation = binaryLocation;
            this.DriverDirectory = driverDirectory;
            this.InstanceId = HelperNET.CreateRandomAlphabeticLowerCaseString(5);

        } // end of method

        public void SetChromeOptions()
        {
            if (BinaryLocation != "")
            {
                ChromeOptions.BinaryLocation = BinaryLocation;
            }

            ChromeOptions.AddArgument("--disable-blink-features=AutomationControlled");
            ChromeOptions.AddExcludedArgument("enable-automation");
            ChromeOptions.AddAdditionalOption("useAutomationExtension", false);
            //ChromeOptions.AddExcludedArgument("test-type");

            if (MaximizeWindow)
            {
                ChromeOptions.AddArgument("start-maximized");
            }

            if (Headless)
            {
                ChromeOptions.AddArgument("headless");
            }

            if (!ShowImages)
            {
                ChromeOptions.AddArgument("--blink-settings=imagesEnabled=false");
            }

            if (WindowSize != "0,0")
            {
                ChromeOptions.AddArgument($"--window-size={WindowSize}");
            }

            if (UserAgent != string.Empty)
            {
                ChromeOptions.AddArgument($"--user-agent={UserAgent}");
            }

            if (DisableExtension)
            {
                ChromeOptions.AddArgument($"--disable-extensions");
            }

            ChromeOptions.AddArgument("--disable-notifications"); // to disable notification
            ChromeOptions.AddArgument("--no-sandbox");
            ChromeOptions.AddArgument("--disable-infobars");
            ChromeOptions.AddArgument("--noerrordialogs");

            if (DisableApplicationCache)
            {
                ChromeOptions.AddArgument("--disable-application-cache"); // to disable cache
            }

            if (Kiosk)
            {
                ChromeOptions.AddArgument("--kiosk");
            }

            // ChromeOptions.AddArgument("--disable-web-security");

            if (Profile != string.Empty)
            {
                string path = Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%");

                ChromeOptions.AddArgument($@"--user-data-dir={path}/Google/Chrome/User Data");
                ChromeOptions.AddArgument($"--profile-directory={Profile}");
            }

            ChromeOptions.AddArgument("--allow-running-insecure-content");

            // ini dari chrome asli
            //ChromeOptions.AddArgument("--flag-switches-begin");
            //ChromeOptions.AddArgument("--flag-switches-end");
            //ChromeOptions.AddArgument("--origin-trial-disabled-features=SecurePaymentConfirmation");

            ChromeOptions.UnhandledPromptBehavior = UnhandledPromptBehavior.Dismiss;

            ChromeOptions.AddUserProfilePreference("profile.default_content_setting_values.notifications", 2);
            ChromeOptions.AddUserProfilePreference("credentials_enable_service", false);

            if (DefaultDownloadDirectory != string.Empty)
            {
                ChromeOptions.AddUserProfilePreference("download.default_directory", DefaultDownloadDirectory);
                //ChromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
                //ChromeOptions.AddUserProfilePreference("download.directory_upgrade", true);
                ChromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");

                ChromeOptions.AddArgument("--browser.download.folderList=2");
                ChromeOptions.AddArgument("--browser.helperApps.neverAsk.saveToDisk=audio/mp3");
                ChromeOptions.AddArgument($"--browser.download.dir={DefaultDownloadDirectory}");
            }

            if (DriverDirectory == "")
            {
                ChromeDriverService = ChromeDriverService.CreateDefaultService();
            }
            else
            {
                ChromeDriverService = ChromeDriverService.CreateDefaultService(DriverDirectory);
            }

            ChromeDriverService.HideCommandPromptWindow = true;


            // dari chat GPT
            if (SaveResources)
            {
                ChromeOptions.AddArgument("--disable-gpu");
                //ChromeOptions.AddArgument("--disable-extensions");
                ChromeOptions.AddArgument("--disable-plugins");
                ChromeOptions.AddArgument("--disable-background-networking");
                ChromeOptions.AddArgument("--disable-software-rasterizer");
                ChromeOptions.AddArgument("--disable-rendering-local-only");
                ChromeOptions.AddArgument("--disable-sync");
                ChromeOptions.AddArgument("--safebrowsing-disable-auto-update");
                ChromeOptions.AddArgument("--metrics-recording-only");
                ChromeOptions.AddArgument("--disable-media-source");
                ChromeOptions.AddUserProfilePreference("disable-prefetch", 1);
            }

            /*
            ChromeOptions.AddArgument("--autoplay-policy=no-user-gesture-required");
            ChromeOptions.AddArgument("--disable-features=InfiniteSessionRestore");
            ChromeOptions.AddArgument("--ignore-certificate-errors");
            */

        } // end of method

        public static void DeleteProfile(string profileName)
        {
            string profilePath = Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%") + $"/Google/Chrome/User Data/{profileName}";

            if (Directory.Exists(profilePath))
            {
                Directory.Delete(profilePath, true);
            }
        } // end of method

        public Browser OpenBrowser()
        {
            SetChromeOptions();

            Browser browser = new Browser();

            try
            {
                Driver = new ChromeDriver(ChromeDriverService, ChromeOptions);

                IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)Driver;

                // 1. Hapus navigator.webdriver
                jsExecutor.ExecuteScript("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});");

                // 3. Modifikasi navigator.plugins
                jsExecutor.ExecuteScript(@"
                    Object.defineProperty(navigator, 'plugins', {
                        get: () => [1, 2, 3]
                    });
                ");

                // 4. Modifikasi navigator.languages
                jsExecutor.ExecuteScript(@"
                    Object.defineProperty(navigator, 'languages', {
                        get: () => ['en-US', 'en']
                    });
                ");

                // 5. Modifikasi screen.availHeight dan screen.availWidth
                jsExecutor.ExecuteScript(@"
                    Object.defineProperty(screen, 'availHeight', { get: () => 900 });
                    Object.defineProperty(screen, 'availWidth', { get: () => 1440 });
                ");

                //Driver = new ChromeDriver(ChromeOptions);
                browser.State = true;
                browser.Message = "Open browser successfully.";
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("This version of ChromeDriver only supports Chrome version"))
                {
                    string browserVersionEtc = ex.Message.Split(new string[] { "urrent browser version is " }, StringSplitOptions.None)[1];
                    int browserVersion = Convert.ToInt32(browserVersionEtc.Split('.')[0]);

                    string driverVersionEtc = ex.Message.Split(new string[] { "only supports Chrome version " }, StringSplitOptions.None)[1];
                    int driverVersion = Convert.ToInt32(driverVersionEtc.Split('\n')[0]);

                    if (browserVersion < driverVersion)
                    {
                        browser.Message = "Browser Chrome Anda perlu diupdate untuk bisa menjalankan Marketplace Magic";
                    }
                    else
                    {
                        browser.Message = "Open browser failed. Exception : " + ex.Message;
                    }
                }
                else
                {
                    browser.Message = "Open browser failed. Exception : " + ex.Message;
                }

                browser.State = false;
            }

            return browser;

        } // end of method

        public void CloseBrowser()
        {

            try
            {
                if (Driver != null)
                {
                    Driver.Close();
                    Driver.Quit();
                    Driver = null;
                }
            }
            catch (Exception ex)
            {
                // Log the exception or handle it as needed
                Debug.WriteLine($"Error occurred while closing the browser: {ex.Message}");
            }

        } // end of method

        public WebPage Navigate(string Url)
        {
            WebPage returnResults = new WebPage();

            try
            {
                if (Driver != null && Driver.GetType() == typeof(ChromeDriver))
                {
                    Driver.Navigate().GoToUrl(Url);
                }

                returnResults.State = true;
                returnResults.Message = "Navigate successfully.";
                returnResults.Url = Url;
            }
            catch (Exception ex)
            {
                returnResults.State = false;
                returnResults.Message = "Navigate failed. Exception : " + ex.Message;
                returnResults.Url = string.Empty;

                Console.WriteLine(ex.Message);
            }

            return returnResults;
        } // end of function

        public WebPage Refresh()
        {
            WebPage returnResults = new WebPage();

            try
            {
                if (Driver != null && Driver.GetType() == typeof(ChromeDriver))
                {
                    Driver.Navigate().Refresh();

                    returnResults.State = true;
                    returnResults.Url = GetCurrentUrl().Url;
                    returnResults.Message = "Refresh successfully.";
                }
            }
            catch (Exception ex)
            {
                returnResults.State = false;
                returnResults.Url = string.Empty;
                returnResults.Message = "Refresh failed. Exception : " + ex.Message;
            }

            return returnResults;
        } // end of function

        public WebPage GetCurrentUrl()
        {
            WebPage returnResults = new WebPage();

            if(Driver == null)
            {
                returnResults.State = false;
                returnResults.Message = "Get URL failed.";

                return returnResults;
            }

            string url;

            try
            {
                url = Driver.Url;

                returnResults.State = true;
                returnResults.Message = "Get URL successfully.";
                returnResults.Url = url;
            }
            catch (Exception ex)
            {
                returnResults.State = false;
                returnResults.Message = "Get URL failed. Exception : " + ex.Message;
            }

            return returnResults;
        } // end of method

        public WebElement FindElementByXPath(string XPathSelector, int timeSpan = 10)
        {
            Debug.WriteLine("Akan select : " + XPathSelector);

            WebElement webElement = new WebElement();

            if (Driver == null)
            {
                webElement.State = false;
                webElement.Message = "Driver is null.";

                return webElement;
            }

            try
            {
                Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(timeSpan);

                webElement.Item = Driver.FindElement(By.XPath(XPathSelector));
                webElement.State = true;
                webElement.Message = "Element is found.";
            }
            catch (Exception ex)
            {
                webElement.State = false;
                webElement.Message = "Element is NOT found. Exception: " + ex;
            }
            finally
            {
                try
                {
                    Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(0);
                }
                catch (Exception ex)
                {
                    webElement.State = false;
                    webElement.Message = "Clear implicit wait failed. Exception: " + ex;
                }
            }

            webElement.Selector = XPathSelector;
            webElement.Driver = Driver;

            return webElement;
        } // end of method

        public WebElements FindElementsByXPath(string XPathSelector, int timeSpan = 10)
        {
            WebElements result = new WebElements();

            if (Driver == null)
            {
                result.State = false;
                result.Message = "No element is found.";
                result.Selector = XPathSelector;
                result.Count = 0;
                result.Items = new List<WebElement>();
                result.Driver = Driver;

                return result;
            }

            try
            {
                Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(timeSpan);
            }
            catch (Exception ex)
            {
                result.State = false;
                result.Message = "Clear implicit wait failed. Exception : " + ex.Message;
                result.Selector = XPathSelector;
                result.Count = 0;
                result.Items = new List<WebElement>();
                result.Driver = Driver;

                return result;
            }

            IReadOnlyCollection<IWebElement>? webElements_ = null;

            try
            {
                webElements_ = Driver.FindElements(By.XPath(XPathSelector));
                int i = 1;

                foreach (IWebElement webElement in webElements_)
                {
                    WebElement singleResult = new WebElement();

                    singleResult.State = true;
                    singleResult.Message = "Element is found.";
                    singleResult.Selector = "(" + XPathSelector + ")[" + i + "]";
                    singleResult.Item = webElement;
                    singleResult.Driver = Driver;

                    result.Items.Add(singleResult);

                    i++;
                }

                result.State = true;
                result.Message = "Elements are found.";
                result.Selector = XPathSelector;
                result.Count = i - 1;
                result.Driver = Driver;

            }
            catch
            {
                result.State = false;
                result.Message = "No element is found.";
                result.Selector = XPathSelector;
                result.Count = 0;
                result.Items = new List<WebElement>();
                result.Driver = Driver;
            }

            try
            {
                Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(0);
            }
            catch (Exception ex)
            {
                result.State = false;
                result.Message = "Clear implicit wait failed. Exception : " + ex.Message;
                result.Selector = XPathSelector;
                result.Count = 0;
                result.Items = new List<WebElement>();
                result.Driver = Driver;

                return result;
            }

            if (result.Items.Count <= 0)
            {
                result.State = false;
                result.Message = "No element is found.";
                result.Selector = XPathSelector;
                result.Items = new List<WebElement>();
                result.Driver = Driver;
            }

            return result;

        } // end of method


        public int CountElementsByXPath(string XPathSelector, int timeSpan = 10)
        {
            WebElements result = new WebElements();

            if (Driver == null)
            {
                return 0;
            }

            try
            {
                Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(timeSpan);
            }
            catch
            {
                return -1;
            }

            IReadOnlyCollection<IWebElement>? webElements_ = null;

            try
            {
                webElements_ = Driver.FindElements(By.XPath(XPathSelector));

            }
            catch
            {
            }

            try
            {
                Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(0);
            }
            catch
            {
                return -1;
            }

            if (webElements_ == null)
                return 0;
            else 
                return webElements_.Count;

        } // end of method


        public void SwitchToIframe(int frameIndex = 0)
        {
            
            if (Driver == null) return;

            Driver.SwitchTo().Frame(frameIndex);

        } // end of method

        public void SwitchToIframe(WebElement iframe)
        {

            if (Driver == null) return;

            Driver.SwitchTo().Frame(iframe.Item);

        } // end of method

        public void SwitchToIframe(IWebElement iframe)
        {

            if (Driver == null) return;

            Driver.SwitchTo().Frame(iframe);

        } // end of method

        public void IframeSwitchBack()
        {

            if (Driver == null) return;

            Driver.SwitchTo().ParentFrame();

        } // end of method

        public void DeleteAllCookies()
        {

            if (Driver == null) return;
            Driver.Manage().Cookies.DeleteAllCookies();

        } // end of method

        public void DeleteAllResources()
        {

            if (Driver == null) return;

            // Hapus semua cookies
            Driver.Manage().Cookies.DeleteAllCookies();

            // Hapus cache
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver!;
            js.ExecuteScript("window.localStorage.clear();");
            js.ExecuteScript("window.sessionStorage.clear();");

        } // end of method

        public CookiesData GetAllCookies()
        {
            CookiesData cookies = new CookiesData();

            if (Driver == null)
            {
                cookies.Status = false;
                cookies.Message = "Get All Cookies failed. Chrome driver was null.";

                return cookies;
            }

            try
            {
                cookies.Items = Driver.Manage().Cookies.AllCookies.ToList();
                cookies.Status = true;
                cookies.Amount = cookies.Items.Count();
                cookies.Message = "Get All Cookies done successfully.";
            }
            catch (WebDriverException ex)
            {
                Driver.Quit();
                cookies.Status = false;
                cookies.Message = "Error retrieving cookies. Exception : " + ex.Message + ".";
                cookies.ExceptionName = ex.GetType().Name;
            }

            return cookies;

        } // end of method

        public string? GetAllCookiesToJson(double addDays = 0)
        {
            CookiesData cookiesData = GetAllCookies();

            if (!cookiesData.Status)
            {
                return null;
            }

            List<CookieInStrings> cookieInStrings = new List<CookieInStrings>();

            foreach (OpenQA.Selenium.Cookie cookie in cookiesData.Items)
            {
                DateTime myExpiry = cookie.Expiry ?? DateTime.Now;
                myExpiry = addDays == 0 ? myExpiry : myExpiry.AddDays(addDays);

                CookieInStrings cookieItem = new CookieInStrings();
                cookieItem.Name = cookie.Name;
                cookieItem.Value = cookie.Value;
                cookieItem.Domain = cookie.Domain;
                cookieItem.Path = cookie.Path;
                cookieItem.Expiry = myExpiry.ToString("M/d/yyyy h:m:s tt", new CultureInfo("en-US"));

                cookieInStrings.Add(cookieItem);
            }

            return JsonConvert.SerializeObject(cookieInStrings);
        } // end of method

        public CookiesFile SaveCookiesToFile(string filename, bool append = false, double addDays = 0)
        {
            CookiesFile cookiesFile = new CookiesFile();

            if (!append)
            {
                if (File.Exists(@filename))
                {
                    File.Delete(@filename);
                }
            }

            cookiesFile.CookiesData = GetAllCookies();

            if (!cookiesFile.CookiesData.Status)
            {
                cookiesFile.Status = false;
                cookiesFile.Message = "Save cookies failed.";
                cookiesFile.filename = null!;

                return cookiesFile;
            }

            string OneLineCookieData;

            foreach (OpenQA.Selenium.Cookie cookie in cookiesFile.CookiesData.Items)
            {

                DateTime myExpiry = cookie.Expiry ?? DateTime.Now;
                myExpiry = addDays == 0 ? myExpiry : myExpiry.AddDays(addDays);

                string myExpiryUS = myExpiry.ToString("M/d/yyyy h:m:s tt", new CultureInfo("en-US"));

                OneLineCookieData = cookie.Name + ";" + cookie.Value + ";" + cookie.Domain + ";" + cookie.Path + ";" + myExpiryUS;

                File.AppendAllText(@filename, OneLineCookieData + Environment.NewLine);
            }

            cookiesFile.Status = true;
            cookiesFile.Message = "Cookies saved succesfully";
            cookiesFile.filename = filename;

            return cookiesFile;
        } // end of method

        public CookiesFile SaveRawCookiesToFile(string filename, bool append = false)
        {
            CookiesFile cookiesFile = new CookiesFile();

            if (!append)
            {
                if (File.Exists(@filename))
                {
                    File.Delete(@filename);
                }
            }

            cookiesFile.CookiesData = GetAllCookies();

            if (!cookiesFile.CookiesData.Status)
            {
                cookiesFile.Status = false;
                cookiesFile.Message = "Get All Cookies failed.";
                cookiesFile.filename = null!;

                return cookiesFile;
            }

            string OneLineCookieData = "";

            for (int i = 0; i < cookiesFile.CookiesData.Items.Count(); i++)
            {
                OneLineCookieData = OneLineCookieData + cookiesFile.CookiesData.Items.ElementAt(i).Name + "=" + cookiesFile.CookiesData.Items.ElementAt(i).Value;

                if (i < cookiesFile.CookiesData.Items.Count() - 1)
                {
                    OneLineCookieData = OneLineCookieData + "; ";
                }
            }

            File.AppendAllText(@filename, OneLineCookieData);

            cookiesFile.Status = true;
            cookiesFile.Message = "Cookies saved succesfully";
            cookiesFile.filename = filename;

            return cookiesFile;
        } // end of method

        public CookieData AddCookie(OpenQA.Selenium.Cookie cookie)
        {
            CookieData cookieData = new CookieData();
            cookieData.Cookie = cookie;

            try
            {
                Driver!.Manage().Cookies.AddCookie(cookie);

                cookieData.Status = true;
                cookieData.Message = "Cookie added succesfully";
            }
            catch (Exception ex)
            {
                cookieData.Status = false;
                cookieData.Message = "Failed when adding cookie. Exception : " + ex.Message;
            }

            return cookieData;
        } // end of method

        public CookieData AddCookie(string name, string value, string domain, string path, DateTime? expiry)
        {
            CookieData cookieData = new CookieData();

            try
            {
                cookieData.Cookie = new OpenQA.Selenium.Cookie(name, value, domain, path, expiry);
                Driver!.Manage().Cookies.AddCookie(cookieData.Cookie);

                cookieData.Status = true;
                cookieData.Message = "Cookie added succesfully";
            }
            catch (Exception ex)
            {
                cookieData.Status = false;
                cookieData.Cookie = null;
                cookieData.Message = "Failed when adding cookie. Exception : " + ex.Message;
            }

            return cookieData;
        } // end of method

        public CookiesFile AddCookiesFromFile(string filename)
        {
            CookiesFile cookiesFile = new CookiesFile();

            string[] lines = File.ReadAllLines(@filename);
            string[] OneCookieData;
            CookieDetails cookieDetails = new CookieDetails();

            int cookieAdded = 0;

            foreach (string line in lines)
            {
                OneCookieData = line.Split(';');

                cookieDetails.Name = OneCookieData[0];
                cookieDetails.Value = OneCookieData[1];
                cookieDetails.Domain = OneCookieData[2];
                cookieDetails.Path = OneCookieData[3];
                cookieDetails.Expiry = OneCookieData[4];

                if (cookieDetails.Expiry == "")
                {
                    try
                    {
                        CookieData cookieData = AddCookie(cookieDetails.Name, cookieDetails.Value, cookieDetails.Domain, cookieDetails.Path, null);
                        cookiesFile.CookiesData!.Items.Add(cookieData.Cookie!);
                        cookieAdded++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                else
                {
                    DateTime expiry = new DateTime();

                    try
                    {
                        expiry = DateTime.ParseExact(cookieDetails.Expiry, "M/d/yyyy h:m:s tt", new CultureInfo("en-US"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        continue;
                    }

                    try
                    {
                        CookieData cookieData = AddCookie(cookieDetails.Name, cookieDetails.Value, cookieDetails.Domain, cookieDetails.Path, expiry);
                        cookiesFile.CookiesData!.Items.Add(cookieData.Cookie!);
                        cookieAdded++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }
            } // end of foreach

            cookiesFile.Status = true;
            cookiesFile.CookiesData!.Amount = cookieAdded;
            cookiesFile.filename = filename;
            cookiesFile.Message = $"{cookieAdded} cookies added.";

            return cookiesFile;
        } // end of method

        public void AddCookiesFromJson(string jsonString)
        {
            List<CookieInStrings> cookieInStrings = JsonConvert.DeserializeObject<List<CookieInStrings>>(jsonString)!;

            foreach (CookieInStrings cookieItem in cookieInStrings!)
            {
                if (cookieItem.Expiry == "")
                {
                    try
                    {
                        CookieData cookieData = AddCookie(cookieItem.Name, cookieItem.Value, cookieItem.Domain, cookieItem.Path, null);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                else
                {
                    DateTime expiry = new DateTime();

                    try
                    {
                        expiry = DateTime.ParseExact(cookieItem.Expiry, "M/d/yyyy h:m:s tt", new CultureInfo("en-US"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        continue;
                    }

                    try
                    {
                        CookieData cookieData = AddCookie(cookieItem.Name, cookieItem.Value, cookieItem.Domain, cookieItem.Path, expiry);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }
            }

        } // end of method


        public long ScrollToBottom(int wait)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver!;
            long scrollHeight = (long)js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight); return document.body.scrollHeight;");

            Thread.Sleep(wait);

            return scrollHeight;
        } // end of method

        public ScrollResult Scroll(int Height)
        {
            ScrollResult scrollResult = new ScrollResult();

            ((IJavaScriptExecutor)Driver!).ExecuteScript($"window.scrollBy(0, {Height})");

            scrollResult.Status = true;
            scrollResult.Message = "Scroll done successfully.";
            scrollResult.Height = Height;

            return scrollResult;
        } // end of method

        public void SendKeys(string key, string keydown = null!)
        {
            Actions actions = new Actions(Driver);
            try
            {
                if (keydown != null)
                {
                    actions.KeyDown(keydown);
                }

                actions.SendKeys(key);

                if (keydown != null)
                {
                    actions.KeyUp(keydown);
                }

                actions.Build().Perform();
            }
            catch
            {

            }
        } // end of method

        public class Version
        {
            public int BrowserMajorVersion { get; set; } = 0;
            public int DriverMajorVersion { get; set; } = 0;

            public string BrowserFullVersion { get; set; } = null!;
            public string DriverFullVersion { get; set; } = null!;

            public string DriverDownloadURL { get; set; } = null!;

            private HttpClient _httpClient = null!;

            public Version()
            {

            } // end of method

            public async Task CurrentVersion(string chromedriverDirectory = null!)
            {
                await Task.Run(() =>
                {
                    if (!System.IO.File.Exists("chromedriver.exe"))
                    {
                        this.BrowserFullVersion = null!;
                        this.DriverFullVersion = null!;

                        this.BrowserMajorVersion = 0;
                        this.DriverMajorVersion = 0;

                        return;
                    }

                    Chrome chrome = new Chrome();

                    chrome.DriverDirectory = chromedriverDirectory;
                    chrome.Headless = true;
                    chrome.SetChromeOptions();

                    try
                    {
                        chrome.Driver = new ChromeDriver(chrome.ChromeDriverService, chrome.ChromeOptions);
                        //chrome.Driver = new ChromeDriver(chrome.ChromeOptions);

                        ICapabilities capabilities = ((ChromeDriver)chrome.Driver).Capabilities;

                        this.BrowserFullVersion = capabilities["browserVersion"].ToString()!;
                        //Console.WriteLine("BrowserFullVersion (in CurrentVersion Method) : " + this.BrowserFullVersion);

                        this.DriverFullVersion = (capabilities["chrome"] as Dictionary<string, object>)!["chromedriverVersion"].ToString()!.Split(' ')[0];

                        this.BrowserMajorVersion = Convert.ToInt32(this.BrowserFullVersion!.Split('.')[0]);
                        //Console.WriteLine("BrowserMajorVersion (in CurrentVersion Method) : " + this.BrowserMajorVersion);

                        this.DriverMajorVersion = Convert.ToInt32(this.DriverFullVersion.Split('.')[0]);


                    }
                    catch (Exception ex)
                    {
                        //session not created: This version of ChromeDriver only supports Chrome version 102
                        //Current browser version is 108.0.5359.125 with binary path C:\Program Files(x86)\Google\Chrome\Application\chrome.exe(SessionNotCreated)

                        if (ex.Message.Contains("This version of ChromeDriver only supports Chrome version"))
                        {
                            this.BrowserFullVersion = ex.Message.Split(new string[] { "urrent browser version is " }, StringSplitOptions.None)[1].Split(' ')[0];
                            this.BrowserMajorVersion = Convert.ToInt32(BrowserFullVersion.Split('.')[0]);

                            this.DriverFullVersion = null!;
                            this.DriverMajorVersion = Convert.ToInt32(ex.Message.Split(new string[] { "only supports Chrome version " }, StringSplitOptions.None)[1].Split('\n')[0]);
                        }
                        else
                        {
                            Console.WriteLine(ex.Message);
                        }

                    }

                    /*
                    Console.WriteLine("BrowserFullVersion : " + this.BrowserFullVersion);
                    Console.WriteLine("DriverFullVersion : " + this.DriverFullVersion);
                    Console.WriteLine("BrowserMajorVersion : " + this.BrowserMajorVersion);
                    Console.WriteLine("DriverMajorVersion : " + this.DriverMajorVersion);
                    */

                    if (chrome.Driver != null)
                    {
                        chrome.Driver.Quit();
                        chrome.Driver = null;
                    }

                });
            } // end of method

            public async Task LatestRelease(HttpClient httpClient)
            {
                await this.BrowserLatestRelease(httpClient);

                Console.WriteLine("BrowserFullVersion (in LatestRelease Method) : " + this.BrowserFullVersion);
                Console.WriteLine("BrowserMajorVersion (in LatestRelease Method) : " + this.BrowserMajorVersion);

                await this.DriverMatchedLatestRelease(this.BrowserFullVersion, httpClient);
            } // end of method

            public async Task BrowserLatestRelease_(HttpClient httpClient)
            {
                this._httpClient = httpClient;

                string csv = await this.GetBrowserVersionData();
                Console.WriteLine("Latest stable (from BrowserLatestRelease Method) : " + csv);

                if (csv == null)
                {
                    return;
                }

                string[] parts = csv.Split(new string[] { "win,stable," }, StringSplitOptions.None);

                if (parts.Length < 2)
                {
                    Console.WriteLine("Couldn't find the version tag in the CSV.");
                    return;
                }

                string[] version = parts[1].Split(new string[] { "," }, StringSplitOptions.None);

                if (version.Length == 0)
                {
                    Console.WriteLine("Couldn't extract the version from the CSV.");
                    return;
                }

                this.BrowserFullVersion = version[0];
                this.BrowserMajorVersion = Convert.ToInt32(version[0].Split('.')[0]);

                Console.WriteLine("BrowserFullVersion (in BrowserLatestRelease Method) : " + this.BrowserFullVersion);
                Console.WriteLine("BrowserMajorVersion (in BrowserLatestRelease Method) : " + this.BrowserMajorVersion);
            } // end of method

            public async Task BrowserLatestRelease(HttpClient httpClient)
            {
                this._httpClient = httpClient;

                this.BrowserFullVersion = await this.GetBrowserVersionData();
                this.BrowserMajorVersion = Convert.ToInt32(this.BrowserFullVersion.Split('.')[0]);

                Console.WriteLine("BrowserFullVersion (in BrowserLatestRelease Method) : " + this.BrowserFullVersion);
                Console.WriteLine("BrowserMajorVersion (in BrowserLatestRelease Method) : " + this.BrowserMajorVersion);
            } // end of method

            public async Task DriverMatchedLatestRelease(string BrowserFullVersion, HttpClient httpClient)
            {
                this._httpClient = httpClient;

                string? json = await this.GetDriverVersionData();

                if (json == null)
                {
                    return;
                }

                JObject jsonObject = JObject.Parse(json);
                JObject milestones = (JObject)jsonObject["milestones"]!;

                // 116
                //Console.WriteLine("BrowserFullVersion : " + BrowserFullVersion);
                string versionMajor = BrowserFullVersion.Split(new string[] { "." }, StringSplitOptions.None)[0];

                if (milestones.ContainsKey(versionMajor))
                {
                    JObject milestone = (JObject)milestones[versionMajor]!; // object 116
                    JArray chromedriverDownloads = (JArray)milestone["downloads"]!["chromedriver"]!;

                    // https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/116.0.5845.96/win32/chromedriver-win32.zip
                    string downloadUrl = chromedriverDownloads
                        .Where(item => item["platform"]!.Value<string>() == "win32")
                        .Select(item => item["url"]!.Value<string>())
                        .FirstOrDefault()!;

                    string fullVersion = (string)milestone["version"]!;

                    this.DriverFullVersion = fullVersion;
                    this.DriverMajorVersion = Convert.ToInt32(fullVersion.Split('.')[0]);
                    this.DriverDownloadURL = downloadUrl;
                }
                else
                {
                    Console.WriteLine($"Versi {versionMajor} tidak ditemukan.");
                    return;
                }

            } // end of method

            private async Task<string> GetBrowserVersionData()
            {
                string url = "https://chromiumdash.appspot.com/fetch_milestones?only_branched=true";

                try
                {
                    HttpResponseMessage response = await _httpClient.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    string responseContent = await response.Content.ReadAsStringAsync();
                    var milestones = JsonConvert.DeserializeObject<List<dynamic>>(responseContent);

                    // Mengambil milestone, v8_branch, dan chromium_branch
                    var latestMilestone = milestones![0];
                    var milestoneNumber = latestMilestone.milestone;
                    var v8Branch = latestMilestone.v8_branch.ToString();
                    var chromiumBranch = latestMilestone.chromium_branch;

                    // Mengambil angka di belakang koma dari v8_branch
                    var v8BranchMinor = v8Branch.Contains(".") ? v8Branch.Split('.')[1] : "0";

                    // Membangun string versi lengkap
                    string fullVersion = $"{milestoneNumber}.{v8BranchMinor}.{chromiumBranch}";

                    return fullVersion;
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine("Terjadi kesalahan saat mengirim permintaan GET:");
                    Console.WriteLine(e.Message);

                    return null!;
                }
            }


            private async Task<string> GetDriverVersionData()
            {
                string url = "https://googlechromelabs.github.io/chrome-for-testing/latest-versions-per-milestone-with-downloads.json";

                try
                {
                    HttpResponseMessage response = await _httpClient.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    //string readed = await response.Content.ReadAsStringAsync();
                    //Console.Out.WriteLineAsync(readed);

                    return await response.Content.ReadAsStringAsync();
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine("Terjadi kesalahan saat mengirim permintaan GET:");
                    Console.WriteLine(e.Message);

                    return null!;
                }
            } // end of method

        } // end of class Version

        public bool IsDriverConnected()
        {
            if (Driver == null) return false;

            try
            {
                string _ = Driver.Url;
                return true;
            }
            catch
            {
                return false;
            }
        } // end of method

        public static string ToLower(string attribute)
        {

            return $"translate({attribute}, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')";

        } // end of method

    } // end of class

    public static class InstanceHelper
    {
        // try click for 10 seconds (1 time per second)
        public static SafeClickResult SafeClick(this WebElement webElement, int times = 10)
        {
            SafeClickResult result = new SafeClickResult();

            if (webElement == null || webElement.Item == null)
            {
                result.Status = false;
                result.Type = 0;
                result.Message = "Element is null";

                Console.WriteLine(result.Message);

                return result;
            }

            int i = 1;

            while (true)
            {
                try
                {
                    webElement.Item.Click();
                    break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(webElement.Selector + " gagal diklik");

                    Thread.Sleep(1000);
                    if (i >= times)
                    {
                        result.Status = false;
                        result.Type = 2;
                        result.Message = $"'{webElement.Selector}' element click is failed for {i} times / seconds. Exception : {ex.Message}";

                        Console.WriteLine(result.Message);

                        return result;
                    }

                    i++;
                }
            }

            result.Status = true;
            result.Type = 1;
            result.Message = $"Element '{webElement.Selector}' is successfully clicked at {i} tries / seconds.";

            Console.WriteLine(result.Message);

            return result;
        } // end of method

        public static SafeSendKeysResult SafeSendKeys(this WebElement webElement, string text, int seconds = 10)
        {
            SafeSendKeysResult result = new SafeSendKeysResult();

            if (webElement == null || webElement.Item == null)
            {
                result.Status = false;
                result.Type = 0;
                result.Message = "Element is null";

                return result;
            }

            int i = 1;

            while (true)
            {
                try
                {
                    webElement.Item.SendKeys(text);
                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);

                    if (i >= seconds)
                    {
                        result.Status = false;
                        result.Type = 2;
                        result.Message = $"Send keys is failed for {i} times / seconds. Exeption : {ex.Message}";

                        return result;
                    }

                    i++;
                }
            }

            result.Status = true;
            result.Type = 1;
            result.Message = $"Send keys is done at {i} tries / seconds.";

            return result;
        } // end of method

        private static object clipboardLock = new object();

        public static SafeSendKeysResult SafeCopyAndPaste_(this WebElement webElement, string text, int seconds = 10)
        {
            SafeSendKeysResult result = new SafeSendKeysResult();

            if (webElement == null || webElement.Item == null)
            {
                result.Status = false;
                result.Type = 0;
                result.Message = "Element is null";
                return result;
            }

            int i = 1;

            while (true)
            {
                try
                {
                    ((IJavaScriptExecutor)webElement.Driver!).ExecuteScript("arguments[0].focus(); arguments[0].innerText = arguments[1];", webElement.Item, text);

                    result.Status = true;
                    result.Type = 1;
                    result.Message = $"Send keys is done at {i} tries / seconds.";
                    return result;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    if (i >= seconds)
                    {
                        result.Status = false;
                        result.Type = 2;
                        result.Message = $"Send keys is failed for {i} times / seconds. Exception : {ex.Message}";
                        return result;
                    }
                    i++;
                }
            }
        } // end of method

        public static SafeSendKeysResult SafeCopyAndPaste(this WebElement webElement, string text, int seconds = 10)
        {
            SafeSendKeysResult result = new SafeSendKeysResult();

            if (webElement == null || webElement.Item == null)
            {
                result.Status = false;
                result.Type = 0;
                result.Message = "Element is null";

                return result;
            }

            int i = 1;

            while (true)
            {
                try
                {
                    // Use clipboard to copy and paste
                    Thread t = new Thread(() =>
                    {
                        lock (clipboardLock)
                        {
                            try
                            {
                                // Copy to clipboard
                                Clipboard.SetText(text);

                                // Simulate a Ctrl+V paste operation
                                webElement.Item.SendKeys(OpenQA.Selenium.Keys.Control + "v");

                                // Optional: clear the clipboard after pasting
                                Clipboard.Clear();
                            }
                            catch
                            {
                                Clipboard.Clear();
                            }
                        }
                    });

                    #if WINDOWS
                    t.SetApartmentState(ApartmentState.STA);
                    t.Start();
                    t.Join();
                    #endif

                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);

                    if (i >= seconds)
                    {
                        result.Status = false;
                        result.Type = 2;
                        result.Message = $"Send keys is failed for {i} times / seconds. Exception : {ex.Message}";

                        return result;
                    }

                    i++;
                }
            }

            result.Status = true;
            result.Type = 1;
            result.Message = $"Send keys is done at {i} tries / seconds.";

            return result;
        } // end of method


        public static string? GetCssValue(this WebElement webElement, string cssAtrribute)
        {
            if (webElement == null || webElement.Item == null) return null;

            try
            {
                string result = webElement.Item.GetCssValue(cssAtrribute);
                return result;
            }
            catch
            {
                return null;
            }
        } // end of method

        public static void ClickJs(this WebElement webElement)
        {
            
            if (webElement.Driver == null) return;

            string javascript = "arguments[0].click();";
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)webElement.Driver;
            jsExecutor.ExecuteScript(javascript, webElement.Item);

        } // end of method

        public static SafeClickResult SafeClickJs(this WebElement webElement, int seconds = 10)
        {
            SafeClickResult result = new SafeClickResult();

            if (webElement == null)
            {
                result.Status = false;
                result.Type = 0;
                result.Message = "Element is null";

                Console.WriteLine(result.Message);

                return result;
            }

            int i = 1;

            while (true)
            {
                try
                {
                    webElement.ClickJs();
                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);

                    if (i >= seconds)
                    {
                        result.Status = false;
                        result.Type = 2;
                        result.Message = $"Element click is failed for {i} times / seconds. Exception : {ex.Message}";

                        Console.WriteLine(result.Message);

                        return result;
                    }

                    i++;
                }
            }

            result.Status = true;
            result.Type = 1;
            result.Message = $"Element is successfully clicked at {i} tries / seconds.";

            Console.WriteLine(result.Message);

            return result;
        } // end of method

        public static WebElement FindElementByXPath(this WebElement webElement, string XPathSelector)
        {
            WebElement result = new WebElement();

            if (!webElement.State || webElement.Item == null)
            {
                result.State = false;
                result.Message = "Input element does NOT exist.";
                result.Selector = XPathSelector;
                result.Item = null;
                result.Driver = webElement.Driver;

                return result;
            }

            IWebElement webElement_;

            try
            {
                webElement_ = webElement.Item.FindElement(By.XPath(XPathSelector));

                result.State = true;
                result.Message = "Element is found.";
                result.Selector = XPathSelector;
                result.Item = webElement_;
                result.Driver = webElement.Driver;
            }
            catch
            {
                result.State = false;
                result.Message = "Element is NOT found.";
                result.Selector = XPathSelector;
                result.Item = null;
                result.Driver = webElement.Driver;
            }

            return result;
        } // end of method

        public static object? GetJavaScriptProperty(this WebElement webElement, string propertyName)
        {
            object? propertyValue = null;

            if (!webElement.State || webElement.Item == null)
            {
                return propertyValue;
            }

            try
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)webElement.Driver!;
                propertyValue = js.ExecuteScript($"return arguments[0].{propertyName};", webElement.Item);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred while getting JavaScript property '{propertyName}': {ex.Message}");
            }

            return propertyValue;
        } // end of method


        public static bool Disappeared(this WebElement webElement, int timeout = 10)
        {
            if (webElement.Item == null) return true;

            for (int i = 0; i < timeout; i++)
            {
                try
                {
                    bool itemEnabled = webElement.Item.Enabled;
                }
                catch
                {
                    return true;
                }

                Thread.Sleep(1000);
            }

            return false;
        } // end of element
    } // end of class

    public class SafeClickResult
    {
        public bool Status { get; set; }
        public int Type { get; set; } // 0. element is null; 1. clicked successfully; 2. click failed
        public string Message { get; set; } = string.Empty;
    } // end of class

    public class SafeSendKeysResult
    {
        public bool Status { get; set; }
        public int Type { get; set; } // 0. element is null; 1. sendkeys successfully; 2. sendkeys failed
        public string Message { get; set; } = string.Empty;
    } // end of class

    public class WebElement
    {
        public bool State { get; set; }
        public string Message { get; set; } = string.Empty;
        public string Selector { get; set; } = string.Empty;
        public IWebElement? Item { get; set; }
        public IWebDriver? Driver { get; set; }
    } // end of class

    public class WebElements
    {
        public bool State { get; set; }
        public string Message { get; set; } = string.Empty;
        public string Selector { get; set; } = string.Empty;
        public int Count { get; set; } = 0;
        public List<WebElement> Items { get; set; } = new List<WebElement>();
        public IWebDriver? Driver { get; set; }
    } // end of class

    public class WebPage
    {
        public bool State { get; set; }
        public string Message { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
    } // end of class

    public class Browser
    {
        public bool State { get; set; } = false;
        public string Message { get; set; } = string.Empty;
    } // end of class

    public class CookiesData
    {
        public bool Status { get; set; }
        public string Message { get; set; } = string.Empty;
        public List<OpenQA.Selenium.Cookie> Items { get; set; } = new List<OpenQA.Selenium.Cookie>();
        public int Amount { get; set; } = 0;
        public string ExceptionName { get; set; } = string.Empty;
    } // end of class

    public class CookieData
    {
        public bool Status { get; set; }
        public string Message { get; set; } = string.Empty;
        public OpenQA.Selenium.Cookie? Cookie { get; set; }
    } // end of class

    public class CookieInStrings
    {
        public string Name { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public string Domain { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public string Expiry { get; set; } = string.Empty;
    } // end of method

    public class CookiesFile
    {
        public bool Status { get; set; }
        public string Message { get; set; } = string.Empty;
        public CookiesData? CookiesData { get; set; }
        public string filename { get; set; } = string.Empty;
    } // end of class

    public class CookieDetails
    {
        public string Name { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public string Domain { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public string Expiry { get; set; } = string.Empty;
    } // end of class

    public class ScrollResult
    {
        public bool Status { get; set; }
        public string Message { get; set; } = string.Empty;
        public int Height { get; set; }
    } // end of class

    public class SelectElement : OpenQA.Selenium.Support.UI.SelectElement
    {
        public SelectElement(IWebElement element) : base(element)
        {

        } // end of method
    } // end of method

} // end of namespace
