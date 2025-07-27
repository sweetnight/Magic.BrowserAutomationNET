using System.Diagnostics;
using System.Globalization;
using System.Management;
using System.Text.Json;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

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
        public string UserDataDir { get; set; } = string.Empty; // prioritas dibanding Profile. Jika diisi, maka Profile diabaikan.
        public string Profile { get; set; } = string.Empty;
        
        // hanya sampai direktori tempat file .exe
        public string? DriverDirectory { get; set; }

        // full path sampai ke filename .exe ATAU bisa jg sampai direktori lokasi .exe nya saja
        public string? BinaryLocation { get; set; }
        public bool SaveResources { get; set; } = true;

        public IWebDriver? Driver { get; set; }

        public ChromeDriverService? ChromeDriverService { get; set; }
        public ChromeOptions ChromeOptions { get; set; } = new ChromeOptions();

        public Chrome(string? binaryLocation = null, string? driverDirectory = null, string? instanceId = null)
        {

            this.BinaryLocation = binaryLocation;
            this.DriverDirectory = driverDirectory;
            this.InstanceId = instanceId == null ? HelperNET.CreateRandomAlphabeticLowerCaseString(5) : instanceId;

        } // end of method

        public void SetChromeOptions()
        {
            if (BinaryLocation != null)
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

            if(!string.IsNullOrWhiteSpace(UserDataDir))
            {
                ChromeOptions.AddArgument($@"--user-data-dir={UserDataDir}");
            }
            else if (Profile != string.Empty)
            {
                string path = Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%");

                ChromeOptions.AddArgument($@"--user-data-dir={path}/Google/Chrome/User Data");
                ChromeOptions.AddArgument($"--profile-directory={Profile}");

                /*
                string path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Google",
                    "Chrome",
                    "User Data"
                );

                ChromeOptions.AddArgument($"--user-data-dir={path}");
                ChromeOptions.AddArgument($"--profile-directory={Profile}");
                */
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

            if (DriverDirectory == null)
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
                ChromeOptions.AddArgument("--disable-dev-shm-usage");
                ChromeOptions.AddArgument("--disable-features=TranslateUI");
                ChromeOptions.AddArgument("--disable-background-timer-throttling");
                ChromeOptions.AddArgument("--disable-backgrounding-occluded-windows");
                ChromeOptions.AddArgument("--disable-renderer-backgrounding");
                ChromeOptions.AddArgument("--disable-default-apps");
                ChromeOptions.AddArgument("--disable-translate");
                ChromeOptions.AddArgument("--no-first-run");
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
                Debug.WriteLine(ex.Message);

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

            Debug.WriteLine("Chrome ==================== : Akan melakukan Navigate");

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

                Debug.WriteLine("Chrome ==================== : Navigate berhasil: " + Url);
            }
            catch (Exception ex)
            {
                returnResults.State = false;
                returnResults.Message = "Navigate failed. Exception : " + ex.Message;
                returnResults.Url = string.Empty;

                Debug.WriteLine("Chrome ==================== : Navigate exception: " + ex.Message);
            }

            return returnResults;
        } // end of function

        public WebPage Refresh()
        {
            
            Debug.WriteLine("Chrome ==================== : Akan melakukan refresh");

            WebPage returnResults = new WebPage();

            try
            {
                if (Driver != null && Driver.GetType() == typeof(ChromeDriver))
                {
                    Driver.Navigate().Refresh();

                    string Url = GetCurrentUrl().Url;
                    returnResults.State = true;
                    returnResults.Url = Url;
                    returnResults.Message = "Refresh successfully.";

                    Debug.WriteLine("Chrome ==================== : Refresh berhasil: " + Url);
                }
            }
            catch (Exception ex)
            {
                returnResults.State = false;
                returnResults.Url = string.Empty;
                returnResults.Message = "Refresh failed. Exception : " + ex.Message;

                Debug.WriteLine("Chrome ==================== : Refresh exception: " + ex.Message);
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

            Debug.WriteLine("Chrome ==================== : Akan select elemen by xpath: " + XPathSelector);

            WebElement webElement = new WebElement();

            if (Driver == null)
            {
                webElement.State = false;
                webElement.Message = "Driver is null.";

                Debug.WriteLine("Chrome ==================== : Gagal select elemen by xpath 'Driver is null': " + XPathSelector);

                return webElement;
            }

            try
            {
                Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(timeSpan);

                webElement.Item = Driver.FindElement(By.XPath(XPathSelector));
                webElement.State = true;
                webElement.Message = "Element is found.";

                Debug.WriteLine("Chrome ==================== : Berhasil select elemen by xpath: " + XPathSelector);

            }
            catch (Exception ex)
            {
                webElement.State = false;
                webElement.Message = "Element is NOT found. Exception: " + ex;

                Debug.WriteLine("Chrome ==================== : Gagal select elemen by xpath: " + XPathSelector + ". Exception: " + ex.Message);
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

                    Debug.WriteLine("Chrome ==================== : Gagal select elemen by xpath: " + XPathSelector + ". Exception: " + ex.Message);
                }
            }

            webElement.Selector = XPathSelector;
            webElement.Driver = Driver;
            webElement.Chrome = this;

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

            Debug.WriteLine("Chrome ==================== : Akan menghapus semua cookies");

            if (Driver == null) return;

            try
            {
                // Hapus semua cookies
                Driver.Manage().Cookies.DeleteAllCookies();
                Debug.WriteLine("Chrome ==================== : Berhasil menghapus semua cookies");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Chrome ==================== : Gagal menghapus semua cookies: " + ex.Message);
            }

        } // end of method

        public string GetLocalStorageJson()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver!;
            return (string)js.ExecuteScript("return JSON.stringify(localStorage);");
        } // end of method

        public void SetLocalStorageJson(string localStorageJson)
        {
            if (Driver is IJavaScriptExecutor js)
            {
                // Escape string agar tidak error di JavaScript
                string escapedJson = localStorageJson
                    .Replace(@"\", @"\\")  // Escape backslash
                    .Replace("`", "\\`");  // Escape backtick karena pakai template literal

                string script = $@"
                    var items = JSON.parse(`{escapedJson}`);
                    for (var key in Object.keys(items)) {{
                        localStorage.setItem(key, items[key]);
                    }}
                ";

                js.ExecuteScript(script);
            }
        } // end of method

        public void WaitForPageToLoad()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver!;
            int timeoutInSeconds = 10;
            WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(timeoutInSeconds));

            wait.Until(driver => js.ExecuteScript("return document.readyState").ToString() == "complete");
        } // end of method

        public void ClearSessionAndReload()
        {
            if (Driver == null) return;

            try
            {
                // Hapus semua cookies
                Driver.Manage().Cookies.DeleteAllCookies();

                // Hapus cache dan storage
                IJavaScriptExecutor js = (IJavaScriptExecutor)Driver!;
                js.ExecuteScript("window.localStorage.clear();");
                js.ExecuteScript("window.sessionStorage.clear();");

                // Paksa reload halaman untuk memastikan state di-reset
                js.ExecuteScript("window.location.reload(true);");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Gagal membersihkan sesi: {ex.Message}");
            }
        } // end of method

        public void DeleteAllResources()
        {

            if (Driver == null) return;

            try
            {
                // Hapus semua cookies
                Driver.Manage().Cookies.DeleteAllCookies();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Gagal menghapus cookies: {ex.Message}");
            }

            try
            {
                // Hapus cache
                IJavaScriptExecutor js = (IJavaScriptExecutor)Driver!;
                js.ExecuteScript("window.localStorage.clear();");
                js.ExecuteScript("window.sessionStorage.clear();");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Gagal menghapus cache: {ex.Message}");
            }

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

            return JsonSerializer.Serialize(cookieInStrings);
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

            Debug.WriteLine("Chrome ==================== : Akan melakukan AddCookiesFromJson");

            List<CookieInStrings> cookieInStrings = JsonSerializer.Deserialize<List<CookieInStrings>>(jsonString)!;

            foreach (CookieInStrings cookieItem in cookieInStrings!)
            {
                if (cookieItem.Expiry == "")
                {
                    try
                    {
                        CookieData cookieData = AddCookie(cookieItem.Name, cookieItem.Value, cookieItem.Domain, cookieItem.Path, null);
                        Debug.WriteLine("Chrome ==================== : Berhasil melakukan AddCookiesFromJson di try pertama");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("Chrome ==================== : Gagal melakukan AddCookiesFromJson di catch pertama: " + ex.Message);
                    }
                }
                else
                {
                    DateTime expiry = new DateTime();

                    try
                    {
                        expiry = DateTime.ParseExact(cookieItem.Expiry, "M/d/yyyy h:m:s tt", new CultureInfo("en-US"));
                    }
                    catch
                    {
                        continue;
                    }

                    try
                    {
                        CookieData cookieData = AddCookie(cookieItem.Name, cookieItem.Value, cookieItem.Domain, cookieItem.Path, expiry);
                        Debug.WriteLine("Chrome ==================== : Berhasil melakukan AddCookiesFromJson di try kedua");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("Chrome ==================== : Gagal melakukan AddCookiesFromJson di catch kedua: " + ex.Message);
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

        public long ScrollToTop(WebElement webElement, int wait = 0)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver!;
            // Scroll elemen ke atas dan ambil nilai scrollTop setelah digulir
            long scrollTop = (long)js.ExecuteScript("arguments[0].scrollTop = 0; return arguments[0].scrollTop;", webElement.Item);

            if(wait > 0)
            {
                Thread.Sleep(wait); // Tunggu sesuai durasi yang diberikan
            }

            return scrollTop;
        } // end of method

        public void ScrollToElement(WebElement webElement, int wait = 0)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver!;

            // Gunakan JavaScript untuk scroll elemen ke dalam viewport
            js.ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", webElement.Item);

            // Tambahkan jeda kecil jika diperlukan

            if(wait > 0)
            {
                Thread.Sleep(wait);
            }
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

        public bool InjectScriptFromFile(string jsFile, out string message)
        {

            string jsCode = File.ReadAllText(jsFile);
            return InjectScript(jsCode, out message);

        } // end of method

        public bool InjectScript(string jsCode, out string message)
        {

            message = string.Empty;

            try
            {
                if (Driver == null)
                {
                    message = "Chrome Driver is not initialized yet.";
                    return false;
                }

                IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)Driver;

                jsExecutor.ExecuteScript(jsCode);

                message = "Script is injected successfully.";
                return true;
            }
            catch (Exception ex)
            {
                message = "Fail inject script. Exception: " + ex.Message;
                return false;
            }

        } // end of method

        public class Version
        {

            public static async Task<(string LatestVersion, string ChromeForTestingDownloadURL, string ChromedriverDownloadURL)> GetLatestChromeData()
            {

                string apiUrl = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json";

                using (HttpClient client = new HttpClient())
                {
                    try
                    {
                        // Mengambil data JSON dari URL
                        string jsonResponse = await client.GetStringAsync(apiUrl);

                        // Mendapatkan data versi stable
                        using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                        {
                            JsonElement root = doc.RootElement;
                            JsonElement stable = root.GetProperty("channels").GetProperty("Stable");
                            string stableVersion = stable.GetProperty("version").GetString()!;

                            string chromeUrl = string.Empty;
                            string chromedriverUrl = string.Empty;

                            // Ambil URL Chrome win64
                            foreach (var download in stable.GetProperty("downloads").GetProperty("chrome").EnumerateArray())
                            {
                                if (download.GetProperty("platform").GetString() == "win64")
                                {
                                    chromeUrl = download.GetProperty("url").GetString()!;
                                    break;
                                }
                            }

                            // Ambil URL Chromedriver win64
                            foreach (var download in stable.GetProperty("downloads").GetProperty("chromedriver").EnumerateArray())
                            {
                                if (download.GetProperty("platform").GetString() == "win64")
                                {
                                    chromedriverUrl = download.GetProperty("url").GetString()!;
                                    break;
                                }
                            }

                            return (stableVersion, chromeUrl, chromedriverUrl);
                        }
                    }
                    catch (HttpRequestException e)
                    {
                        Console.WriteLine($"Error saat mengakses URL: {e.Message}");
                    }
                    catch (System.Text.Json.JsonException e)
                    {
                        Console.WriteLine($"Error saat memproses data JSON: {e.Message}");
                    }

                    return ("0.0.0.0", string.Empty, string.Empty);
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

        public bool KillBrowserByUserDataDir(string userDataDir)
        {

            bool killed = false;

            Process[] processes = Process.GetProcessesByName("chrome");

            foreach (Process proc in processes)
            {
                try
                {
                    string query = $"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {proc.Id}";

                    using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(query))
                    {
                        foreach (ManagementObject obj in searcher.Get())
                        {
                            string? cmdLine = obj["CommandLine"]?.ToString();
                            //Debug.WriteLine($"{cmdLine}");

                            //if (!string.IsNullOrEmpty(cmdLine) && cmdLine.Contains($"--user-data-dir=\"{userDataDir}\"", StringComparison.OrdinalIgnoreCase))
                            if (!string.IsNullOrEmpty(cmdLine) && cmdLine.Contains($"--user-data-dir=\"{userDataDir}\"", StringComparison.OrdinalIgnoreCase) && !cmdLine.Contains("--type=", StringComparison.OrdinalIgnoreCase)) // Hanya parent
                            {
                                proc.Kill();
                                proc.WaitForExit();
                                Debug.WriteLine($"Killed process {proc.Id} with user-data-dir match");
                                killed = true;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error: {ex.Message}");
                }
            }

            return killed;

        } // end of method

        /*
        public bool KillBrowserByUserDataDir(string userDataDir)
        {
            bool killed = false;

            string psCommand = $@"
                Get-CimInstance Win32_Process -Filter ""Name='chrome.exe'"" |
                Where-Object {{ $_.CommandLine -like '*--user-data-dir=""{userDataDir}""*' -and $_.CommandLine -notlike '*--type=*' }} |
                Select-Object -ExpandProperty ProcessId
            ";

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "powershell",
                Arguments = $"-NoProfile -Command \"{psCommand}\"",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            Process psProc = Process.Start(psi)!;
            string output = psProc.StandardOutput.ReadToEnd();
            psProc.WaitForExit();

            string[] lines = output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines)
            {
                if (int.TryParse(line.Trim(), out int pid))
                {
                    try
                    {
                        Process p = Process.GetProcessById(pid);
                        p.Kill();
                        p.WaitForExit();
                        Debug.WriteLine($"Killed process {pid} with user-data-dir match");
                        killed = true;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error killing process {pid}: {ex.Message}");
                    }
                }
            }

            return killed;
        }
        */


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
                        result.Message = $"'{webElement.Selector}' element click is failed for {i} timeout / seconds. Exception : {ex.Message}";

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
                        result.Message = $"Send keys is failed for {i} timeout / seconds. Exeption : {ex.Message}";

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
                        result.Message = $"Send keys is failed for {i} timeout / seconds. Exception : {ex.Message}";
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
                        result.Message = $"Send keys is failed for {i} timeout / seconds. Exception : {ex.Message}";

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
                        result.Message = $"Element click is failed for {i} timeout / seconds. Exception : {ex.Message}";

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

        public static string? SafeGetAttribute(this WebElement webElement, string attribute, int timeout = 10)
        {

            for (int attempt = 0; attempt < timeout; attempt++)
            {
                try
                {
                    if (webElement.Item == null) return null;

                    return webElement.Item.GetAttribute(attribute);
                }
                catch (StaleElementReferenceException)
                {
                    Debug.WriteLine($"Chrome ==================== : SafeGetAttribute: StaleElementReferenceException di attempt {attempt + 1} dari {timeout}");

                    if (webElement.Chrome == null || string.IsNullOrEmpty(webElement.Selector))
                        return null;

                    // Refresh elemen
                    WebElement fresh = webElement.Chrome.FindElementByXPath(webElement.Selector, 5);

                    if (!fresh.State || fresh.Item == null)
                        return null;

                    // Update referensi elemen lama
                    webElement.Item = fresh.Item;
                    webElement.State = fresh.State;
                    webElement.Message = fresh.Message;

                    Thread.Sleep(1000); // delay kecil
                }
                catch
                {
                    return null;
                }
            }

            return null;

        } // end of method


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
        public Chrome? Chrome { get; set; }
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
