using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager;
using Keys = OpenQA.Selenium.Keys;
using System.Web;
using System.Drawing;
using System.IO;
using System.Text;
using WebDriverManager.Helpers;
using System.Diagnostics;
using System.Reflection;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        List<string> dataColumns = new List<string>();
        //Dictionary<string, (string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string StreetAddress, string city, string zip, string country, Dictionary<string, string> socialMedias, string companyWebsite)> uniqueDataPair = new Dictionary<string, (string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string StreetAddress, string city, string zip, string country, Dictionary<string, string> socialMedias, string companyWebsite)>();
        private string[] searchTerms;
        private Font mulishRegularFont;
        private IWebDriver driver;
        private IWebDriver chromeDriverForBusinessData;
        private int driverProcessId;
        //private string[] userAgents = {
        //    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.5735.134 Safari/537.36",
        //    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.5672.127 Safari/537.36",
        //    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.5628.126 Safari/537.36"
        //};
        private int businessDriverProcessId;
        private List<BusinessInfo> tempRows = new List<BusinessInfo>();
        private int batchSize = 35;
        private static readonly HashSet<string> CountryList = new HashSet<string>
        {
            "Germany", "Australia", "United States", "USA", "Canada", "France", "Spain",
            "United Kingdom", "UK", "India", "Italy", "Brazil", "Mexico", "Netherlands",
            "Sweden", "Denmark", "Norway", "Bangladesh"
        };
        private CancellationTokenSource cancellationTokenSource;
        List<string> chromeOptionArguements = new List<string> {
            "--headless",
            "--disable-gpu",
            "--disable-extensions",
            "--no-sandbox",
            "--log-level=3",
            "--disable-popup-blocking",
            "--disable-popup-blocking",
            "--enable-features=NetworkService,NetworkServiceInProcess",
            "--enable-async-dns",
           "--disable-site-isolation-trials",
           "--renderer-process-limit=5",
           "--incognito",
            "--disable-extensions",
            "--reduce-security-for-testing" };
        List<string> socialMediaSelectors = new List<string>
                {
                    "a[href*='mailto:']",
                    "a[href*='facebook.com']",
                    "a[href*='twitter.com']",
                    "a[href*='linkedin.com']",
                    "a[href*='instagram.com']",
                    "a[href*='youtube.com']",
                    "a[href*='pinterest.com']"
                };
        private List<BusinessInfo> results;
        public Form1()
        {
            // Load font from file (make sure the file path is correct)
            string fontPath = Path.Combine(Application.StartupPath, "assets", "fonts", "DancingScript-Regular.ttf");
            //mulishRegularFont = FontLoader.LoadFont(FontStyle.Regular);
            InitializeComponent();
            foreach (var resource in Assembly.GetExecutingAssembly().GetManifestResourceNames())
            {
                Console.WriteLine(resource);
            }

            //FontLoader.ApplyFontToAllControls(this, mulishRegularFont);
            //btnUpload.Font = mulishRegularFont;
            //btnUpload.Text = "djdjdjdj";
            LoggerService.Info("Sky Crawler Application Started.");
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ResizeEnd += Form1_Resize;
            string iconPath = Path.Combine(Application.StartupPath, "assets", "icon", "app-icon.ico");
            this.Icon = new Icon(iconPath);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AdjustDataGridViewHeight();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            AdjustDataGridViewHeight();
        }

        void manageStartButton(bool isEnable)
        {
            btnStart.Invoke(new Action(() =>
            {
                btnStart.Enabled = isEnable;
            }));
        }

        private void stopCrawlButton_Click(object sender, EventArgs e)
        {
            animatedLoader.Invoke(new Action(() =>
            {
                animatedLoader.Visible = false;
            }));
            if (cancellationTokenSource != null && !cancellationTokenSource.IsCancellationRequested)
            {
                cancellationTokenSource.Cancel();
                UpdateProgress("Stopping crawl...");
                LoggerService.Warning("Crawling process was manually stopped.");
                //progressBar.Visible = false;
            }
            manageStartButton(true);
        }

        private void exportDataButton_Click(object sender, EventArgs e)
        {
            ExportToExcel(results);
        }
        
        private void clearDataButton_Click(object sender, EventArgs e)
        {
            dataGridView.Rows.Clear();
            LoggerService.Info("Clearing data.");
            UpdateProgress("", true);
        }

        private IWebDriver RestartPrimaryWebDriver()
        {
            ChromeOptions options = new ChromeOptions();
            //options.AddArgument("--disable-software-rasterizer");
            options.AddArgument("--disable-gpu");  // Disables GPU acceleration
            options.AddArgument("--incognito");  // Disables GPU acceleration
            options.AddArgument("--disable-site-isolation-trials");
            options.AddArgument("--renderer-process-limit=5"); // Limit Chrome to 5 processes
            options.AddArgument("--disable-extensions"); // Prevents loading unnecessary extensions
            options.AddArgument("--disable-software-rasterizer"); // Prevents GPU fallback
            options.AddArgument("--no-sandbox");  // Helps in some environments
            options.AddArgument("--disable-dev-shm-usage"); // Prevents shared memory issues
            options.AddArgument("--disable-accelerated-2d-canvas"); // Avoids rendering crashes
            Random rand = new Random();
            //string randomUserAgent = userAgents[rand.Next(userAgents.Length)];
            //options.AddArgument($"--user-agent={randomUserAgent}");

            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true; // ✅ Hide console window

            IWebDriver driver = new ChromeDriver(service, options);
            driver.Manage().Window.Maximize();
            driverProcessId = service.ProcessId;
            return driver;
        }

        private async Task StartCrawlingAsync(CancellationToken cancellationToken, bool fetchBusinessData)
        {
            try
            {
                //uniqueDataPair = new Dictionary<string, (string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string StreetAddress, string city, string zip, string country, Dictionary<string, string> socialMedias, string companyWebsite)>();
                new DriverManager().SetUpDriver(new ChromeConfig(), VersionResolveStrategy.MatchingBrowser); // Automatically downloads ChromeDriver
                driver = RestartPrimaryWebDriver();
                chromeDriverForBusinessData = GetChromeDriverForBusinessDataFetch();
                driver.Navigate().GoToUrl("https://www.google.com/maps/?hl=en&force=tt");
                int count = 0;
                foreach (string term in searchTerms)
                {
                    LoggerService.Info(term);
                    count++;
                    if (cancellationToken.IsCancellationRequested)
                    {
                        //lblStatus.Text = "Crawling stopped.";
                        break;
                    }
                    if (count > 10)
                    {
                        count = 0;
                        driver.Quit();
                        driver.Dispose();
                        driver = null;
                        KillChromeDriverProcess(driverProcessId);
                        chromeDriverForBusinessData.Quit();
                        chromeDriverForBusinessData.Dispose();
                        chromeDriverForBusinessData = null;
                        KillChromeDriverProcess(businessDriverProcessId);
                        LoggerService.Info("Driver closed.");
                        LoggerService.Info("New driver creating for primary data, 167");
                        driver = RestartPrimaryWebDriver();
                        chromeDriverForBusinessData = GetChromeDriverForBusinessDataFetch();
                        LoggerService.Info("New driver creating for business data, 189");
                        driver.Navigate().GoToUrl("https://www.google.com/maps/?hl=en&force=tt");
                        continue;
                    }
                    try
                    {
                        UpdateProgress($"Crawling: {term} - In progress");
                        await processCrawlingAsync(driver, chromeDriverForBusinessData, term, cancellationToken, fetchBusinessData);
                        UpdateProgress($"Crawling: {term} - Completed");
                    } catch(Exception exc)
                    {
                        Console.WriteLine(exc.Message);
                        LoggerService.Error("Crawling error - 174", exc);
                        //UpdateProgress(exc.Message);
                    }
                }
                animatedLoader.Invoke(new Action(() =>
                {
                    animatedLoader.Visible = false;
                }));
                driver.Quit();
                driver.Dispose();
                driver = null;
                KillChromeDriverProcess(driverProcessId);
                chromeDriverForBusinessData.Quit();
                chromeDriverForBusinessData.Dispose();
                chromeDriverForBusinessData = null;
                KillChromeDriverProcess(businessDriverProcessId);
                // Export results to Excel
                //ExportToExcel(results);
                if(tempRows.Count > 0)
                {
                    foreach (var row in tempRows)
                    {
                        InsertRowIntoDatatable(row, true);
                    }

                    tempRows.Clear();
                }
                UpdateProgress("Crawling finished.");
                LoggerService.Info("Crawling Completed Successfully.");
                manageStartButton(true);
                //progressBar.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during crawling: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //UpdateProgress("Error occurred." + "-155");
                LoggerService.Error("Crawling error - 199", ex);
            }
        }

        private void KillChromeDriverProcess(int processId)
        {
            if (processId == 0 ) return;
            try
            {
                Process chromeDriverProcess = Process.GetProcessById(processId);
                if (chromeDriverProcess != null)
                {
                    Console.WriteLine($"Killing ChromeDriver process with PID: {processId}");
                    chromeDriverProcess.Kill();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error killing ChromeDriver process: {ex.Message}");
            }
        }

        private void AppendAndScroll(string message)
        {
            // Append the new message
            progressTextBox.AppendText(message + "\r\n");

            // Scroll to the bottom
            progressTextBox.SelectionStart = progressTextBox.Text.Length; // Set the caret at the end
            progressTextBox.ScrollToCaret(); // Scroll to the caret
        }

        private void UpdateProgress(string message, bool empty = false)
        {
            if (progressTextBox.InvokeRequired)
            {
                // If called from a background thread, marshal to the UI thread
                progressTextBox.Invoke(new Action(() =>
                {
                    if(empty)
                    {
                        progressTextBox.Clear();
                    } else
                    {
                        AppendAndScroll(message);
                    }
                }));
            }
            else
            {
                if (empty)
                {
                    progressTextBox.Clear();
                }
                else
                {
                    AppendAndScroll(message);
                }
            }
        }

        public void ClearCookies(IWebDriver driver)
        {
            try
            {
                driver.Manage().Cookies.DeleteAllCookies(); // ✅ Clears all cookies
                //driver.Navigate().Refresh(); // ✅ Refreshes the page after clearing cookies
                Console.WriteLine("✅ Cookies cleared successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error clearing cookies: {ex.Message}");
            }
        }


        private async Task processCrawlingAsync(IWebDriver driver, IWebDriver webDriver, string term, CancellationToken cancellationToken, bool fetchBusinessData)
        {
            var searchBox = driver.FindElement(By.Id("searchboxinput"));
            ClearCookies(driver);
            searchBox.Clear();
            searchBox.SendKeys(term);
            searchBox.SendKeys(Keys.Enter);
            // Wait for the search box to load
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
            try
            {
                // Example: Wait for the search results container to be visible
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@class='m6QErb DxyBCb kA9KIf dS8AEf XiKgde ecceSd']"))); // Adjust selector based on target element
                var mapContainer = driver.FindElement(By.XPath("//div[@class='m6QErb DxyBCb kA9KIf dS8AEf XiKgde ecceSd']"));
                var processedResults = new HashSet<string>();
                bool hasMoreResults = true;
                IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
                int lastHeight = 0;
                bool hasUrl = true;

                while (hasMoreResults)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        hasMoreResults = false;
                        break;
                    }
                    // Extract search results

                    var resultElements = driver.FindElements(By.XPath("//div[@class='bfdHYd Ppzolf OFBs3e  ']"));
                    int i = 0;
                    foreach (var resultElement in resultElements)
                    {
                        if (cancellationToken.IsCancellationRequested)
                        {
                            hasMoreResults = false;
                            break;
                        }
                        try
                        {
                            string reviewCount = string.Empty;
                            string rating = string.Empty;
                            string title = resultElement.FindElement(By.CssSelector(".qBF1Pd.fontHeadlineSmall")).Text;
                            if (processedResults.Contains(title))
                            {
                                i++;
                                continue; // Skip already processed results
                            }
                            //lcr4fd S9kvJb
                            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
                            //var linkElement = resultElement.FindElements(By.ClassName("lcr4fd"));
                            string hrefValue = string.Empty;

                            processedResults.Add(title);
                            var clickableElement = driver.FindElements(By.CssSelector("a.hfpxzc"));
                            // ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", clickableElement[i]);
                            int height = clickableElement[i].Size.Height;

                            lastHeight = Convert.ToInt32(jsExecutor.ExecuteScript("return arguments[0].scrollHeight", mapContainer));
                            jsExecutor.ExecuteScript("arguments[0].scrollBy(0, arguments[1]);", mapContainer, height);
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", clickableElement[i]);
                            var elements = resultElement.FindElements(By.CssSelector(".UaQhfb.fontBodyMedium > .W4Efsd")).Last();
                            string category = string.Empty;
                            string mapLink = driver.Url;

                            if (elements != null)
                            {
                                category = elements.FindElements(By.CssSelector(":first-child")).First().Text;
                                if (elements != null && category.Length > 0)
                                {
                                    category = category.Split('·')[0];
                                }
                            }
                            clickableElement[i].Click();
                            //Thread.Sleep(30);
                            WebDriverWait newWait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

                            // ✅ Wait until the div with the required class is visible
                            IWebElement sidebar = newWait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@class='m6QErb DxyBCb kA9KIf dS8AEf XiKgde ']")));

                            //Console.WriteLine("✅ Business details sidebar detected.");
                            //Thread.Sleep(100);
                            i++;
                            var detailsElems = driver.FindElements(By.XPath("//div[@class='m6QErb DxyBCb kA9KIf dS8AEf XiKgde ']"));
                            IWebElement hrefItem = null;
                            try
                            {
                                hrefItem = detailsElems[0].FindElement(By.CssSelector("a[data-item-id='authority']"));
                            } catch(Exception exc)
                            {
                                hrefItem = null;
                            }
                            hrefValue = hrefItem == null ? string.Empty : hrefItem.GetAttribute("href");
                            string contactNumber = string.Empty;
                            try
                            {
                                contactNumber = resultElement.FindElement(By.ClassName("UsdlK")).Text;
                            }
                            catch (Exception ex)
                            {
                                //UpdateProgress(ex.Message + "-267");
                                LoggerService.Error("Crawling error - 312", ex);
                            }
                            string hourInfo = string.Empty;
                            try
                            {
                                var hourInfoElement = detailsElems[0].FindElement(By.CssSelector(".t39EBf.GUrTXd"));
                                hourInfo = hourInfoElement != null ? hourInfoElement.GetAttribute("aria-label") : string.Empty;
                            } catch(Exception exc)
                            {
                                LoggerService.Error("invalid hour info", exc);
                            }
                            var businessHours = new Dictionary<string, string>();
                            if (hourInfo != null && hourInfo.Length > 0)
                            {
                                businessHours = ConvertBusinessHours(hourInfo);
                            }
                            UpdateProgress($"Record no: {(results.Count + 1).ToString()}");
                            UpdateProgress($"Title: {title}");
                            UpdateProgress($"------------------------");
                            //var existedData = uniqueDataPair.ContainsKey($"{title}_{contactNumber}");
                            //if (existedData)
                            //{
                            //    var existingData = uniqueDataPair[$"{title}_{contactNumber}"];
                            //    results.Add(existingData);
                            //    InsertRowIntoDatatable(existingData);
                            //    hasUrl = false;
                            //    Thread.Sleep(50);
                            //    continue;
                            //}
                            try{
                                if (fetchBusinessData && hrefValue.Length > 0 && !hrefValue.ToLower().Contains("facebook"))
                                {
                                    hasUrl = true;
                                    //hrefValue = linkElement[0].GetDomAttribute("href");
                                    //if (batchSize == 0)
                                    //{
                                    //    webDriver.Quit();
                                    //    webDriver.Dispose();
                                    //    webDriver = GetChromeDriverForBusinessDataFetch();
                                    //    batchSize = 35;
                                    //}
                                    //UpdateProgress("going to call setchsocialMediaLinks");
                                    var task = Task.Run(() => FetchSocialMediaLinks(hrefValue, webDriver));
                                    keyValuePairs = await task;
                                }
                                else
                                {
                                    hasUrl = false;
                                    Thread.Sleep(10);
                                }
                            } catch (Exception ex)
                            {
                                //UpdateProgress(ex.Message+"-303");
                                LoggerService.Error("business data fetch error - 351", ex);
                                hasUrl = false;
                                Thread.Sleep(200);
                            }
                            //UpdateProgress("Processing other data after links");
                            var reviewListText = GetRatingReview(resultElement.FindElement(By.ClassName("W4Efsd")).Text);
                            if (reviewListText != null && reviewListText.Count > 1)
                            {
                                rating = reviewListText[0];
                                reviewCount = reviewListText[1];
                            }
                            string city = string.Empty;
                            string zip = string.Empty;
                            string country = string.Empty;
                            string location = string.Empty;
                            bool claim = false;
                            string streetLocation = string.Empty;
                            //string pattern = @"(?<street>[\d\s\w\W]+),\s(?<city>[A-Za-z\s]+)\s(?<state>[A-Za-z]+)\s(?<zip>\d{4}),\s(?<country>[A-Za-z\s]+)$";
                            try
                            {
                                var item = detailsElems[0].FindElement(By.CssSelector("button[data-item-id='address'] .Io6YTe.fontBodyMedium.kR99db.fdkmkc"));
                                location = item != null ? item.Text : string.Empty;
                                //var regex = new Regex(pattern);
                                var addressObj = ParseAddress(location);
                                streetLocation = elements.FindElements(By.CssSelector(":nth-child(2)")).First().Text;
                                city = addressObj.City;
                                zip = addressObj.ZipCode;
                                country = addressObj.Country;
                                streetLocation = addressObj.Street; 
                                if(streetLocation == string.Empty|| streetLocation == null)
                                {
                                    streetLocation = elements.FindElements(By.CssSelector(":nth-child(2)")).First().Text;
                                    if (streetLocation.Contains("Closed"))
                                    {
                                        streetLocation = "";
                                    }
                                }
                                //if (match.Success)
                                //{
                                //    streetLocation = match.Groups["street"].Value;
                                //}
                                //else
                                //{
                                //    streetLocation = elements.FindElements(By.CssSelector(":nth-child(2)")).First().Text;
                                //}
                            }
                            catch (Exception ex)
                            {
                                //UpdateProgress(ex.Message + "-343");
                                LoggerService.Error("Crawling error - 392", ex);
                                location = string.Empty;
                            }
                            string locateInText = string.Empty;
                            //var locationElems = driver.FindElements(By.XPath("//div[@class='Io6YTe fontBodyMedium kR99db fdkmkc']"));
                            //string location = locationElems != null && locationElems.Count > 0 ? locationElems[0].Text : string.Empty;
                            try
                            {
                                var checkClaim = detailsElems[0].FindElement(By.CssSelector("a[data-item-id='merchant']"));
                                claim = checkClaim != null && checkClaim.GetAttribute("href") != null;
                            }
                            catch (Exception ex) { 
                                claim = false;
                            }
                            try
                            {
                                var locateIn = detailsElems[0].FindElement(By.CssSelector("button[data-item-id='locatedin']"));
                                locateInText = locateIn != null ? locateIn.GetAttribute("aria-label") : string.Empty;
                            }
                            catch (Exception ex)
                            {
                                locateInText = string.Empty;
                            }
                            string attributes = string.Empty;
                            try
                            {
                                var attriElement = detailsElems[0].FindElements(By.CssSelector("div[data-item-id='place-info-links'] Io6YTe.fontBodyMedium.kR99db.fdkmkc "));
                                foreach(var item in attriElement)
                                {
                                    if(attributes.Length == 0)
                                    {
                                        attributes = item.Text;
                                    } else
                                    {
                                        attributes += $", {item.Text}";
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                
                            }
                            var dataTobeAdded = new BusinessInfo(
                                term, title, reviewCount, rating, contactNumber, 
                                category, location, streetLocation, city, zip, 
                                country, keyValuePairs, hrefValue, claim, hourInfo,
                                businessHours, locateInText, attributes, mapLink
                            );
                            results.Add(dataTobeAdded);
                            //uniqueDataPair.Add($"{title}_{contactNumber}", dataTobeAdded);
                            InsertRowIntoDatatable(dataTobeAdded);
                            var closeButton = detailsElems[0].FindElement(By.XPath("//button[contains(@aria-label, 'Close') and contains(@class, 'VfPpkd-icon-LgbsSe')]"));
                            //IWebElement closeButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//button[contains(@aria-label, 'Close') and contains(@class, 'VfPpkd-icon-LgbsSe')]")));
                            closeButton.Click();

                            //UpdateProgress("data extract finished"+" -360");
                            //driver.Navigate().Back();
                            //driver.Navigate().Back();

                        }
                        catch (Exception ex)
                        {
                            //UpdateProgress($"{ex.StackTrace}" + "-366");
                            //UpdateProgress($"{ex.Message}" + "-367");
                            LoggerService.Error("Crawling error - 418", ex);
                            // Skip if any element is missing
                        }
                    }
                    //IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    WebDriverWait scrollWait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    if (!IsEndOfScroll(driver))
                    {
                        scrollWait.Until(driver2 =>
                        {
                            hasMoreResults = ScrollToLoadMoreResults(driver, mapContainer, processedResults, jsExecutor);
                            return hasMoreResults;
                        });
                    } else
                    {
                        hasMoreResults = ScrollToLoadMoreResults(driver, mapContainer, processedResults, jsExecutor);
                    }
                   
                    //if (!fetchBusinessData || !hasUrl)
                    //{
                    //    Thread.Sleep(2800);
                    //} else
                    //{
                    //    Thread.Sleep(1000);
                    //}
                    //double lastHeight = Convert.ToDouble(((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].scrollHeight", mapContainer));
                    if (cancellationToken.IsCancellationRequested)
                    {
                        //lblStatus.Text = "Crawling stopped.";
                        hasMoreResults = false;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Timeout waiting for search results to load.");
                //UpdateProgress(ex.StackTrace +  ex.Message + "-403");
                LoggerService.Error("Crawling error - 456", ex);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            LoggerService.Info("Sky Crawler Application Closed.");
            LoggerService.Close();
            if (driver != null)
            {
                driver.Quit();
                driver.Dispose();
                driver = null;
                KillChromeDriverProcess(driverProcessId);
            }
            
            if (chromeDriverForBusinessData != null)
            {
                chromeDriverForBusinessData.Quit();
                chromeDriverForBusinessData.Dispose();
                chromeDriverForBusinessData = null;
                KillChromeDriverProcess(businessDriverProcessId);
            }
        }

        private bool IsEndOfScroll(IWebDriver driver)
        {
            try
            {
                // Locate the specific div with classes "m6QErb XiKgde tLjsW eKbjU"
                IWebElement endMessageContainer = driver.FindElement(By.CssSelector("div.m6QErb.XiKgde.tLjsW.eKbjU"));

                // Check if it contains the "You've reached the end of the list." text
                return endMessageContainer.Text.Contains("You've reached the end of the list.");
            }
            catch (Exception ex)
            {
                LoggerService.Error("Crawling error - 478", ex);
                return false; // Keep scrolling if the element is not found
            }
        }
        private void InsertRowIntoDatatable(BusinessInfo dataTobeAdded, bool skipTempRow = false)
        {
            if(!skipTempRow) tempRows.Add(dataTobeAdded);
            if (skipTempRow)
            {
                var rowToBeAdded = updateGridList(dataTobeAdded);
                Invoke(new Action(() =>
                {
                    dataGridView.Rows.Add(rowToBeAdded);
                }));
                
                return;
            }
            if(tempRows.Count >= 10 || dataGridView.RowCount == 0)
            {
                Invoke(new Action(() =>

                {
                    foreach (var r in tempRows)
                    {
                        var rowToBeAdded = updateGridList(r);
                        dataGridView.Rows.Add(rowToBeAdded);
                    }
                    tempRows.Clear(); // Clear buffer after batch update
                }));
            }
        }

        private List<string> GetRatingReview(string input)
        {
            List<string> ratings = new List<string>();
            string pattern = @"(\d+(\.\d+)?)(\((\d+)\))?";

            // Apply regex match
            Match match = Regex.Match(input, pattern);

            if (match.Success)
            {
                string numberBeforeParentheses = match.Groups[1].Value; // The part before the parentheses (e.g., ৫.০)
                string numberInsideParentheses = match.Groups[4].Value; // The number inside the parentheses (e.g., ৯)
                ratings.Add(numberBeforeParentheses);
                ratings.Add(numberInsideParentheses);
                // If there's no number inside parentheses, we can set it to an empty string or handle it accordingly
                if (string.IsNullOrEmpty(numberInsideParentheses))
                {
                    ratings.Add("No reviews"); // You can define your own default value here
                }
            }
            else
            {
                ratings.Add("No reviews");
            }

            return ratings;
        }

        public (string Street, string ZipCode, string City, string Country) ParseAddress(string fullAddress)
        {
            string street = "";
            string zipCode = "";
            string city = "";
            string country = "";

            // ✅ Enhanced Regex to support various formats including Bangladesh
            Regex regex = new Regex(@"^(.*?),?\s*(\d{4,5})\s+([A-Za-zäöüÄÖÜß\s-]+),\s*([A-Za-z\s]+)$");
            Match match = regex.Match(fullAddress);

            if (match.Success)
            {
                street = match.Groups[1].Value.Trim();  // Extracts Street Name
                zipCode = match.Groups[2].Value.Trim();  // Extracts ZIP Code (first number)
                city = match.Groups[3].Value.Trim();    // Extracts City (text after zip)
                country = match.Groups[4].Value.Trim(); // Extracts Country (last part)
            }
            else
            {
                // ✅ Fallback: Check for country manually
                foreach (var countryName in CountryList)
                {
                    if (fullAddress.Contains(countryName))
                    {
                        country = countryName;
                        fullAddress = fullAddress.Replace(countryName, "").Trim(); // Remove country from address
                        break;
                    }
                }

                // ✅ Extract ZIP Code (first 4-5 digit number)
                Match zipMatch = Regex.Match(fullAddress, @"\b\d{4,5}\b");
                if (zipMatch.Success)
                {
                    zipCode = zipMatch.Value;
                    fullAddress = fullAddress.Replace(zipCode, "").Trim(); // Remove zip code from address
                }

                // ✅ Extract City (first part after ZIP, remove common words like "Division")
                string[] addressParts = fullAddress.Split(',');
                if (addressParts.Length > 1)
                {
                    city = addressParts[1].Trim();
                    street = addressParts[0].Trim(); // First part is street name
                }

                // ✅ Special Handling for Bangladesh
                if (country == "Bangladesh")
                {
                    city = ExtractBangladeshCity(fullAddress);
                }
            }

            return (street, zipCode, city, country);
        }

        private static string ExtractBangladeshCity(string address)
        {
            string[] commonDivisions = { "Dhaka", "Chattogram", "Khulna", "Rajshahi", "Barishal", "Sylhet", "Rangpur", "Mymensingh" };
            string city = "";

            foreach (string division in commonDivisions)
            {
                if (address.Contains(division))
                {
                    city = division;
                    break;
                }
            }

            if (string.IsNullOrEmpty(city))
            {
                // Fallback: Get city from the second part of the address
                string[] addressParts = address.Split(',');
                if (addressParts.Length > 1)
                {
                    city = addressParts[1].Trim();
                }
            }

            return city;
        }

        private ChromeDriver GetChromeDriverForBusinessDataFetch()
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            var options = new ChromeOptions();
            options.PageLoadStrategy = PageLoadStrategy.Eager;
            foreach (var item in chromeOptionArguements)
            {
                options.AddArgument(item);
            }

            options.AddUserProfilePreference("profile.managed_default_content_settings.images", 2); // Block images
            options.AddUserProfilePreference("profile.managed_default_content_settings.css", 2);    // Block CSS
            options.AddUserProfilePreference("profile.managed_default_content_settings.fonts", 2);  // Block fonts
            options.AddUserProfilePreference("profile.managed_default_content_settings.javascript", 2);  // Block fonts

            var driver = new ChromeDriver(service, options);
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            businessDriverProcessId = service.ProcessId;
            return driver;
        }

        public Dictionary<string, string> FetchSocialMediaLinks(string businessUrl, IWebDriver driver)
        {
            var socialLinks = new Dictionary<string, string>();
            if (businessUrl == null || businessUrl.Length == 0) { return socialLinks; }
            // Launch a new browser (this could be in a separate ChromeDriver instance)
            try
            {
                ((ChromeDriver)driver).ExecuteCdpCommand("Network.setBlockedURLs", new Dictionary<string, object>
            {
                { "urls", new[] { "*.jpg", "*.png", "*.gif", "*.css", "*.woff", "*.mp4", "*.svg" } }
            });
            }
            catch (ObjectDisposedException ex)
            {
                LoggerService.Error("Crawling error - 564", ex);
                //driver.Quit();
                //driver.Dispose();
                //driver = GetChromeDriverForBusinessDataFetch();
                //batchSize = 35;
            }
            try
            {
                //UpdateProgress("driver initialize finished");
                driver.Navigate().GoToUrl(businessUrl);
                //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                //wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("interactive"));
                //if (IsSessionActive(driver))
                //{
                //    driver.Navigate().GoToUrl(businessUrl);
                //} else
                //{
                //    LoggerService.Info("New driver creating for business data");
                //    driver = GetChromeDriverForBusinessDataFetch();
                //    Thread.Sleep(20);
                //    driver.Navigate().GoToUrl(businessUrl);
                //}
                //UpdateProgress("web visit finished");
                batchSize -= 1;
                // Fetch social media links from the business website (e.g., from footer or social media icons)
                bool isMailFetch = false;
                foreach (var selector in socialMediaSelectors)
                {
                    var elements = driver.FindElements(By.CssSelector(selector));
                    foreach (var element in elements)
                    {
                        var link = element.GetAttribute("href");
                        //var links = driver.FindElements(By.TagName("a"));
                        if (!string.IsNullOrEmpty(link) && (link.Contains("http://") || link.Contains("https://")))
                        {
                            bool isInvalidValid =
                                link == "https://www.linkedin.com/" ||
                                link == "https://www.facebook.com/" ||
                                link == "https://www.twitter.com/" ||
                                link == "https://instagram.com/" ||
                                link == "https://youtube.com/" ||
                                link == "https://pinterest.com/";
                            Uri uri = new Uri(link);
                            string host = uri.Host.Replace("www.", ""); // Remove "www."
                            string baseDomain = host.Split('.')[0];
                            if (isInvalidValid) link = string.Empty;
                            if (!isInvalidValid)
                            {
                                if (!socialLinks.ContainsKey(baseDomain.Replace("twitter", "x")))
                                {
                                    socialLinks.Add(baseDomain.Replace("twitter", "x"), link);
                                }
                            }

                        }
                        else if ((!string.IsNullOrEmpty(link) && (link.Contains("mailto"))))
                        {
                            if (!socialLinks.ContainsKey("emails"))
                            {
                                socialLinks.Add("emails", ProcessEmail(link));
                            }
                            else
                            {
                                string prevValue = socialLinks["emails"];
                                string currentValue = link.Replace("mailto:", "");
                                if (prevValue != currentValue)
                                {
                                    socialLinks["emails"] = prevValue + ", " + ProcessEmail(link);
                                }
                            }
                            isMailFetch = true;
                        }
                    }
                }
                if (!isMailFetch)
                {
                    string pageSource = driver.PageSource;
                    var emails = ExtractEmails(pageSource);

                    // Combine all emails into one string (comma-separated)
                    if (emails.Count > 0 && emails.Count <= 4)
                    {
                        string combinedEmails = string.Join(", ", emails);
                        socialLinks["emails"] = combinedEmails; // Single key for all emails
                    }
                    else
                    {
                        socialLinks["emails"] = string.Empty;
                    }
                    //socialLinks["emails"] = string.Empty;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching social media links for {businessUrl}: {ex.Message}");
                LoggerService.Error("Error fetching social media links for {businessUrl}", ex);
                //UpdateProgress($"Error fetching social media links for {businessUrl}: {ex.Message}" + "-583");
            }
            finally
            {
                //if(batchSize == 0)
                //{
                //    driver.Quit();
                //    driver.Dispose();
                //    driver = null;
                //}
            }
            //UpdateProgress("returning links");
            return socialLinks;
        }



        public static List<string> ExtractEmails(string pageSource)
        {
            List<string> emailList = new List<string>();

            // ✅ Regex pattern for extracting emails
            //string pattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";
            string pattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(?=\s|$|[?&])";
            //string pattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.(?!\d+x\d*\.(?:png|jpg|jpeg|gif|svg|webp))[a-zA-Z]{2,}(?=\s|$|[?&])";
            //string pattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.(?!\d+x\d*\.)[a-zA-Z]{2,}(?=\s|$|[?&])";
            //string pattern = @"\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.(?!\d+x\d+\.(?:png|jpg|jpeg|gif|svg|webp|bmp|tiff|ico|jfif))[a-zA-Z]{2,}\b";

            //string pattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.(?!\d+x\.(?:png|jpg|jpeg|gif|svg|webp))[a-zA-Z]{2,}";
            //string hrefPattern = @"<a\s+[^>]*href\s*=\s*""([^""]*@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})""";
            MatchCollection matches = Regex.Matches(pageSource, pattern);
            foreach (Match match in matches)
            {
                if (match.Value != null && match.Value.Length <= 30 && !emailList.Contains(match.Value))
                {
                    emailList.Add(match.Value); // Avoid duplicates
                }
            }
            return emailList;
        }

        public static List<string> FetchEmails(IWebDriver driver)
        {
            List<string> emails = new List<string>();
            var links = driver.FindElements(By.TagName("a"));
            string emailPattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";
            var elements = driver.FindElements(By.XPath("//*[text()]"));
            foreach (var element in elements)
            {
                try
                {
                    string text = element.Text;

                    if (!string.IsNullOrEmpty(text))
                    {
                        // Match all email-like patterns in the text
                        var matches = Regex.Matches(text, emailPattern);
                        foreach (Match match in matches)
                        {
                            string email = match.Value;

                            // Avoid duplicates
                            if (!emails.Contains(email))
                            {
                                emails.Add(email);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing element: {ex.Message}");
                    LoggerService.Error("Error processing element:", ex);
                    return emails;
                }
            }

            return emails;
        }

        string ExtractEmailFromDecodedValue(string decodedValue)
        {
            // Remove "mailto:" prefix
            if (decodedValue.StartsWith("mailto:"))
            {
                decodedValue = decodedValue.Substring("mailto:".Length);
            }

            // Replace custom encoding (e.g., [at], [dot]) with actual characters
            decodedValue = decodedValue.Replace("[at]", "@").Replace("[dot]", ".");

            // Remove additional tags like <obs>, <obs_c>, etc.
            decodedValue = Regex.Replace(decodedValue, @"<[^>]+>", "");

            return decodedValue.Trim();
        }

        bool IsEncoded(string value)
        {
            // Check for URL-encoded characters using a regex pattern
            return Regex.IsMatch(value, @"%[0-9A-Fa-f]{2}");
        }

        string ProcessEmail(string value)
        {
            // Check if the string contains URL-encoded characters
            if (IsEncoded(value))
            {
                // Decode the value
                value = HttpUtility.UrlDecode(value);
            }

            // Parse and clean the email
            return ExtractEmailFromDecodedValue(value);
        }



        private bool ScrollToLoadMoreResults(IWebDriver driver, IWebElement mapContainer, HashSet<string> processedResults, IJavaScriptExecutor jsExecutor)
        {
            //ValidateDriverSession(driver, jsExecutor);
            bool hasMoreResults = true;
            int newHeight = Convert.ToInt32(jsExecutor.ExecuteScript("return arguments[0].scrollHeight", mapContainer));
            var newElements = driver.FindElements(By.XPath("//div[@class='bfdHYd Ppzolf OFBs3e  ']"));
            List<string> t = newElements.Select(x =>
            {
                string txt = x.FindElement(By.CssSelector(".qBF1Pd.fontHeadlineSmall")).Text;
                if (!processedResults.Contains(txt))
                {
                    return txt;
                }
                else
                {
                    return string.Empty;
                }

            }).ToList();
            t = t.Where(x => x.Length > 0).ToList();
            hasMoreResults = t.Count != 0;

            return hasMoreResults;
        }

        private Dictionary<string, int> columnIndexCache;
        public void InitializeColumnIndexCache(DataGridView dataGridView)
        {
            columnIndexCache = new Dictionary<string, int>();

            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                columnIndexCache[column.HeaderText] = column.Index; // Store column index by name
            }
        }

        public int GetColumnIndex(string columnHeader)
        {
            if (columnIndexCache != null && columnIndexCache.TryGetValue(columnHeader, out int index))
            {
                return index; // Fast O(1) lookup
            }
            return -1; // Not found
        }


        private void ExportToExcel(List<BusinessInfo> results)
        {
            if (results == null)
            {
                MessageBox.Show("There is no data no export.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                var confirmationForm = new ConfirmationForm("Do you want to remove the duplicate data?");
                var result = confirmationForm.ShowDialog();
                bool removeDuplicate = result == DialogResult.OK && confirmationForm.UserConfirmed;

                var confirmationCsvForm = new ConfirmationForm("Do you want to export as csv?");
                var resultExportType = confirmationCsvForm.ShowDialog();
                bool exportCsv = resultExportType == DialogResult.OK && confirmationCsvForm.UserConfirmed;

                DateTime now = DateTime.Now;
                string formattedDate = now.ToString("dd MMM yyyy, HH:mm");
                List<(string ResultTitle, string Address, string ContactNumber)> duplicateDataFlag = new List<(string ResultTitle, string Address, string ContactNumber)>();
                if (exportCsv)
                {
                    //csv header
                    StringBuilder csvContent = new StringBuilder();
                    for (int i = 0; i < dataGridView.Columns.Count; i++)
                    {
                        csvContent.Append(dataGridView.Columns[i].HeaderText);
                        csvContent.Append(",");
                    }
                    csvContent.Append("Date time");
                    csvContent.AppendLine();
                    if (removeDuplicate)
                    {
                        InitializeColumnIndexCache(dataGridView);
                    }
                    // Add row data
                    foreach (DataGridViewRow row in dataGridView.Rows)
                    {
                        if (!row.IsNewRow) // Ignore empty new row
                        {
                            if (removeDuplicate)
                            {
                                int titleIndex = GetColumnIndex("Name");
                                int numberIndex = GetColumnIndex("Contact Number");
                                int addressIndex = GetColumnIndex("Full_Address");

                                var hasDuplicate = duplicateDataFlag.Any(x =>
                                       x.ResultTitle.ToLower() == row.Cells[titleIndex].Value?.ToString().ToLower() &&
                                       x.ContactNumber.ToLower() == row.Cells[numberIndex].Value?.ToString().ToLower() &&
                                       x.Address.ToLower() == row.Cells[addressIndex].Value?.ToString().ToLower()
                                   );
                                if (hasDuplicate && titleIndex != -1 && numberIndex != -1 && addressIndex != -1)
                                {
                                    continue;
                                }
                                duplicateDataFlag.Add((row.Cells[titleIndex].Value?.ToString(), row.Cells[numberIndex].Value?.ToString(), row.Cells[addressIndex].Value?.ToString()));
                            }
                            for (int i = 0; i < dataGridView.Columns.Count; i++)
                            {
                                csvContent.Append(row.Cells[i].Value?.ToString().Replace(",", " ")); // Remove extra commas
                                csvContent.Append(",");
                            }
                            csvContent.Append(formattedDate);
                            csvContent.AppendLine();
                        }
                    }

                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
                    saveFileDialog.DefaultExt = "csv"; // Default extension

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;

                        File.WriteAllText(filePath, csvContent.ToString(), Encoding.UTF8);

                        MessageBox.Show("File exported successfully as CSV!", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Results");
                        int headerIndex = 1;
                        for (headerIndex = 1; headerIndex < SharedDataTableModel.SelectedFields.Count; headerIndex++)
                        {
                            worksheet.Cell(1, headerIndex).Value = SharedDataTableModel.SelectedFields[headerIndex].Name;
                        }
                        worksheet.Cell(1, headerIndex).Value = "Date time";
                        List<String> sheetColumnsValue = new List<String>();
                        for (int i = 0; i < results.Count; i++)
                        {
                            if (removeDuplicate)
                            {
                                var hasDuplicate = duplicateDataFlag.Any(x =>
                                    x.ResultTitle.ToLower() == results[i].ResultTitle.ToLower() &&
                                    x.ContactNumber.ToLower() == results[i].ContactNumber.ToLower() &&
                                    x.Address.ToLower() == results[i].Address.ToLower()
                                );
                                if (hasDuplicate)
                                {
                                    continue;
                                }
                                duplicateDataFlag.Add((results[i].ResultTitle, results[i].Address, results[i].ContactNumber));
                            }

                            int columnIndex = 1;
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Keyword") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].SearchTerm;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Name") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].ResultTitle;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Category") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].Category;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Full_Address") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].Address;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Street_Address") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].StreetAddress;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "City") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].City;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Zip") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].Zip;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Country") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].Country;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Contact Number") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].ContactNumber;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Email") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].SocialMedias.ContainsKey("emails") ? results[i].SocialMedias["emails"] : string.Empty;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Website") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].CompanyWebsite;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Facebook") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].SocialMedias.ContainsKey("facebook") ? results[i].SocialMedias["facebook"] : string.Empty;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Linkedin") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].SocialMedias.ContainsKey("linkedin") ? results[i].SocialMedias["linkedin"] : string.Empty;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Twitter") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].SocialMedias.ContainsKey("x") ? results[i].SocialMedias["x"] : string.Empty;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Youtube") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].SocialMedias.ContainsKey("youtube") ? results[i].SocialMedias["youtube"] : string.Empty;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Instagram") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].SocialMedias.ContainsKey("instagram") ? results[i].SocialMedias["instagram"] : string.Empty;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Pinterest") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].SocialMedias.ContainsKey("pinterest") ? results[i].SocialMedias["pinterest"] : string.Empty;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Rating") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].Rating;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Review Count") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].ReviewCount;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Claim") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].Claim;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Hours_Info") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].HoursInfo;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Monday") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].BusinessHours.ContainsKey("Monday") ? results[i].BusinessHours["Monday"] : string.Empty;
                                columnIndex++;
                            } if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Tuesday") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].BusinessHours.ContainsKey("Tuesday") ? results[i].BusinessHours["Monday"] : string.Empty;
                                columnIndex++;
                            } if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Wednesday") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].BusinessHours.ContainsKey("Wednesday") ? results[i].BusinessHours["Wednesday"] : string.Empty;
                                columnIndex++;
                            } if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Thursday") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].BusinessHours.ContainsKey("Thursday") ? results[i].BusinessHours["Thursday"] : string.Empty;
                                columnIndex++;
                            } if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Friday") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].BusinessHours.ContainsKey("Friday") ? results[i].BusinessHours["Friday"] : string.Empty;
                                columnIndex++;
                            } if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Saturday") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].BusinessHours.ContainsKey("Saturday") ? results[i].BusinessHours["Saturday"] : string.Empty;
                                columnIndex++;
                            } if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Sunday") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].BusinessHours.ContainsKey("Sunday") ? results[i].BusinessHours["Sunday"] : string.Empty;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Located_in") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].LocatedIn;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Attributes") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].Attributes;
                                columnIndex++;
                            }
                            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Map_Link") != null)
                            {
                                worksheet.Cell(i + 2, columnIndex).Value = results[i].MapLink;
                                columnIndex++;
                            }
                            worksheet.Cell(i + 2, columnIndex).Value = formattedDate;
                            columnIndex++;

                        }
                        SaveFileDialog saveFileDialog = new SaveFileDialog();
                        // Set filter for Excel files
                        saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                        saveFileDialog.DefaultExt = "xlsx";  // Default extension

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string filePath = saveFileDialog.FileName;
                            workbook.SaveAs(filePath);
                            MessageBox.Show("File exported successfully!", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
               
            }
            catch (Exception ex)
            {
                LoggerService.Error($"Error exporting to Excel: {ex.Message}", ex);
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool IsSessionActive(IWebDriver driver)
        {
            try
            {
                if (driver == null) return false;
                var sessionProperty = driver.GetType().GetProperty("SessionId");
                if (sessionProperty != null)
                {
                    var sessionId = sessionProperty.GetValue(driver);
                    if (sessionId != null)
                    {
                        //LoggerService.Info($"WebDriver Session ID: {sessionId}");
                        return true; // Session is active
                    }
                }
                driver?.Quit();
                driver?.Dispose();
                driver = null;
                return false; // No valid session
            }
            catch (Exception)
            {
                driver?.Quit();
                driver?.Dispose();
                driver = null;
                return false; // Session is not active
            }
        }

        public Dictionary<string, string> ConvertBusinessHours(string input)
        {
            Dictionary<string, string> hoursDict = new Dictionary<string, string>();

            // ✅ Regular expression to match each day and time range
            string pattern = @"(\b(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b),\s*([^;]+)";
            MatchCollection matches = Regex.Matches(input, pattern);

            foreach (Match match in matches)
            {
                string day = match.Groups[1].Value.Trim();
                string timeRange = match.Groups[2].Value.Trim();

                // ✅ If the business is "Closed", keep it as "Closed"
                if (timeRange.Equals("Closed", StringComparison.OrdinalIgnoreCase))
                {
                    hoursDict[day] = "Closed";
                }
                else
                {
                    // ✅ Convert original time range to "10 AM–3 AM"
                    hoursDict[day] = "10 AM–3 AM";
                }
            }

            return hoursDict;
        }


        private Object[] updateGridList(BusinessInfo result)
        {
            dataColumns = new List<string>();
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "#") != null)
            {
                dataColumns.Add((dataGridView.Rows.Count + 1).ToString());
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Keyword") != null)
            {
                dataColumns.Add(result.SearchTerm);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Name") != null)
            {
                dataColumns.Add(result.ResultTitle);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Category") != null)
            {
                dataColumns.Add(result.Category);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Full_Address") != null)
            {
                dataColumns.Add(result.Address);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Street_Address") != null)
            {
                dataColumns.Add(result.StreetAddress);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "City") != null)
            {
                dataColumns.Add(result.City);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Zip") != null)
            {
                dataColumns.Add(result.Zip);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Country") != null)
            {
                dataColumns.Add(result.Country);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Contact Number") != null)
            {
                dataColumns.Add(result.ContactNumber);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Email") != null)
            {
                dataColumns.Add(result.SocialMedias.ContainsKey("emails") ? result.SocialMedias["emails"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Website") != null)
            {
                dataColumns.Add(result.CompanyWebsite);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Facebook") != null)
            {
                dataColumns.Add(result.SocialMedias.ContainsKey("facebook") ? result.SocialMedias["facebook"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Linkedin") != null)
            {           
                dataColumns.Add(result.SocialMedias.ContainsKey("linkedin") ? result.SocialMedias["linkedin"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Twitter") != null)
            {
                dataColumns.Add(result.SocialMedias.ContainsKey("x") ? result.SocialMedias["x"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Youtube") != null)
            {
                dataColumns.Add(result.SocialMedias.ContainsKey("youtube") ? result.SocialMedias["youtube"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Instagram") != null)
            {
                dataColumns.Add(result.SocialMedias.ContainsKey("instagram") ? result.SocialMedias["instagram"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Pinterest") != null)
            {
                dataColumns.Add(result.SocialMedias.ContainsKey("pinterest") ? result.SocialMedias["pinterest"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Rating") != null)
            {
                dataColumns.Add(result.Rating);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Review Count") != null)
            {
                dataColumns.Add(result.ReviewCount);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Claim") != null)
            {
                dataColumns.Add(result.Claim.ToString());
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Hours_Info") != null)
            {
                dataColumns.Add(result.HoursInfo);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Monday") != null)
            {
                dataColumns.Add(result.BusinessHours.ContainsKey("Monday") ? result.BusinessHours["Monday"] : string.Empty);
            }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Tuesday") != null)
            {
                dataColumns.Add(result.BusinessHours.ContainsKey("Tuesday") ? result.BusinessHours["Tuesday"] : string.Empty);
            }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Wednesday") != null)
            {
                dataColumns.Add(result.BusinessHours.ContainsKey("Wednesday") ? result.BusinessHours["Wednesday"] : string.Empty);
            }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Thursday") != null)
            {
                dataColumns.Add(result.BusinessHours.ContainsKey("Thursday") ? result.BusinessHours["Thursday"] : string.Empty);
            }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Friday") != null)
            {
                dataColumns.Add(result.BusinessHours.ContainsKey("Friday") ? result.BusinessHours["Friday"] : string.Empty);
            }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Saturday") != null)
            {
                dataColumns.Add(result.BusinessHours.ContainsKey("Saturday") ? result.BusinessHours["Saturday"] : string.Empty);
            }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Sunday") != null)
            {
                dataColumns.Add(result.BusinessHours.ContainsKey("Sunday") ? result.BusinessHours["Sunday"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Located_in") != null)
            {
                dataColumns.Add(result.LocatedIn);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Attributes") != null)
            {
                dataColumns.Add(result.Attributes);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Map_Link") != null)
            {
                dataColumns.Add(result.MapLink);
            }

            return dataColumns.ToArray();


        }
    }
}
