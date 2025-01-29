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
using DocumentFormat.OpenXml.Spreadsheet;
using System.Web;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        List<string> dataColumns = new List<string>();
        Dictionary<string, (string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string StreetAddress, string city, string zip, string country, Dictionary<string, string> socialMedias, string companyWebsite)> uniqueDataPair = new Dictionary<string, (string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string StreetAddress, string city, string zip, string country, Dictionary<string, string> socialMedias, string companyWebsite)>();
        private string[] searchTerms;
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
        private List<(string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string StreetAddress, string city, string zip, string country, Dictionary<string, string> socialMedias, string companyWebsite)> results;
        public Form1()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ResizeEnd += Form1_Resize;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AdjustDataGridViewHeight();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            AdjustDataGridViewHeight();
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
                //progressBar.Visible = false;
            }
        }

        private void exportDataButton_Click(object sender, EventArgs e)
        {
            ExportToExcel(results);
        }
        
        private void clearDataButton_Click(object sender, EventArgs e)
        {
            dataGridView.Rows.Clear();
            UpdateProgress("", true);
        }

        private async Task StartCrawlingAsync(CancellationToken cancellationToken, bool fetchBusinessData)
        {
            try
            {
                uniqueDataPair = new Dictionary<string, (string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string StreetAddress, string city, string zip, string country, Dictionary<string, string> socialMedias, string companyWebsite)>();
                new DriverManager().SetUpDriver(new ChromeConfig()); // Automatically downloads ChromeDriver
                ChromeOptions options = new ChromeOptions();
                IWebDriver driver = new ChromeDriver(options);
                driver.Manage().Window.Maximize();
                IWebDriver chromeDriverForBusinessData = GetChromeDriverForBusinessDataFetch();
                driver.Navigate().GoToUrl("https://www.google.com/maps?hl=en");
                foreach (string term in searchTerms)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        //lblStatus.Text = "Crawling stopped.";
                        break;
                    }
                    try
                    {
                        UpdateProgress($"Crawling: {term} - In progress");
                        await processCrawlingAsync(driver, chromeDriverForBusinessData, term, cancellationToken, fetchBusinessData);
                        UpdateProgress($"Crawling: {term} - Completed");
                    } catch(Exception exc)
                    {
                        Console.WriteLine(exc.Message);
                    }
                }
                animatedLoader.Invoke(new Action(() =>
                {
                    animatedLoader.Visible = false;
                }));
                driver.Quit();
                chromeDriverForBusinessData.Quit();
                // Export results to Excel
                //ExportToExcel(results);
                UpdateProgress("Crawling finished.");
                //progressBar.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during crawling: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateProgress("Error occurred.");
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

        private async Task processCrawlingAsync(IWebDriver driver, IWebDriver webDriver, string term, CancellationToken cancellationToken, bool fetchBusinessData)
        {
            var searchBox = driver.FindElement(By.Id("searchboxinput"));
            searchBox.Clear();
            searchBox.SendKeys(term);
            searchBox.SendKeys(Keys.Enter);
            // Wait for the search box to load
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
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
                            var linkElement = resultElement.FindElements(By.ClassName("lcr4fd"));
                            string hrefValue = string.Empty;

                            processedResults.Add(title);
                            var clickableElement = driver.FindElements(By.CssSelector("a.hfpxzc"));
                            // ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", clickableElement[i]);
                            int height = clickableElement[i].Size.Height;

                            lastHeight = Convert.ToInt32(jsExecutor.ExecuteScript("return arguments[0].scrollHeight", mapContainer));
                            jsExecutor.ExecuteScript("arguments[0].scrollBy(0, arguments[1]);", mapContainer, height);
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", clickableElement[i]);
                            clickableElement[i].Click();
                            i++;
                            string contactNumber = string.Empty;
                            try
                            {
                                contactNumber = resultElement.FindElement(By.ClassName("UsdlK")).Text;
                            }
                            catch (Exception ex)
                            {
                            }
                            UpdateProgress($"Record no: {(dataGridView.Rows.Count + 1).ToString()}");
                            UpdateProgress($"Title: {title}");
                            UpdateProgress($"------------------------");
                            var existedData = uniqueDataPair.ContainsKey($"{title}_{contactNumber}");
                            if (existedData)
                            {
                                var existingData = uniqueDataPair[$"{title}_{contactNumber}"];
                                results.Add(existingData);
                                InsertRowIntoDatatable(existingData);
                                hasUrl = false;
                                Thread.Sleep(50);
                                continue;
                            }

                            if (fetchBusinessData && linkElement.Count > 0)
                            {
                                hasUrl = true;
                                hrefValue = linkElement[0].GetDomAttribute("href");
                                var task = Task.Run(() => FetchSocialMediaLinks(hrefValue, webDriver));
                                keyValuePairs = await task;
                            }
                            else
                            {
                                hasUrl = false;
                                Thread.Sleep(200);
                            }
                            var reviewListText = GetRatingReview(resultElement.FindElement(By.ClassName("W4Efsd")).Text);
                            if (reviewListText != null && reviewListText.Count > 1)
                            {
                                reviewCount = reviewListText[0];
                                rating = reviewListText[1];
                            }
                            string category = string.Empty;
                            string city = string.Empty;
                            string zip = string.Empty;
                            string country = string.Empty;
                            string location = string.Empty;
                            string streetLocation = string.Empty;
                            string pattern = @"(?<street>[\d\s\w\W]+),\s(?<city>[A-Za-z\s]+)\s(?<state>[A-Za-z]+)\s(?<zip>\d{4}),\s(?<country>[A-Za-z\s]+)$";
                            var elements = resultElement.FindElements(By.CssSelector(".UaQhfb.fontBodyMedium > .W4Efsd")).Last();
                            var detailsElems = driver.FindElements(By.XPath("//div[@class='m6QErb DxyBCb kA9KIf dS8AEf XiKgde ']"));
                            try
                            {
                                var item = detailsElems[0].FindElement(By.CssSelector("button[data-item-id='address'] .Io6YTe.fontBodyMedium.kR99db.fdkmkc"));
                                location = item != null ? item.Text : string.Empty;
                                var regex = new Regex(pattern);
                                var match = regex.Match(location);
                                if (match.Success)
                                {
                                    streetLocation = match.Groups["street"].Value;
                                    city = match.Groups["city"].Value;
                                    zip = match.Groups["zip"].Value;
                                    country = match.Groups["country"].Value;
                                }
                                else
                                {
                                    streetLocation = elements.FindElements(By.CssSelector(":nth-child(2)")).First().Text;
                                }
                            }
                            catch (NoSuchElementException)
                            {
                                location = string.Empty;
                            }
                            //var locationElems = driver.FindElements(By.XPath("//div[@class='Io6YTe fontBodyMedium kR99db fdkmkc']"));
                            //string location = locationElems != null && locationElems.Count > 0 ? locationElems[0].Text : string.Empty;
                            if (elements != null)
                            {
                                category = elements.FindElements(By.CssSelector(":first-child")).First().Text;
                                if (elements != null && category.Length > 0)
                                {
                                    category = category.Split('·')[0];
                                }
                            }
                            var dataTobeAdded = (term, title, reviewCount, rating, contactNumber, category, location, streetLocation, city, zip, country, keyValuePairs, hrefValue);
                            results.Add(dataTobeAdded);
                            uniqueDataPair.Add($"{title}_{contactNumber}", dataTobeAdded);
                            InsertRowIntoDatatable(dataTobeAdded);
                            //driver.Navigate().Back();
                            //driver.Navigate().Back();

                        }
                        catch (Exception ex)
                        {
                            // Skip if any element is missing
                        }
                    }
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
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
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Timeout waiting for search results to load.");
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
            catch (NoSuchElementException)
            {
                return false; // Keep scrolling if the element is not found
            }
        }

        private void InsertRowIntoDatatable((string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string StreetAddress, string city, string zip, string country, Dictionary<string, string> socialMedias, string companyWebsite) dataTobeAdded)
        {
            var rowToBeAdded = updateGridList(dataTobeAdded);
            Invoke(new Action(() =>

            {
                dataGridView.Rows.Add(rowToBeAdded);
            }));
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
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);

            return driver;
        }

        public Dictionary<string, string> FetchSocialMediaLinks(string businessUrl, IWebDriver driver)
        {
            var socialLinks = new Dictionary<string, string>();
            if (businessUrl == null || businessUrl.Length == 0) { return socialLinks; }
            // Launch a new browser (this could be in a separate ChromeDriver instance)
            
            ((ChromeDriver)driver).ExecuteCdpCommand("Network.setBlockedURLs", new Dictionary<string, object>
            {
                { "urls", new[] { "*.jpg", "*.png", "*.gif", "*.css", "*.woff", "*.mp4", "*.svg" } }
            });
            try
            {
                driver.Navigate().GoToUrl(businessUrl);

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
                    var emails = FetchEmails(driver);

                    // Combine all emails into one string (comma-separated)
                    if (emails.Count > 0)
                    {
                        string combinedEmails = string.Join(", ", emails);
                        socialLinks["emails"] = combinedEmails; // Single key for all emails
                    }
                    else
                    {
                        socialLinks["emails"] = string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching social media links for {businessUrl}: {ex.Message}");
            }
            finally
            {
                //driver.Quit();
            }

            return socialLinks;
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

        private void ExportToExcel(List<(string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string streetAddress, string city, string zip, string country, Dictionary<string, string> socialMedias, string hrefValue)> results)
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
                List<(string ResultTitle, string Address, string ContactNumber)> duplicateDataFlag = new List<(string ResultTitle, string Address, string ContactNumber)>();
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
                    DateTime now = DateTime.Now;

                    // Format it as "24 Jan 2025, hh:mm"
                    string formattedDate = now.ToString("dd MMM yyyy, HH:mm");
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
                        } if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Full_Address") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].Address;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Street_Address") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].streetAddress;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "City") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].city;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Zip") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].zip;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Country") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].country;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Contact Number") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].ContactNumber;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Email") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].socialMedias.ContainsKey("emails") ? results[i].socialMedias["emails"] : string.Empty;
                            columnIndex++;
                        }
                        if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Website") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].hrefValue;
                            columnIndex++;
                        }
                        if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Facebook") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].socialMedias.ContainsKey("facebook") ? results[i].socialMedias["facebook"] : string.Empty;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Linkedin") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].socialMedias.ContainsKey("linkedin") ? results[i].socialMedias["linkedin"] : string.Empty;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Twitter") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].socialMedias.ContainsKey("x") ? results[i].socialMedias["x"] : string.Empty;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Youtube") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].socialMedias.ContainsKey("youtube") ? results[i].socialMedias["youtube"] : string.Empty;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Instagram") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].socialMedias.ContainsKey("instagram") ? results[i].socialMedias["instagram"] : string.Empty;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Pinterest") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].socialMedias.ContainsKey("pinterest") ? results[i].socialMedias["pinterest"] : string.Empty;
                            columnIndex++;
                        }if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Rating") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].Rating;
                            columnIndex++;
                        }
                        if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Review Count") != null)
                        {
                            worksheet.Cell(i + 2, columnIndex).Value = results[i].ReviewCount;
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
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Object[] updateGridList((string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string StreetAddress, string city, string zip, string country, Dictionary<string, string> socialMedias, string companyWebsite) result)
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
                dataColumns.Add(result.city);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Zip") != null)
            {
                dataColumns.Add(result.zip);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Country") != null)
            {
                dataColumns.Add(result.country);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Contact Number") != null)
            {
                dataColumns.Add(result.ContactNumber);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Email") != null)
            {
                dataColumns.Add(result.socialMedias.ContainsKey("emails") ? result.socialMedias["emails"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Website") != null)
            {
                dataColumns.Add(result.companyWebsite);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Facebook") != null)
            {
                dataColumns.Add(result.socialMedias.ContainsKey("facebook") ? result.socialMedias["facebook"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Linkedin") != null)
            {
                dataColumns.Add(result.socialMedias.ContainsKey("linkedin") ? result.socialMedias["linkedin"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Twitter") != null)
            {
                dataColumns.Add(result.socialMedias.ContainsKey("x") ? result.socialMedias["x"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Youtube") != null)
            {
                dataColumns.Add(result.socialMedias.ContainsKey("youtube") ? result.socialMedias["youtube"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Instagram") != null)
            {
                dataColumns.Add(result.socialMedias.ContainsKey("instagram") ? result.socialMedias["instagram"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Pinterest") != null)
            {
                dataColumns.Add(result.socialMedias.ContainsKey("pinterest") ? result.socialMedias["pinterest"] : string.Empty);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Rating") != null)
            {
                dataColumns.Add(result.Rating);
            }
            if (SharedDataTableModel.SelectedFields.Find(x => x.Name == "Review Count") != null)
            {
                dataColumns.Add(result.ReviewCount);
            }

            return dataColumns.ToArray();


        }
    }
}
