using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using Keys = OpenQA.Selenium.Keys;

namespace WindowsFormsApp2
{
    partial class Form1 : Form
    {
        private string[] searchTerms;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.Button btnStart;
        //private System.Windows.Forms.DataGridView dataGridView = new System.Windows.Forms.DataGridView();
        //private DataTable resultsDataTable = new DataTable();
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnUpload = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            //((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();

            //this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            //this.dataGridView.Location = new System.Drawing.Point(12, 12);
            //this.dataGridView.Name = "dataGridView";
            //this.dataGridView.RowTemplate.Height = 24;
            //this.dataGridView.Size = new System.Drawing.Size(760, 400);
            //this.dataGridView.TabIndex = 0;


            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 50); // Position it appropriately
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(70, 13); // Adjust size as needed
            this.lblStatus.TabIndex = 2;
            this.lblStatus.Text = "Status: Idle";

            this.btnUpload.Location = new System.Drawing.Point(20, 20); // Set position
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(150, 40); // Set size
            this.btnUpload.TabIndex = 0;
            this.btnUpload.Text = "Upload File";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);

            this.btnStart.Location = new System.Drawing.Point(200, 20); // Set position
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(150, 40); // Set size
            this.btnStart.TabIndex = 1;
            this.btnStart.Text = "Start Crawling";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 200); // Adjust size as needed
            //this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.btnStart);
            this.Name = "MainForm";
            this.Text = "Google Maps Crawler";
            //((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        //private void InitializeDataTable()
        //{
        //    // Define the DataTable with columns
        //    resultsDataTable = new DataTable();
        //    resultsDataTable.Columns.Add("Keyword");
        //    resultsDataTable.Columns.Add("Result Title");
        //    resultsDataTable.Columns.Add("Category");
        //    resultsDataTable.Columns.Add("Location");
        //    resultsDataTable.Columns.Add("Contact Number");
        //    resultsDataTable.Columns.Add("Email");
        //    resultsDataTable.Columns.Add("Website");
        //    resultsDataTable.Columns.Add("Facebook");
        //    resultsDataTable.Columns.Add("Linkedin");
        //    resultsDataTable.Columns.Add("Twitter");
        //    resultsDataTable.Columns.Add("Youtube");
        //    resultsDataTable.Columns.Add("Instagram");
        //    resultsDataTable.Columns.Add("Pinterest");
        //    resultsDataTable.Columns.Add("Ratiung");
        //    resultsDataTable.Columns.Add("Review Count");

        //    // Bind the DataTable to the DataGridView
        //    dataGridView.DataSource = resultsDataTable;
        //}

        private void btnUpload_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Text Files (*.txt)|*.txt";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        searchTerms = File.ReadAllLines(openFileDialog.FileName).Select(line => line.Trim()).ToArray();
                        lblStatus.Text = $"Loaded {searchTerms.Length} search terms.";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error loading file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (searchTerms == null || searchTerms.Length == 0)
            {
                MessageBox.Show("Please upload a file with search terms first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            lblStatus.Text = "Crawling started...";
            //InitializeDataTable();
            new Thread(async () => await StartCrawlingAsync()).Start();
        }


        private async Task StartCrawlingAsync()
        {
            try
            {
                var results = new System.Collections.Generic.List<(string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, Dictionary<string, string>, string companyWebsite)>();

                new DriverManager().SetUpDriver(new ChromeConfig()); // Automatically downloads ChromeDriver
                ChromeOptions options = new ChromeOptions();
                IWebDriver driver = new ChromeDriver(options);
                //ChromeOptions options = new ChromeOptions();
                //string driverPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
                //IWebDriver driver = new ChromeDriver(driverPath);
                //options.AddArgument("C:\\Users\\Skytech\\Downloads\\chrome-win64\\chrome-win64");
                //IWebDriver driver = new ChromeDriver(options);

                driver.Navigate().GoToUrl("https://www.google.com/maps");

                foreach (string term in searchTerms)
                {
                    // Wait for the search box to load
                    Thread.Sleep(2000);
                    var searchBox = driver.FindElement(By.Id("searchboxinput"));
                    searchBox.Clear();
                    searchBox.SendKeys(term);
                    searchBox.SendKeys(Keys.Enter);

                    Thread.Sleep(5000); // Wait for search results to load

                    var mapContainer = driver.FindElement(By.XPath("//div[@class='m6QErb DxyBCb kA9KIf dS8AEf XiKgde ecceSd']"));
                    var processedResults = new HashSet<string>();
                    bool hasMoreResults = true;
                    IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
                    int lastHeight = 0;
                    while (hasMoreResults)
                    {
                        // Extract search results
                        var resultElements = driver.FindElements(By.XPath("//div[@class='bfdHYd Ppzolf OFBs3e  ']"));
                        int i = 0;
                        ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                        service.HideCommandPromptWindow = true;
                        foreach (var resultElement in resultElements)
                        {
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
                                if (linkElement.Count > 0)
                                {
                                    hrefValue = linkElement[0].GetAttribute("href");
                                    var task = Task.Run(() => FetchSocialMediaLinks(hrefValue, service));
                                    keyValuePairs = await task;
                                }
                                processedResults.Add(title);
                                var clickableElement = driver.FindElements(By.CssSelector("a.hfpxzc"));
                               // ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", clickableElement[i]);
                                int height = clickableElement[i].Size.Height;
                                //((IJavaScriptExecutor)driver).ExecuteScript($"arguments[0].scrollTop(0, {height});", mapContainer);
                                //((IJavaScriptExecutor)driver).ExecuteScript($"arguments[0].scrollTop = {height}", mapContainer);
                                //((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", clickableElement[i]);
                                
                                lastHeight = Convert.ToInt32(jsExecutor.ExecuteScript("return arguments[0].scrollHeight", mapContainer));
                                jsExecutor.ExecuteScript("arguments[0].scrollBy(0, arguments[1]);", mapContainer, height);
                                clickableElement[i].Click();
                                i++;
                                Thread.Sleep(3000);
                                var reviewListText = GetRatingReview(resultElement.FindElement(By.ClassName("W4Efsd")).Text);
                                if (reviewListText != null && reviewListText.Count > 1)
                                {
                                    reviewCount = reviewListText[0];
                                    rating = reviewListText[1];
                                }
                                string category = string.Empty;
                                string location = string.Empty;
                                var elements = resultElement.FindElements(By.CssSelector(".UaQhfb.fontBodyMedium > .W4Efsd")).Last();
                                if (elements != null)
                                {
                                    category = elements.FindElements(By.CssSelector(":first-child")).First().Text;
                                    location = elements.FindElements(By.CssSelector(":nth-child(2)")).First().Text;
                                    if (elements != null && category.Length > 0)
                                    {
                                        category = category.Split('.')[0];
                                    }
                                }
                                string contactNumber = string.Empty;
                                try
                                {
                                    contactNumber = resultElement.FindElement(By.ClassName("UsdlK")).Text;
                                }
                                catch (Exception ex)
                                {
                                }

                                results.Add((term, title, reviewCount, rating, contactNumber, category, location, keyValuePairs, hrefValue));
                                //Invoke(new Action(() =>
                                //{
                                //    // Add a new row with the extracted data
                                //    resultsDataTable.Rows.Add(term, title, reviewCount, rating, contactNumber, category, location, "", "", "", "", "", "", hrefValue);
                                //}));
                                driver.Navigate().Back();
                                Thread.Sleep(2000);
                            }
                            catch (Exception ex)
                            {
                                // Skip if any element is missing
                            }
                        }

                        //double lastHeight = Convert.ToDouble(((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].scrollHeight", mapContainer));
                        int newHeight = Convert.ToInt32(jsExecutor.ExecuteScript("return arguments[0].scrollHeight", mapContainer));
                        var newElements = driver.FindElements(By.XPath("//div[@class='bfdHYd Ppzolf OFBs3e  ']"));
                        var t = newElements.Select(x =>
                        {
                            string txt = x.FindElement(By.CssSelector(".qBF1Pd.fontHeadlineSmall")).Text;
                            if (!processedResults.Contains(txt))
                            {
                                return txt;
                            } else
                            {
                                return string.Empty;
                            }
                            
                        }).ToList();
                        t = t.Where(x => x.Length > 0).ToList();
                        if (t.Count == 0)
                        {
                            hasMoreResults = false;
                        }
                        //// Scroll further down
                        //((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollBy(0, 500);", mapContainer);
                        //Thread.Sleep(3000);
                        //// Check if more results are loaded (adjust this condition as per your application's behavior)
                        //bool isEndOfResults = !mapContainer.FindElements(By.XPath(".//div[@class='bfdHYd Ppzolf OFBs3e  ']")).Any();
                        //if (isEndOfResults)
                        //{
                        //    hasMoreResults = false;
                        //}
                        ////hasMoreResults = ScrollToLoadMoreResults(driver, mapContainer);
                    }

                    
                }

                driver.Quit();

                // Export results to Excel
                ExportToExcel(results);

                Invoke(new Action(() => lblStatus.Text = "Crawling finished. Results exported."));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during crawling: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Invoke(new Action(() => lblStatus.Text = "Error occurred."));
            }
        }

    //    private string GetRandomUserAgent()
    //    {
    //        var userAgents = new List<string>
    //{
    //    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    //    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36",
    //    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/89.0"
    //};

    //        Random rand = new Random();
    //        return userAgents[rand.Next(userAgents.Count)];
    //    }

        private List<string> GetRatingReview(string input)
        {
            List<string>  ratings = new List<string>();
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

        public Dictionary<string, string> FetchSocialMediaLinks(string businessUrl, ChromeDriverService service)
        {
            var socialLinks = new Dictionary<string, string>();
            if (businessUrl == null || businessUrl.Length == 0) {  return socialLinks; }
            // Launch a new browser (this could be in a separate ChromeDriver instance)
            var options = new ChromeOptions();
            options.AddArguments("headless", "disable-gpu", "no-sandbox");
            options.AddArguments("--blink-settings=imagesEnabled=false");  // Disable images
            options.AddArguments("--disable-css");  // Disable CSS
            options.AddArguments("--disable-javascript");  // Disable JavaScript
            

            var driver = new ChromeDriver(service, options);
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromMinutes(3);
            try
            {
                driver.Navigate().GoToUrl(businessUrl);

                // Fetch social media links from the business website (e.g., from footer or social media icons)
                var socialMediaSelectors = new List<string>
                {
                    "a[href*='facebook.com']",
                    "a[href*='x.com']",
                    "a[href*='linkedin.com']",
                    "a[href*='instagram.com']",
                    "a[href*='youtube.com']",
                    "a[href*='pinterest.com']",
                };

                foreach (var selector in socialMediaSelectors)
                {
                    var elements = driver.FindElements(By.CssSelector(selector));
                    foreach (var element in elements)
                    {
                        var link = element.GetAttribute("href");
                        var links = driver.FindElements(By.TagName("a"));
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
                            if(!isInvalidValid)
                            {
                                if (!socialLinks.ContainsKey(baseDomain))
                                {
                                    socialLinks.Add(baseDomain, link);
                                }
                            }
                                
                        }
                    }
                }

                var emails = FetchEmails(driver);

                // Combine all emails into one string (comma-separated)
                if (emails.Count > 0)
                {
                    string combinedEmails = string.Join(", ", emails);
                    socialLinks["emails"] = combinedEmails; // Single key for all emails
                }
                else
                {
                    socialLinks["emails"] = "No emails found";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching social media links for {businessUrl}: {ex.Message}");
            }
            finally
            {
                driver.Quit();
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
            //foreach (var link in links)
            //{
            //    try
            //    {
            //        string href = link.GetAttribute("href");

            //        if (!string.IsNullOrEmpty(href) && href.StartsWith("mailto:"))
            //        {
            //            string email = href.Replace("mailto:", "").Trim();
            //            if (!emails.Contains(email) && Regex.IsMatch(email, emailPattern))
            //            {
            //                emails.Add(email);
            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine($"Error processing link: {ex.Message}");
            //    }
            //}

            return emails;
        }



        private bool ScrollToLoadMoreResults(IWebDriver driver, IWebElement mapContainer)
        {
            // Execute JavaScript to scroll the results container down (not just the window)
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;

            // Get the map container (or search results container)

            // Get the initial scroll height
            int lastHeight = Convert.ToInt32(jsExecutor.ExecuteScript("return arguments[0].scrollHeight", mapContainer));
            // Scroll down the results container
            jsExecutor.ExecuteScript("arguments[0].scrollTop = arguments[0].scrollHeight", mapContainer);
            Thread.Sleep(3000);
            int newHeight = Convert.ToInt32(jsExecutor.ExecuteScript("return arguments[0].scrollHeight", mapContainer));
            return newHeight > lastHeight;
            //bool reachedEnd = false;

            //while (!reachedEnd)
            //{

            //    // Wait for new results to load
            //    Thread.Sleep(3000); // Adjust this time as necessary

            //    // Get the new scroll height
            //    int newHeight = Convert.ToInt32(jsExecutor.ExecuteScript("return arguments[0].scrollHeight", mapContainer));
            //    Console.WriteLine($"New Height after Scroll: {newHeight}");

            //    // If the height hasn't changed, we've reached the end of the results
            //    if (newHeight == lastHeight)
            //    {
            //        reachedEnd = true; // No more results loaded
            //        Console.WriteLine("Reached the end of the results.");
            //    }
            //    else
            //    {
            //        // Otherwise, update the height and continue scrolling
            //        lastHeight = newHeight;
            //    }
            //}
        }

        private void ExportToExcel(System.Collections.Generic.List<(string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, Dictionary<string, string> socialMedias, string hrefValue)> results)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Results");
                    worksheet.Cell(1, 1).Value = "Keyword";
                    worksheet.Cell(1, 2).Value = "Result Title";
                    worksheet.Cell(1, 3).Value = "Category";
                    worksheet.Cell(1, 4).Value = "Location";
                    worksheet.Cell(1, 5).Value = "Contact Number";
                    worksheet.Cell(1, 6).Value = "Email";
                    worksheet.Cell(1, 7).Value = "Website";
                    worksheet.Cell(1, 8).Value = "Facebook";
                    worksheet.Cell(1, 9).Value = "LinkedIn";
                    worksheet.Cell(1, 10).Value = "Twitter";
                    worksheet.Cell(1, 11).Value = "Youtube";
                    worksheet.Cell(1, 12).Value = "Instagram";
                    worksheet.Cell(1, 13).Value = "Pinterest";
                    worksheet.Cell(1, 14).Value = "Rating";
                    worksheet.Cell(1, 15).Value = "Review Count";

                    for (int i = 0; i < results.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = results[i].SearchTerm;
                        worksheet.Cell(i + 2, 2).Value = results[i].ResultTitle;
                        worksheet.Cell(i + 2, 3).Value = results[i].Category;
                        worksheet.Cell(i + 2, 4).Value = results[i].Address;
                        worksheet.Cell(i + 2, 5).Value = results[i].ContactNumber;
                        worksheet.Cell(i + 2, 6).Value = results[i].socialMedias.ContainsKey("emails") ? results[i].socialMedias["emails"] : string.Empty;
                        worksheet.Cell(i + 2, 7).Value = results[i].hrefValue;
                        worksheet.Cell(i + 2, 8).Value = results[i].socialMedias.ContainsKey("facebook") ? results[i].socialMedias["facebook"] : string.Empty;
                        worksheet.Cell(i + 2, 9).Value = results[i].socialMedias.ContainsKey("linkedin") ? results[i].socialMedias["linkedin"] : string.Empty;
                        worksheet.Cell(i + 2, 10).Value = results[i].socialMedias.ContainsKey("x") ? results[i].socialMedias["x"] : string.Empty;
                        worksheet.Cell(i + 2, 11).Value = results[i].socialMedias.ContainsKey("youtube") ? results[i].socialMedias["youtube"] : string.Empty;
                        worksheet.Cell(i + 2, 12).Value = results[i].socialMedias.ContainsKey("instagram") ? results[i].socialMedias["instagram"] : string.Empty;
                        worksheet.Cell(i + 2, 13).Value = results[i].socialMedias.ContainsKey("pinterest") ? results[i].socialMedias["pinterest"] : string.Empty;
                        worksheet.Cell(i + 2, 14).Value = results[i].ReviewCount;
                        worksheet.Cell(i + 2, 15).Value = results[i].Rating;
                    }

                    string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "GoogleMapsResults.xlsx");
                    workbook.SaveAs(filePath);
                    MessageBox.Show($"Results exported to {filePath}", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

