using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Keys = OpenQA.Selenium.Keys;

namespace WindowsFormsApp2
{
    partial class Form1 : Form
    {
        private string[] searchTerms;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.Button btnStart;
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
            this.SuspendLayout();
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
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.btnStart);
            this.Name = "MainForm";
            this.Text = "Google Maps Crawler";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

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
            new Thread(() => StartCrawling()).Start();
        }


        private void StartCrawling()
        {
            try
            {
                var results = new System.Collections.Generic.List<(string SearchTerm, string ResultTitle, string Address)>();
                ChromeOptions options = new ChromeOptions();
                options.AddArgument("C:\\Users\\Skytech\\Downloads\\chrome-win64\\chrome-win64");
                IWebDriver driver = new ChromeDriver(options);

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

                    ScrollToLoadMoreResults(driver);

                    // Extract search results
                    var resultElements = driver.FindElements(By.XPath("//div[@class='bfdHYd Ppzolf OFBs3e  ']"));
                    foreach (var resultElement in resultElements)
                    {
                        try
                        {
                            string title = resultElement.FindElement(By.CssSelector(".qBF1Pd.fontHeadlineSmall")).Text;
                            string address = resultElement.FindElement(By.ClassName("W4Efsd")).Text;
                            results.Add((term, title, address));
                        }
                        catch(Exception ex) 
                        {
                            // Skip if any element is missing
                        }
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

        private void ScrollToLoadMoreResults(IWebDriver driver)
        {
            // Execute JavaScript to scroll the results container down (not just the window)
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;

            // Get the map container (or search results container)
            var mapContainer = driver.FindElement(By.XPath("//div[@class='m6QErb DxyBCb kA9KIf dS8AEf XiKgde ecceSd']"));

            // Get the initial scroll height
            int lastHeight = Convert.ToInt32(jsExecutor.ExecuteScript("return arguments[0].scrollHeight", mapContainer));
            Console.WriteLine($"Initial Height: {lastHeight}");

            bool reachedEnd = false;

            while (!reachedEnd)
            {
                // Scroll down the results container
                jsExecutor.ExecuteScript("arguments[0].scrollTop = arguments[0].scrollHeight", mapContainer);

                // Wait for new results to load
                Thread.Sleep(3000); // Adjust this time as necessary

                // Get the new scroll height
                int newHeight = Convert.ToInt32(jsExecutor.ExecuteScript("return arguments[0].scrollHeight", mapContainer));
                Console.WriteLine($"New Height after Scroll: {newHeight}");

                // If the height hasn't changed, we've reached the end of the results
                if (newHeight == lastHeight)
                {
                    reachedEnd = true; // No more results loaded
                    Console.WriteLine("Reached the end of the results.");
                }
                else
                {
                    // Otherwise, update the height and continue scrolling
                    lastHeight = newHeight;
                }
            }
        }

        private void ExportToExcel(System.Collections.Generic.List<(string SearchTerm, string ResultTitle, string Address)> results)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Results");
                    worksheet.Cell(1, 1).Value = "Search Term";
                    worksheet.Cell(1, 2).Value = "Result Title";
                    worksheet.Cell(1, 3).Value = "Address";

                    for (int i = 0; i < results.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = results[i].SearchTerm;
                        worksheet.Cell(i + 2, 2).Value = results[i].ResultTitle;
                        worksheet.Cell(i + 2, 3).Value = results[i].Address;
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

