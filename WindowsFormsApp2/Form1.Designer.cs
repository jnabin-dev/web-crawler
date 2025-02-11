﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    partial class Form1 : Form
    {
        private Button btnUpload;
        private Button btnStart;
        private Button stopCrawlButton;
        private Button clearButton;
        private Button exportButton;
        private GroupBox instructionsGroupBox;
        private GroupBox progressGroupBox;
        private TextBox instructionsTextBox;
        private TextBox progressTextBox;
        private DataGridView dataGridView;
        private ToolStrip toolStrip;
        private ToolStripDropDownButton helpDropDownButton;
        private ToolStripMenuItem helpMenuItem;
        private ToolStripMenuItem instructionsMenuItem;
        private ToolStripMenuItem addLicenseMenuItem;
        private ToolStripMenuItem contactBotsolMenuItem;
        private ToolStripMenuItem customizeDataFieldsMenuItem;
        private ToolStripMenuItem changeCrawlerMenuItem;
        private ToolStripMenuItem advancedOptionsMenuItem;
        private ToolStripMenuItem checkForUpdatesMenuItem;
        private ToolStripMenuItem aboutMenuItem;
        private PictureBox animatedLoader;
        //private ProgressBar progressBar;
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
            this.ClientSize = new Size(1100, 600);
            this.btnUpload = new Button();
            this.btnStart = new Button();
            this.stopCrawlButton = new Button();
            this.exportButton = new Button();
            this.clearButton = new Button();
            this.instructionsGroupBox = new GroupBox();
            this.progressGroupBox = new GroupBox();
            this.instructionsTextBox = new TextBox();
            this.progressTextBox = new TextBox();
            this.toolStrip = new ToolStrip();
            this.helpMenuItem = new ToolStripMenuItem();
            this.instructionsMenuItem = new ToolStripMenuItem();
            this.addLicenseMenuItem = new ToolStripMenuItem();
            this.contactBotsolMenuItem = new ToolStripMenuItem();
            this.customizeDataFieldsMenuItem = new ToolStripMenuItem();
            this.changeCrawlerMenuItem = new ToolStripMenuItem();
            this.advancedOptionsMenuItem = new ToolStripMenuItem();
            this.checkForUpdatesMenuItem = new ToolStripMenuItem();
            this.aboutMenuItem = new ToolStripMenuItem();
            this.helpDropDownButton = new ToolStripDropDownButton();
            //((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();

            // 
            // ToolStrip
            // 
            this.toolStrip.Dock = System.Windows.Forms.DockStyle.None; // Allow manual positioning
            this.toolStrip.Location = new System.Drawing.Point(this.ClientSize.Width - 112, 10); // Adjust based on form width
            this.toolStrip.Anchor = AnchorStyles.Top | AnchorStyles.Right; // Anchor to the top-right
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                this.helpDropDownButton
            });

            // 
            // Help DropDownButton
            // 
            this.helpDropDownButton.Text = "Help";
            this.helpDropDownButton.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.instructionsMenuItem,
            this.addLicenseMenuItem,
            this.contactBotsolMenuItem,
            this.customizeDataFieldsMenuItem,
            this.changeCrawlerMenuItem,
            this.advancedOptionsMenuItem,
            this.checkForUpdatesMenuItem,
            this.aboutMenuItem
            });

            this.helpDropDownButton.DropDownDirection = ToolStripDropDownDirection.Left; // Open the menu on the left side


            // 
            // Dropdown Items
            // 
            this.instructionsMenuItem.Text = "Instructions";
            this.instructionsMenuItem.Click += new System.EventHandler(this.InstructionsMenuItem_Click);

            this.addLicenseMenuItem.Text = "Add License";
            this.addLicenseMenuItem.Click += new System.EventHandler(this.AddLicenseMenuItem_Click);

            this.contactBotsolMenuItem.Text = "Contact Crawler";
            this.contactBotsolMenuItem.Click += new System.EventHandler(this.ContactBotsolMenuItem_Click);

            this.customizeDataFieldsMenuItem.Text = "Customize Data Fields";
            this.customizeDataFieldsMenuItem.Click += new System.EventHandler(this.CustomizeDataFieldsMenuItem_Click);

            this.changeCrawlerMenuItem.Text = "Change Crawler";
            this.changeCrawlerMenuItem.Click += new System.EventHandler(this.ChangeCrawlerMenuItem_Click);

            this.advancedOptionsMenuItem.Text = "Advanced Options";
            this.advancedOptionsMenuItem.Click += new System.EventHandler(this.AdvancedOptionsMenuItem_Click);

            this.checkForUpdatesMenuItem.Text = "Check For Updates";
            this.checkForUpdatesMenuItem.Click += new System.EventHandler(this.CheckForUpdatesMenuItem_Click);

            this.aboutMenuItem.Text = "About";
            this.aboutMenuItem.Click += new System.EventHandler(this.AboutMenuItem_Click);

            // 
            // instructionsGroupBox
            // 
            this.instructionsGroupBox.Text = "Instructions";
            this.instructionsGroupBox.Location = new System.Drawing.Point(10, 10); // Top-left corner
            this.instructionsGroupBox.Controls.Add(this.instructionsTextBox);

            // 
            // instructionsTextBox
            // 
            this.instructionsTextBox.Multiline = true;
            this.instructionsTextBox.ReadOnly = true;
            this.instructionsTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.instructionsTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.instructionsTextBox.Text = "The bot will open a Chrome window...\r\n" + // Add your instructions here
                                        "* FREE version only extracts 20 results...\r\n" +
                                        "* Make sure language is set to English...";

            // 
            // progressGroupBox
            // 
            this.progressGroupBox.Text = "Progress";
            this.progressGroupBox.Controls.Add(this.progressTextBox);
            string spinnerPath = Path.Combine(Application.StartupPath, "assets", "img", "spinner.gif");
            animatedLoader = new PictureBox
            {
                Image = Image.FromFile(spinnerPath), // Path to your animated GIF
                Location = new System.Drawing.Point(55, progressTextBox.Top-1), // Position beside the label
                Size = new System.Drawing.Size(15, 15), // Adjust size as needed
                SizeMode = PictureBoxSizeMode.StretchImage,
                Visible = false // Initially hidden
            };
            progressGroupBox.Controls.Add(animatedLoader);

            // ProgressBar
            //progressBar = new ProgressBar
            //{
            //    Location = new System.Drawing.Point(progressTextBox.Right + 10, progressTextBox.Top), // Position beside the label
            //    Size = new System.Drawing.Size(200, 20),
            //    Style = ProgressBarStyle.Marquee, // Indeterminate progress style
            //    Visible = false // Initially hidden
            //};
            //progressGroupBox.Controls.Add(progressBar);

            // 
            // progressTextBox
            // 
            this.progressTextBox.Multiline = true;
            this.progressTextBox.ReadOnly = true;
            this.progressTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.progressTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressTextBox.ForeColor = Color.DarkBlue; // Set text color to blue
            this.progressTextBox.BackColor = Color.White; // Optional: Ensure background is white for contrast

            this.btnUpload.Location = new System.Drawing.Point(10, 180); // Set position
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(100, 40); // Set size
            this.btnUpload.TabIndex = 0;
            this.btnUpload.Text = "Upload File";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);

            this.btnStart.Location = new System.Drawing.Point(130, 180); // Set position
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(100, 40); // Set size
            this.btnStart.TabIndex = 1;
            this.btnStart.Text = "Start Crawling";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);

            this.stopCrawlButton.Location = new System.Drawing.Point(250, 180); // Adjust position
            this.stopCrawlButton.Name = "stopCrawlButton";
            this.stopCrawlButton.Size = new System.Drawing.Size(100, 40); // Adjust size
            this.stopCrawlButton.Text = "Stop Crawl";
            this.stopCrawlButton.UseVisualStyleBackColor = true;
            this.stopCrawlButton.Click += new System.EventHandler(this.stopCrawlButton_Click);
            
            this.exportButton.Location = new System.Drawing.Point(370, 180); // Adjust position
            this.exportButton.Name = "exportDataButton";
            this.exportButton.Size = new System.Drawing.Size(100, 40); // Adjust size
            this.exportButton.Text = "Export";
            this.exportButton.UseVisualStyleBackColor = true;
            this.exportButton.Click += new System.EventHandler(this.exportDataButton_Click);
            
            this.clearButton.Location = new System.Drawing.Point(490, 180); // Adjust position
            this.clearButton.Name = "clearDataButton";
            this.clearButton.Size = new System.Drawing.Size(100, 40); // Adjust size
            this.clearButton.Text = "Clear";
            this.clearButton.UseVisualStyleBackColor = true;
            this.clearButton.Click += new System.EventHandler(this.clearDataButton_Click);

            //Panel panel = new Panel
            //{
            //    Dock = DockStyle.Fill // Fill remaining space after the buttons
            //};
            this.dataGridView = new System.Windows.Forms.DataGridView();
            //this.dataGridView.Dock = DockStyle.Fill;
            this.dataGridView.Location = new System.Drawing.Point(10, 260);
            AdjustDataGridViewHeight();
            this.dataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            this.dataGridView.AutoResizeColumns();
            //this.dataGridView.ColumnCount = SharedDataTableModel.SelectedFields.Count;
            SharedDataTableModel.SelectedFields = SharedDataTableModel.DataGridColumns;
            this.dataGridView.ReadOnly = true;
            this.dataGridView.AllowUserToAddRows = false;
            UpdateDataTableColumns();
            //panel.Controls.Add(dataGridView);
            // 
            // Form1
            Label lblFooter = new Label();
            lblFooter.Text = "A Product of Codezzi | All Rights Reserved";
            lblFooter.Font = new Font("Arial", 10, FontStyle.Italic);
            lblFooter.ForeColor = Color.Gray;
            lblFooter.AutoSize = false;
            lblFooter.TextAlign = ContentAlignment.MiddleRight;
            lblFooter.Dock = DockStyle.Bottom;
            lblFooter.Height = 25; // Adjust as needed
            lblFooter.Padding = new Padding(0, 2, 10, 0); // Right padding for alignment
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            //this.ClientSize = new System.Drawing.Size(400, 200); // Adjust size as needed
            //this.Controls.Add(this.dataGridView);
            //this.Controls.Add(panel);
            this.Controls.Add(this.toolStrip); // Add the MenuStrip to the form
            this.Controls.Add(this.instructionsGroupBox);
            this.Controls.Add(this.progressGroupBox);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(lblFooter);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.stopCrawlButton);
            this.Controls.Add(this.exportButton);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.dataGridView);
            this.Name = "MainForm";
            this.Text = "Sky Crawler";
            this.FormClosing += Form1_FormClosing;
            //((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private void AdjustDataGridViewHeight()
        {
            int formWidth = this.ClientSize.Width;
            int groupBoxWidth = (formWidth - 30) / 2; // 30 accounts for margins (10px left, 10px middle gap, 10px right)

            // Instructions GroupBox
            this.instructionsGroupBox.Size = new System.Drawing.Size(groupBoxWidth, 150); // 150px height
            this.instructionsGroupBox.Location = new System.Drawing.Point(10, 25);

            // Progress GroupBox
            this.progressGroupBox.Size = new System.Drawing.Size(groupBoxWidth, 150); // Match height
            this.progressGroupBox.Location = new System.Drawing.Point(20 + groupBoxWidth, 25); // Offset by left margin + width of first GroupBox

            int topButtonsBottom = btnUpload.Bottom; // Get the bottom position of the topmost button
            int bottomPadding = 30; // Add some padding at the bottom
            int availableHeight = this.ClientSize.Height - topButtonsBottom - bottomPadding - 10;
            dataGridView.Width = groupBoxWidth*2;
            // Set DataGridView's top position and height
            dataGridView.Top = topButtonsBottom + 15; // Add some padding below the buttons
            dataGridView.Height = availableHeight;
        }

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
                        string selectedFileName = openFileDialog.FileName;
                        string fileNameOnly = System.IO.Path.GetFileName(selectedFileName); // Get only the file name
                        UpdateProgress($"File Selected: {fileNameOnly}");
                        UpdateProgress($"Loaded {searchTerms.Length} search terms.");
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
            manageStartButton(false);
            //progressBar.Visible = true;
            //InitializeDataTable();
            if (cancellationTokenSource != null)
            {
                cancellationTokenSource.Dispose();
            }
            cancellationTokenSource = new CancellationTokenSource();
            dataGridView.Rows.Clear();
            //UpdateProgress("", true);
            results = new System.Collections.Generic.List<(string SearchTerm, string ResultTitle, string ReviewCount, string Rating, string ContactNumber, string Category, string Address, string StreetAddress, string city, string zip, string country, Dictionary<string, string>, string companyWebsite)>();
            var confirmationForm = new ConfirmationForm("Do you want to scrape email and other social media links from the business website? (It can slow down the crawler speed.)");
            var result = confirmationForm.ShowDialog();
            animatedLoader.Visible = true;
            UpdateProgress("Crawling started...");
            if (result == DialogResult.OK && confirmationForm.UserConfirmed)
            {
                // User clicked "Yes"
                new Thread(async () => await StartCrawlingAsync(cancellationTokenSource.Token, true)).Start();
                // Start the crawler with additional features
            }
            else
            {
                // User clicked "No" or closed the dialog
                new Thread(async () => await StartCrawlingAsync(cancellationTokenSource.Token, false)).Start();
                // Start the crawler without additional features
            }
        }

        // Menu Item Event Handlers
        private void InstructionsMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("The bot will open a Chrome window to search for businesses...", "Instructions");
        }

        private void AddLicenseMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Add License functionality is not yet implemented.", "Add License");
        }

        private void ContactBotsolMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("You can contact Botsol via support@example.com.", "Contact Botsol");
        }

        private void CustomizeDataFieldsMenuItem_Click(object sender, EventArgs e)
        {
            // Open the new form with checkboxes
            CustomizeDataFieldsForm customizeForm = new CustomizeDataFieldsForm();
            customizeForm.ShowDialog(); // Open as a modal dialog

            UpdateDataTableColumns();
        }

        private void UpdateDataTableColumns()
        {
            // Clear existing columns
            dataGridView.Columns.Clear();

            for (int index = 0; index < SharedDataTableModel.SelectedFields.Count; index++)
            {
                this.dataGridView.Columns.Add(SharedDataTableModel.SelectedFields[index].Name, SharedDataTableModel.SelectedFields[index].Name);
                if (SharedDataTableModel.SelectedFields[index].Width != null)
                {
                    this.dataGridView.Columns[index].Width = Convert.ToInt32(SharedDataTableModel.SelectedFields[index].Width);
                }
            }
        }

        private void ChangeCrawlerMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Change Crawler functionality is not yet implemented.", "Change Crawler");
        }

        private void AdvancedOptionsMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Advanced Options functionality is not yet implemented.", "Advanced Options");
        }

        private void CheckForUpdatesMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Check For Updates functionality is not yet implemented.", "Check For Updates");
        }

        private void AboutMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Codezzi Crawler Version 1.0", "About");
        }
    }
}

