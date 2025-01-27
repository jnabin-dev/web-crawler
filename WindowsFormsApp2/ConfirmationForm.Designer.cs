using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Windows.Forms;
using System;

namespace WindowsFormsApp2
{
    partial class ConfirmationForm
    {
        public bool UserConfirmed { get; private set; } = false; // Track if "Yes" was clicked

        private Button yesButton;
        private Button noButton;
        private Label messageLabel;
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
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Text = "ConfirmationForm";

            // Label
            messageLabel = new Label
            {
                Text = "Do you want to scrape email and other social media links from the business website? (It can slow down the crawler speed.)",
                AutoSize = true,
                Location = new System.Drawing.Point(10, 10),
                MaximumSize = new System.Drawing.Size(300, 0), // Word wrap within 300px
            };

            // Yes Button
            yesButton = new Button
            {
                Text = "Yes",
                Location = new System.Drawing.Point(10, messageLabel.Bottom + 40),
                Size = new System.Drawing.Size(75, 30)
            };
            yesButton.Click += YesButton_Click;

            // No Button
            noButton = new Button
            {
                Text = "No",
                Location = new System.Drawing.Point(110, messageLabel.Bottom + 40),
                Size = new System.Drawing.Size(75, 30)
            };
            noButton.Click += NoButton_Click;

            // Form properties
            this.Text = "Crawling Preference Required";
            this.ClientSize = new System.Drawing.Size(350, yesButton.Bottom + 20);
            this.Controls.Add(messageLabel);
            this.Controls.Add(yesButton);
            this.Controls.Add(noButton);
            this.StartPosition = FormStartPosition.CenterParent; // Center relative to the parent form
        }

        #endregion

        private void YesButton_Click(object sender, EventArgs e)
        {
            UserConfirmed = true; // Indicate "Yes" was clicked
            this.DialogResult = DialogResult.OK; // Close the form with an OK result
            this.Close();
        }

        private void NoButton_Click(object sender, EventArgs e)
        {
            UserConfirmed = false; // Indicate "No" was clicked
            this.DialogResult = DialogResult.Cancel; // Close the form with a Cancel result
            this.Close();
        }
    }
}