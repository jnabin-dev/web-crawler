using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    partial class CustomizeDataFieldsForm
    {
        private List<CheckBox> checkBoxes = new List<CheckBox>();
        private Button saveButton;

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
            this.Text = "CustomizeDataFieldsForm";

            int columnCount = 0;
            int checkBoxPerColumn = 10; // Maximum checkboxes per column
            int columnWidth = 120; // Width of each column
            int startX = 10; // Starting X position for the first column
            int startY = 10; // Starting Y position for the first checkbox
            int yPosition = startY;
            // Dynamically create checkboxes for each available field
            foreach (var field in SharedDataTableModel.DataGridColumns)
            {
                // Create a checkbox
                CheckBox checkBox = new CheckBox
                {
                    Text = field.Name,
                    Checked = field.IsSelected, // Initially selected
                    Location = new System.Drawing.Point(startX + (columnWidth * columnCount), yPosition),
                    AutoSize = true
                };

                this.Controls.Add(checkBox);
                checkBoxes.Add(checkBox);

                // Update position for the next checkbox
                yPosition += 30; // Space between checkboxes
                // Move to a new column if we've added 10 checkboxes in the current column
                if (checkBoxes.Count % checkBoxPerColumn == 0)
                {
                    columnCount++;
                    yPosition = startY; // Reset Y position for the new column
                }
            }

            // Save Button
            saveButton = new Button
            {
                Text = "Save",
                Location = new System.Drawing.Point(10, checkBoxPerColumn * 30 + startY) // Dynamically position below the last checkbox
            };
            this.Controls.Add(saveButton);
            saveButton.Click += SaveButton_Click;

            // Form properties
            int formHeight = Math.Max(yPosition + 80, checkBoxPerColumn * 30 + startY + 30); // Ensure enough height for rows
            int formWidth = startX + (columnWidth * (columnCount + 1)) + 10; // Ensure enough width for columns
            this.Text = "Customize Data Fields";
            this.ClientSize = new System.Drawing.Size(formWidth, formHeight);
        }

        #endregion

        private void SaveButton_Click(object sender, EventArgs e)
        {
            foreach (var item in SharedDataTableModel.DataGridColumns)
            {
                item.IsSelected = checkBoxes.Where(x => x.Text == item.Name).FirstOrDefault().Checked;
            }

            SharedDataTableModel.SelectedFields = checkBoxes
            .Where(cb => cb.Checked)
            .Select(cb => new GridColumnConfig(cb.Text, null))
            .ToList();

            // Close the form
            this.Close();

            //MessageBox.Show(selectedFields, "Save Confirmation");

            // Close the form after saving
            //this.Close();
        }
    }
}