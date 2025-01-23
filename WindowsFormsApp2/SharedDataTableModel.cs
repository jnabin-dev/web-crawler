using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp2
{
    public class SharedDataTableModel
    {
        public static List<GridColumnConfig> SelectedFields { get; set; }
        public static List<GridColumnConfig> DataGridColumns { get; set; } = new List<GridColumnConfig> {
                new GridColumnConfig("#", 50),
                new GridColumnConfig("Keyword", null),
                new GridColumnConfig("Name", null),
                new GridColumnConfig("Category", null),
                new GridColumnConfig("Full_Address", 150),
                new GridColumnConfig("Street_Address", null),
                new GridColumnConfig("City", 70),
                new GridColumnConfig("Zip", 70),
                new GridColumnConfig("Country", null),
                new GridColumnConfig("Contact Number", null),
                new GridColumnConfig("Email", null),
                new GridColumnConfig("Website", null),
                new GridColumnConfig("Facebook", null),
                new GridColumnConfig("Linkedin", null),
                new GridColumnConfig("Twitter", null),
                new GridColumnConfig("Youtube", null),
                new GridColumnConfig("Instagram", null),
                new GridColumnConfig("Pinterest", null),
                new GridColumnConfig("Rating", null),
                new GridColumnConfig("Review Count", null),
            };

        public SharedDataTableModel()
        {
            SelectedFields = DataGridColumns;
        }
    }
}
