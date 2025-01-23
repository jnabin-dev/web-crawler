using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp2
{
    public class GridColumnConfig
    {
        public GridColumnConfig(string name, int? widht, bool selected = true)
        {
            Name = name;
            IsSelected = selected;
            if (widht != null)
            {
                Width = widht;
            }
            else
            {
                switch (Name)
                {
                    case "#":
                        Width = 50;
                        break;
                    case "Full_Address":
                        Width = 150;
                        break;
                    case "City":
                        Width = 70;
                        break;
                    case "Zip":
                        Width = 70;
                        break;
                    default:
                        Width = null;
                        break;

                }
            }

        }
        public string Name { get; set; }
        public int? Width { get; set; }
        public bool IsSelected { get; set; } = true;
    }
}
