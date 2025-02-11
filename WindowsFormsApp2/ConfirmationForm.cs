using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class ConfirmationForm : Form
    {
        private string ConfirmationText;
        public ConfirmationForm(string confirmationText)
        {
            ConfirmationText = confirmationText;
            InitializeComponent();
            string fontPath = Path.Combine(Application.StartupPath, "assets", "fonts", "Mulish-Regular.ttf");
            Font mulishRegularFont = FontLoader.LoadCustomFont(fontPath, FontLoader.fontSize, FontStyle.Regular);

            FontLoader.ApplyFontToAllControls(this, mulishRegularFont);
        }
    }
}
