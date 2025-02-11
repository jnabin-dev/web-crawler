using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Text;
using System.Runtime.InteropServices;

namespace WindowsFormsApp2
{
    public class FontLoader
    {
        private static PrivateFontCollection privateFontCollection = new PrivateFontCollection();
        public static float fontSize = 8.5f;

        public static Font LoadCustomFont(string fontFilePath, float fontSize, FontStyle fontStyle = FontStyle.Regular)
        {
            if (!File.Exists(fontFilePath))
            {
                throw new FileNotFoundException("Font file not found!", fontFilePath);
            }

            byte[] fontData = File.ReadAllBytes(fontFilePath);
            IntPtr fontPtr = Marshal.AllocCoTaskMem(fontData.Length);
            Marshal.Copy(fontData, 0, fontPtr, fontData.Length);

            privateFontCollection.AddMemoryFont(fontPtr, fontData.Length);
            Marshal.FreeCoTaskMem(fontPtr);

            return new Font(privateFontCollection.Families[0], fontSize, fontStyle);
        }

        // ✅ Apply font to all controls recursively
        public static void ApplyFontToAllControls(Control parentControl, Font font)
        {
            foreach (Control ctrl in parentControl.Controls)
            {
                ctrl.Font = font; // Apply font
                if (ctrl.HasChildren)
                {
                    ApplyFontToAllControls(ctrl, font); // Recursive call for nested controls
                }
            }
        }
    }
}
