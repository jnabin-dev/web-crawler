using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Reflection;

namespace WindowsFormsApp2
{
    public class FontLoader
    {
        private static PrivateFontCollection privateFonts = new PrivateFontCollection(); // Keep it in memory

        public static Font LoadFont(FontStyle style, float size = 10)
        {
            if (privateFonts.Families.Length > 0) // Check if already loaded
            {
                return new Font(privateFonts.Families[0], size, style);
            }

            // ✅ Corrected Embedded Resource Path (Confirm it using GetManifestResourceNames())
            string resourceName = "WindowsFormsApp2.assets.fonts.DancingScript-Regular.ttf";

            using (Stream fontStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (fontStream == null)
                {
                    MessageBox.Show("❌ Font file not found. Check your embedded resource path.");
                    return new Font("Arial", size, style); // Fallback font
                }

                byte[] fontData = new byte[fontStream.Length];
                fontStream.Read(fontData, 0, fontData.Length);
                IntPtr fontPtr = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(fontData.Length);
                System.Runtime.InteropServices.Marshal.Copy(fontData, 0, fontPtr, fontData.Length);
                privateFonts.AddMemoryFont(fontPtr, fontData.Length);
                System.Runtime.InteropServices.Marshal.FreeCoTaskMem(fontPtr);
            }

            return new Font(privateFonts.Families[0], size, style);
        }

        // ✅ Apply font to all controls recursively
        public static void ApplyFontToAllControls(Control parentControl, Font font)
        {
            //foreach (Control ctrl in parentControl.Controls)
            //{
            //    ctrl.Font = font; // Apply font
            //    if (ctrl.HasChildren)
            //    {
            //        ApplyFontToAllControls(ctrl, font); // Recursive call for nested controls
            //    }
            //}
        }
    }
}
