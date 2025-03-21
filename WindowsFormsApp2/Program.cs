using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                Application.Run(new Form1());
            }
            catch(Exception exc)
            {
                LoggerService.Error("root excaption", exc);
            }
            finally
            {
                //PowerHelper.RestoreSleep(); // ✅ Restore sleep even if app crashes
            }
        }
    }
}
