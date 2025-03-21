using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp2
{
    public class PowerHelper
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetThreadExecutionState(ExecutionState esFlags);

        [Flags]
        public enum ExecutionState : uint
        {
            ES_CONTINUOUS = 0x80000000,
            ES_DISPLAY_REQUIRED = 0x00000002,
            ES_SYSTEM_REQUIRED = 0x00000001
        }

        // ✅ Call this when the bot starts to **prevent sleep mode**
        public static void PreventSleep()
        {
            SetThreadExecutionState(ExecutionState.ES_CONTINUOUS | ExecutionState.ES_SYSTEM_REQUIRED);
        }

        // ✅ Call this when the bot closes to **restore normal sleep behavior**
        public static void RestoreSleep()
        {
            SetThreadExecutionState(ExecutionState.ES_CONTINUOUS); // ✅ Allows Windows to sleep again
        }
    }
}
