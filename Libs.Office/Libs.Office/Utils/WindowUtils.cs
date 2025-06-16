using System.Runtime.InteropServices;
using System;

namespace Libs.Office.Utils
{
    public class WindowUtils
    {
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);
    }
}
