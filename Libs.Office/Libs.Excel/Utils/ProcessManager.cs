using System.Diagnostics;

namespace Libs.Excel.Utils
{
    public class ProcessManager
    {
        public static void Kill(int processId)
        {
            Process process = Process.GetProcessById(processId);
            if (process != null)
            {
                process.Kill();
                process.WaitForExit();
            }
        }
    }
}