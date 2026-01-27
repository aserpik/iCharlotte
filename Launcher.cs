using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

class Program
{
    [DllImport("shell32.dll", SetLastError = true)]
    static extern void SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);

    static void Main(string[] args)
    {
        // Set AppUserModelID for taskbar grouping
        SetCurrentProcessExplicitAppUserModelID("iCharlotte.LegalSuite.1");

        // Get the directory where this exe is located
        string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string exeDir = Path.GetDirectoryName(exePath);
        string scriptPath = Path.Combine(exeDir, "iCharlotte.py");

        // Launch pythonw with the script
        ProcessStartInfo psi = new ProcessStartInfo();
        psi.FileName = "pythonw.exe";
        psi.Arguments = "\"" + scriptPath + "\"";
        psi.WorkingDirectory = exeDir;
        psi.UseShellExecute = false;

        // Pass through any command line arguments
        foreach (string arg in args)
        {
            psi.Arguments += " \"" + arg + "\"";
        }

        Process.Start(psi);
    }
}
