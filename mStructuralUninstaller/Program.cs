using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System;
using System.Windows.Forms;
using System.Reflection;

namespace mStructuralUninstaller
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Run Bat file
            ProcessStartInfo processInfo;
            Process process;
            bool rResult = false;

            string guid = "7dabe4bf-0e94-411a-88e6-a5464fa4b01e";

            string installloc = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            processInfo = new ProcessStartInfo(installloc + @"\Uninstall.bat");
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = false;
            processInfo.RedirectStandardInput = true;
            processInfo.Verb = "runas";
            try
            {
                process = Process.Start(processInfo);
                int count = 0;

                do
                {
                    if (!process.HasExited)
                    {
                        if (count > 10)
                        {
                            MessageBox.Show("Uninstallation failed, please contact admin");
                            return;
                        }
                        count++;
                    }
                }
                while (!process.WaitForExit(500));

                rResult = true;

                // if success, delete registry
                if (rResult)
                {
                    RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", true);
                    key.DeleteSubKeyTree("{" + guid.ToUpper() + "}");
                    key.Close();
                }
                else
                {
                    MessageBox.Show("Failed to Uninstall");
                    return;
                }

                ProcessStartInfo psi = new ProcessStartInfo("cmd.exe",
                    String.Format("/k {0} & {1} & {2}",
                        "timeout /T 1 /NOBREAK >NUL",
                        "rmdir /s /q \"" + Application.StartupPath + "\"",
                        "exit"
                    )
                );

                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.WindowStyle = ProcessWindowStyle.Hidden;

                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }
    }
}
