using System;
using System.Security.Principal;
using Microsoft.Deployment.WindowsInstaller;

namespace DTFTest
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!IsAdmin())
            {
                Console.Write("Please re-launch Visual Studio with administrator rights (right click icon = run as admin).");
                Console.ReadLine();
                System.Environment.Exit(-1);
            }

            try
            {
                Installer.SetInternalUI(InstallUIOptions.Silent);
                Installer.EnableLog(InstallLogModes.Verbose, @"E:\WixTest4.log");
                Installer.InstallProduct(@"E:\WixTest.msi", "");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception:\n\n " + ex.Message);
            }
        }

        static bool IsAdmin ()
        {
            using (var identity = WindowsIdentity.GetCurrent())
            {
                var principal = new WindowsPrincipal(identity);
                if (!principal.IsInRole(WindowsBuiltInRole.Administrator))
                {
                    return false;
                }
            }

            return true;
        }
    }
}

// System.Environment.Exit(-1);
// Installer.ConfigureProduct(, 0, InstallState.Absent, "REBOOT=\"R\"");
// Installer.EnableLog(InstallLogModes.Error, errorLogPath);
