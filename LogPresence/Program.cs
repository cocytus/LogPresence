using System;
using System.Globalization;
using System.ServiceProcess;

namespace LogPresence
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
            if (Environment.UserInteractive)
            {
                PresenceSaver.PostProcessIfNewDay();
            }
            else
            {
                var servicesToRun = new ServiceBase[]
                {
                    new PresenceSaver()
                };
                ServiceBase.Run(servicesToRun);
            }
        }
    }
}
