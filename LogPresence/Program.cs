using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using System;
using System.Globalization;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;

namespace LogPresence
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public static void Main()
        {
            CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;

            var cb = new ConfigurationBuilder();
            cb.AddJsonFile("appsettings.json");
            var config = cb.Build();

            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            try
            {
                if (!Console.IsInputRedirected)
                {
                    new PresenceSaver(config).PostProcessIfNewDay();
                }
                else
                {
                    Console.WriteLine("Redirected -  run SVC");
                    ServiceBase.Run(new PresenceSaver(config));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.WriteLine("Press any key to continue!");
                Console.ReadKey();
            }
        }
    }
}
