using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Extensions.Configuration;
using SimpleImpersonation;

namespace FileReader.FileWriter
{
    public static class LocalStorageService
    {
        private static readonly IConfiguration Configuration = AppConfiguration.ReadConfigurationFromAppSettings();



        public static void StoreBytes(byte[] bytes, string path, string name, string extension)
        {
            try
            {
                var username = Configuration["Network:Username"];
                var password = Configuration["Network:Password"];
                var domain = Configuration["Network:Domain"];
                var location = Configuration["StoragePath"];
                Console.WriteLine("Username At Download:  " + username);
                Console.WriteLine("Password At Download:  " + password);
                Console.WriteLine("Domain At Download:  " + domain);
                Console.WriteLine("Location At Download:  " + location);
                var credentials = new UserCredentials(domain, username, password);
                Impersonation.RunAsUser(credentials, LogonType.Interactive, () =>
               {
                   var uploadsFolderPath = Path.Combine($"{path}");
                   if (!Directory.Exists(uploadsFolderPath))
                       Directory.CreateDirectory(uploadsFolderPath);
                   var newFileName = name + "."+ extension;
                   var filePath = Path.Combine(uploadsFolderPath, newFileName);
                   using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
                   fs.Write(bytes, 0, bytes.Length);
               });

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

        }

        public static IEnumerable<object> GetDirectoriesAsync()
        {
            try
            {
                var username = Configuration["Network:Username"];
                var password = Configuration["Network:Password"];
                var domain = Configuration["Network:Domain"];
                var location = Configuration["StoragePaths:Location"];
                Console.WriteLine("Username At Download:  " + username);
                Console.WriteLine("Password At Download:  " + password);
                Console.WriteLine("Domain At Download:  " + domain);
                Console.WriteLine("Location At Download:  " + location);
                var credentials = new UserCredentials(domain, username, password);
                var result = Impersonation.RunAsUser(credentials, LogonType.Interactive, () => Directory.GetFiles(@location));
                foreach (var item in result)
                {
                    Console.WriteLine(item);
                }
               
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}
