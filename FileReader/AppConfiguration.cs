using System.IO;
using Microsoft.Extensions.Configuration;

namespace FileReader
{
    public static class AppConfiguration
    {
        public static IConfigurationRoot ReadConfigurationFromAppSettings()
        {
            var configurationBuilder = new ConfigurationBuilder();
            var path = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
            configurationBuilder.AddJsonFile(path, false);
            var root = configurationBuilder.Build();
            return root;
        }
    }
}