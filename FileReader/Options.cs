using CommandLine;
using Microsoft.Extensions.Configuration;

namespace FileReader
{
    public class Options
    {
        private readonly IConfiguration _configuration = AppConfiguration.ReadConfigurationFromAppSettings();

        public Options(string tableName, string dbFileName)
        {
            Server = _configuration["SqlCredentials:Server"];
            Database = _configuration["SqlCredentials:Database"];
            UserId = _configuration["SqlCredentials:Username"];
            Password = _configuration["SqlCredentials:Password"];
            Dbf = _configuration["DBFiles:" + dbFileName];
            Table = _configuration["Tables:" + tableName];
        }

        [Option(Required = true, HelpText = "The database server")]
        public string Server { get; set; }

        [Option(Required = true, HelpText = "The name of the database")]
        public string Database { get; set; }

        [Option(Required = true, HelpText = "The UserID used to connect to the database server")]
        public string UserId { get; set; }

        [Option(Required = true, HelpText = "The password used to connect to the database server")]
        public string Password { get; set; }

        [Option(Required = true, HelpText = "Path to the DBF file to import")]
        public string Dbf { get; set; }

        [Option(Required = true, HelpText = "The name of the database table to import into")]
        public string Table { get; set; }

        [Option(Default = 3000, HelpText = "The connection timeout used in the bulk copy operation")]
        public int BulkCopyTimeout { get; set; } = 50000000;

        [Option(Default = true, HelpText = "Whether to truncate the table before copying")]
        public bool Truncate { get; set; } = true;

        [Option(Default = true, HelpText = "Whether to skip deleted records")]
        public bool SkipDeletedRecords { get; set; } = true;
    }
}