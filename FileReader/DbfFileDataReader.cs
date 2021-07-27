using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Text;
using CommandLine.Text;
using DbfDataReader;

namespace FileReader
{
    public static class DbfFileDataReader
    {
         /// <summary>
        /// Run And Read Options
        /// </summary>
        /// <param name="options"></param>
        /// <returns></returns>
        public static int RunAndReturnExitCode(Options options)
        {
            Console.WriteLine(HeadingInfo.Default);
            Console.WriteLine();
            Console.WriteLine("Bulk copy from:");
            Console.WriteLine($"  DBF: {options.Dbf}");
            Console.WriteLine("to:");
            Console.WriteLine($"  Server: {options.Server}");
            Console.WriteLine($"  Database: {options.Database}");
            Console.WriteLine($"  Table: {options.Table}");
            Console.WriteLine($"  BulkCopyTimeout: {options.BulkCopyTimeout}");
            Console.WriteLine($"  UserID: {options.UserId}");
            Console.WriteLine($"  Truncate: {options.Truncate}");
            Console.WriteLine($"  SkipDeletedRecords: {options.SkipDeletedRecords}");
            Console.WriteLine();

            var connectionString = BuildConnectionString(options);

            using var connection = new SqlConnection(connectionString);
            connection.Open();

            if (options.Truncate)
            {
                TruncateTable(connection, options.Table);
            }

            DoBulkCopy(connection, options);

            return 0;
        }

        /// <summary>
        /// Truncate Table
        /// </summary>
        /// <param name="connection"></param>
        /// <param name="table"></param>
        private static void TruncateTable(SqlConnection connection, string table)
        {
            Console.WriteLine($"Truncating table '{table}'");

            var sql = $"truncate table {table};";

            var stopwatch = new Stopwatch();
            stopwatch.Start();

            var command = new SqlCommand(sql, connection);
            command.ExecuteNonQuery();

            stopwatch.Stop();

            Console.WriteLine($"Truncating table '{table}' completed in {GetElapsedTime(stopwatch)}s");
        }

        /// <summary>
        /// Bulk Copy To SQL
        /// </summary>
        /// <param name="connection"></param>
        /// <param name="options"></param>
        private static void DoBulkCopy(SqlConnection connection, Options options)
        {
            Console.WriteLine("Begin bulk copy");

            var stopwatch = new Stopwatch();
            stopwatch.Start();

            var rowsCopied = 0L;
            var dbfRecordCount = 0L;
            var dbfDataReaderOptions = new DbfDataReaderOptions
            {
                SkipDeletedRecords = options.SkipDeletedRecords
            };
            using var dbfDataReader = new DbfDataReader.DbfDataReader(options.Dbf, dbfDataReaderOptions);
            dbfRecordCount = dbfDataReader.DbfTable.Header.RecordCount;

            using var bulkCopy = new SqlBulkCopy(connection)
            {
                BulkCopyTimeout = options.BulkCopyTimeout, DestinationTableName = options.Table
            };

            try
            {
                bulkCopy.WriteToServer(dbfDataReader);
                rowsCopied = bulkCopy.RowsCopied();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing: dbf file: '{options.Dbf}', exception: {ex.Message}");
            }

            stopwatch.Stop();
            Console.WriteLine($"Bulk copy completed in {GetElapsedTime(stopwatch)}s");
            Console.WriteLine($"Copied {rowsCopied} of {dbfRecordCount} rows");
        }

        /// <summary>
        /// Get Elapsed Time
        /// </summary>
        /// <param name="stopwatch"></param>
        /// <returns></returns>
        private static string GetElapsedTime(Stopwatch stopwatch)
        {
            var ts = stopwatch.Elapsed;
            return $"{ts.Hours:00}:{ts.Minutes:00}:{ts.Seconds:00}.{ts.Milliseconds / 10:00}";
        }

        /// <summary>
        /// Build Connection String
        /// </summary>
        /// <param name="options"></param>
        /// <returns></returns>
        public  static string BuildConnectionString(Options options)
        {
            return
                $"Server={options.Server};Database={options.Database};User ID={options.UserId};Password={options.Password};";
        }

        /// <summary>
        /// Ge Columns Details
        /// </summary>
        public static void GetColumnsDetails()
        {
            var dbfPath = "D:\\OCS\\RBUCH.DBF";
            using var dbfTable = new DbfTable(dbfPath, Encoding.UTF8);
            var header = dbfTable.Header;

            var versionDescription = header.VersionDescription;
            var hasMemo = dbfTable.Memo != null;
            var recordCount = header.RecordCount;

            foreach (var dbfColumn in dbfTable.Columns)
            {
                var name = dbfColumn.ColumnName;
                var columnType = dbfColumn.ColumnType;
                var length = dbfColumn.Length;
                var decimalCount = dbfColumn.DecimalCount;
            }
        }

        public static void ReadRecordsAsDataReader()
        {
            var options = new DbfDataReaderOptions
            {
                SkipDeletedRecords = true
                // Encoding = EncodingProvider.GetEncoding(1252);
            };

            var dbfPath = "D:\\OCS\\RBUCH.DBF";
            using var dbfDataReader = new DbfDataReader.DbfDataReader(dbfPath, options);
            var records = new List<object>();
            while (dbfDataReader.Read())
            {
                var record = new
                {
                    RZELLE = dbfDataReader.GetInt32(0),
                    DATUM = dbfDataReader.GetDateTime(1),
                    ZEIT = dbfDataReader.GetString(2),
                    ROHSTOFF = dbfDataReader.GetInt32(3),
                    BESTANDALT = dbfDataReader.GetDecimal(4),
                    BUCHUNG = dbfDataReader.GetDecimal(5),
                    BENUTZER = dbfDataReader.GetString(6),
                    BEMERKUNG = dbfDataReader.GetString(7),
                    ZU = dbfDataReader.GetString(8),
                    DOSANW = dbfDataReader.GetInt32(9)
                };
                records.Add(record);
            }
        }
    }
}