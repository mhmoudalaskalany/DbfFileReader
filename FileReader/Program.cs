using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using FileReader.Models;
using Microsoft.Extensions.Configuration;


namespace FileReader
{
    public class Program
    {
        private static Options _options;
        private static IConfiguration _configuration = AppConfiguration.ReadConfigurationFromAppSettings();
        public static void Main(string[] args)
        {
            Console.WriteLine($"");

            _options = new Options();
            //ReadDbfFile();
            //DbfFileDataReader.GetColumnsDetails();
            DbfFileDataReader.RunAndReturnExitCode(_options);
            //OledbReader.ReadOledbFile();
            var lastDayData = GetDataForExcel();
            GenerateExcelFile(lastDayData);
        }

        private static List<SqlRecord> GetDataForExcel()
        {
            var yesterday = DateTime.Now.AddDays(-30);
            var connectionString = DbfFileDataReader.BuildConnectionString(_options);
            var connection = new SqlConnection(connectionString);
            var command = new SqlCommand($"select * from {_options.Table} where DATUM >= @yesterday");
            command.Parameters.AddWithValue("@yesterday", yesterday);
            command.Connection = connection;
            connection.Open();
            using SqlDataReader dr = command.ExecuteReader();
            var list = dr.MapToList<SqlRecord>();
            Console.WriteLine(list.Count);
            return list;

        }

        private static void GenerateExcelFile(List<SqlRecord> entities)
        {
            if (entities.Any())
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Last Day");
                var currentRow = 1;
                var properties = entities.First().GetType().GetProperties();
                var columnNumber = 0;
                // set header columns
                foreach (var prop in properties)
                {
                    worksheet.Cell(currentRow, ++columnNumber).Value = prop.Name;
                }

                foreach (var user in entities)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = user.RZELLE;
                    worksheet.Cell(currentRow, 2).Value = user.DATUM;
                    worksheet.Cell(currentRow, 3).Value = user.ZEIT;
                    worksheet.Cell(currentRow, 4).Value = user.ROHSTOFF;
                    worksheet.Cell(currentRow, 5).Value = user.BESTANDALT;
                    worksheet.Cell(currentRow, 6).Value = user.BUCHUNG;
                    worksheet.Cell(currentRow, 7).Value = user.BENUTZER;
                    worksheet.Cell(currentRow, 8).Value = user.BEMERKUNG;
                    worksheet.Cell(currentRow, 9).Value = user.ZU;
                    worksheet.Cell(currentRow, 10).Value = user.DOSANW;
                }

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                var content = stream.ToArray();
                var path = _configuration["ExcelPath"];
                bool exists = Directory.Exists(path);

                if (!exists)
                    System.IO.Directory.CreateDirectory(path);
                File.WriteAllBytes(path + "Last Day.xlsx", content);
            }
            
        }
    }
}