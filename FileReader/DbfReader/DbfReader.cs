using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using FileReader.Extensions;
using FileReader.Models;
using Microsoft.Extensions.Configuration;

namespace FileReader.DbfReader
{
    public class DbfReader
    {
        private static Options _options;
        private static readonly DateTime Day = new DateTime(2021, 7, 27);
        private static readonly IConfiguration Configuration = AppConfiguration.ReadConfigurationFromAppSettings();
        #region Private Methods

        /// <summary>
        /// ReadUser Inputs
        /// </summary>
        public  static void ReadUserInput()
        {
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("Welcome To DBF File Reader (OCS-INFOTECH)");
            Console.WriteLine($"Starting Reading File At {DateTime.Now} .....");
            Console.ForegroundColor = ConsoleColor.Blue;

            Console.WriteLine($"Day Data : {Day}");
            Console.WriteLine("Please Enter Option For Files");
            Console.WriteLine("1 - AUFTRAGA");
            Console.WriteLine("2 - AUFTRAGB");
            Console.WriteLine("3 - AUFTRAGC");
            Console.WriteLine("4 - RBUCH");
            Console.WriteLine("5 - LFS");
            Console.WriteLine("6 - Generate Excel");
            Console.WriteLine("7 - Generate Summary Excel");
            Console.WriteLine("8 - Exit");
            switch (Console.ReadLine())
            {
                case "1":
                    {
                        _options = new Options("AUFTRAGA", "AUFTRAGA");
                        DbfFileDataReader.RunAndReturnExitCode(_options);
                        ReadUserInput();
                        break;
                    }
                case "2":
                    {
                        _options = new Options("AUFTRAGB", "AUFTRAGB");
                        DbfFileDataReader.RunAndReturnExitCode(_options);
                        ReadUserInput();
                        break;
                    }
                case "3":
                    {
                        _options = new Options("AUFTRAGC", "AUFTRAGC");
                        DbfFileDataReader.RunAndReturnExitCode(_options);
                        ReadUserInput();
                        break;
                    }
                case "4":
                    {
                        _options = new Options("RBUCH", "RBUCH");
                        DbfFileDataReader.RunAndReturnExitCode(_options);
                        ReadUserInput();
                        break;
                    }
                case "5":
                    {
                        _options = new Options("LFS", "LFS");
                        DbfFileDataReader.RunAndReturnExitCode(_options);
                        ReadUserInput();
                        break;
                    }
                case "6":
                    {
                        _options = new Options();
                        GenerateFinalExcel();
                        ReadUserInput();
                        break;
                    }
                case "7":
                    {
                        _options = new Options();
                        GenerateSummaryExcel();
                        ReadUserInput();
                        break;
                    }
                case "8":
                    {
                        break;
                    }
            }
        }


        ////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////// Header Excel Sheet /////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////

        #region Header Excel Sheet


        /// <summary>
        /// Generate Final Excel 
        /// </summary>
        private static void GenerateFinalExcel()
        {
            var data = GetFinalData();
            var finalData = PrepareExcelData(data);
            GenerateFile(finalData, Day);
        }


        /// <summary>
        /// Prepare FinalData
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static List<FinalModel> PrepareExcelData(List<CombinedModel> data)
        {
            var finalData = new List<FinalModel>();
            var grouped = data.GroupBy(a => a.ProcessOrderNumber);
            foreach (var group in grouped)
            {
                var finalModel = new FinalModel
                {
                    FgMaterialCode = group.First().FgMaterialCode,
                    ProductVersion = group.First().ProductVersion,
                    PlantCode = "",
                    ProductionLine = "",
                    ProcessOrderType = "",
                    ProcessOrderNumber = group.Key,
                    TotalQty = group.First().Charge * group.First().BatchCount,
                    UnitOfMeasure = "KG",
                    BatchCount = group.First().BatchCount,
                    EndDate = group.Last().EndDate,
                    EndTime = group.Last().EndTime,
                    StartDate = group.First().StartDate,
                    StartTime = group.First().StartTime,
                    RawCodeMaterial = group.First().RawCodeMaterial,
                    ActualQty = group.Sum(x => x.CurrentQuantity)
                };
                finalData.Add(finalModel);
            }

            return finalData;
        }

        /// <summary>
        /// Generating Excel File For Final Model Data
        /// </summary>
        /// <param name="entities"></param>
        /// <param name="day"></param>
        private static void GenerateFile(List<FinalModel> entities, DateTime day)
        {
            if (entities.Any())
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine($"Generating Excel File Starting From Day : {day.ToShortDateString()}...");
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

                foreach (var record in entities)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = record.FgMaterialCode;
                    worksheet.Cell(currentRow, 2).Value = record.ProductVersion;
                    worksheet.Cell(currentRow, 3).Value = record.PlantCode;
                    worksheet.Cell(currentRow, 4).Value = record.ProductionLine;
                    worksheet.Cell(currentRow, 5).Value = record.ProcessOrderType;
                    worksheet.Cell(currentRow, 6).Value = record.ProcessOrderNumber;
                    worksheet.Cell(currentRow, 7).Value = record.TotalQty;
                    worksheet.Cell(currentRow, 8).Value = record.UnitOfMeasure;
                    worksheet.Cell(currentRow, 9).Value = record.BatchCount;
                    worksheet.Cell(currentRow, 10).Value = record.EndDate;
                    worksheet.Cell(currentRow, 11).Value = record.EndTime;
                    worksheet.Cell(currentRow, 12).Value = record.StartDate;
                    worksheet.Cell(currentRow, 13).Value = record.StartTime;
                    worksheet.Cell(currentRow, 14).Value = record.RawCodeMaterial;
                    worksheet.Cell(currentRow, 15).Value = record.ActualQty;
                }

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                var content = stream.ToArray();
                var path = Configuration["ExcelPath"];
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Writing Excel File To Path {path}");
                bool exists = Directory.Exists(path);

                if (!exists)
                    Directory.CreateDirectory(path);
                File.WriteAllBytes($"{path}PrdHdr_{day.Year}{day.Month}{day.Day}.xlsx", content);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Writing File Success");
            }
        }


        /// <summary>
        /// Get Final Data From Sql
        /// </summary>
        /// <returns></returns>
        private static List<CombinedModel> GetFinalData()
        {
            try
            {
                var connectionString = DbfFileDataReader.BuildConnectionString(_options);
                var connection = new SqlConnection(connectionString);
                var query = $"GetLastDayData";
                var command = new SqlCommand(query) { CommandType = CommandType.StoredProcedure };
                command.Parameters.AddWithValue("@Day", Day);
                command.Connection = connection;
                connection.Open();
                using SqlDataReader dr = command.ExecuteReader();
                var list = dr.MapToList<CombinedModel>();
                Console.WriteLine(list.Count);
                return list;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        #endregion


        ////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////// Summary Excel Sheet /////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////
        #region Summary Excel Sheet

        /// <summary>
        /// Generate Summary Excel Sheet
        /// </summary>
        private static void GenerateSummaryExcel()
        {
            var data = GetSummaryExcelData();
            var finalData = PrepareSummaryExcelData(data);
            GenerateFinalSummaryExcel(finalData, Day);
        }

        /// <summary>
        /// Get Summary Excel Data From Database
        /// </summary>
        /// <returns></returns>
        private static List<SummaryModel> GetSummaryExcelData()
        {
            var connectionString = DbfFileDataReader.BuildConnectionString(_options);
            var connection = new SqlConnection(connectionString);
            var query = $"GetSummaryLastDayData";
            var command = new SqlCommand(query) { CommandType = CommandType.StoredProcedure };
            command.Parameters.AddWithValue("@Day", Day);
            command.Connection = connection;
            connection.Open();
            using SqlDataReader dr = command.ExecuteReader();
            var list = dr.MapToList<SummaryModel>();
            Console.WriteLine(list.Count);
            return list;
        }

        /// <summary>
        /// Map Data To Excel Model
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static List<FinalSummaryModel> PrepareSummaryExcelData(List<SummaryModel> data)
        {
            var newList = data
                .GroupBy(x => new { x.ProcessOrderNumber, x.RawMaterialCode, })
                .Select(y => new FinalSummaryModel()
                {
                    ProcessOrderNumber = y.Key.ProcessOrderNumber,
                    RawMaterialCode = y.Key.RawMaterialCode,
                    StdQty = y.Sum(x => x.StartQty),
                    ActualQty = y.Sum(x => x.CurrentQty)
                }
                ).OrderBy(x => x.ProcessOrderNumber).ThenBy(x => x.RawMaterialCode).ToList();
            return newList;
        }

        /// <summary>
        /// Generate Final Summary Excel
        /// </summary>
        /// <param name="entities"></param>
        /// <param name="day"></param>
        private static void GenerateFinalSummaryExcel(List<FinalSummaryModel> entities, DateTime day)
        {
            if (entities.Any())
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine($"Generating Excel File Starting From Day : {day.ToShortDateString()}...");
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

                foreach (var record in entities)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = record.ProcessOrderNumber;
                    worksheet.Cell(currentRow, 2).Value = record.RawMaterialCode;
                    worksheet.Cell(currentRow, 3).Value = record.StdQty;
                    worksheet.Cell(currentRow, 4).Value = record.ActualQty;
                }

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                var content = stream.ToArray();
                var path = Configuration["ExcelPath"];
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Writing Excel File To Path {path}");
                bool exists = Directory.Exists(path);

                if (!exists)
                    Directory.CreateDirectory(path);
                File.WriteAllBytes($"{path}PrdDtl_{day.Year}{day.Month}{day.Day}.xlsx", content);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Writing File Success");
            }
        }
        #endregion


        // /// <summary>
        // /// Getting Data For AUFTRAGA Excel Data
        // /// </summary>
        // /// <param name="day"></param>
        // /// <returns></returns>
        // private static List<AufTragA> GetDataForAufTragAExcel(DateTime day)
        // {
        //     var connectionString = DbfFileDataReader.BuildConnectionString(_options);
        //     var connection = new SqlConnection(connectionString);
        //     var command = new SqlCommand($"select * from {_options.Table} where DATUM >= @yesterday");
        //     command.Parameters.AddWithValue("@yesterday", day);
        //     command.Connection = connection;
        //     connection.Open();
        //     using SqlDataReader dr = command.ExecuteReader();
        //     var list = dr.MapToList<AufTragA>();
        //     Console.WriteLine(list.Count);
        //     return list;
        // }
        //
        // /// <summary>
        // /// Getting Data For AUFTRAGC Excel Data
        // /// </summary>
        // /// <param name="day"></param>
        // /// <returns></returns>
        // private static List<AufTragB> GetDataForAufTragBExcel(DateTime day)
        // {
        //     var connectionString = DbfFileDataReader.BuildConnectionString(_options);
        //     var connection = new SqlConnection(connectionString);
        //     var command = new SqlCommand($"select * from {_options.Table} where DATUM >= @yesterday");
        //     command.Parameters.AddWithValue("@yesterday", day);
        //     command.Connection = connection;
        //     connection.Open();
        //     using SqlDataReader dr = command.ExecuteReader();
        //     var list = dr.MapToList<AufTragB>();
        //     Console.WriteLine(list.Count);
        //     return list;
        // }
        //
        // /// <summary>
        // /// Getting Data For AUFTRAGC Excel Data
        // /// </summary>
        // /// <param name="day"></param>
        // /// <returns></returns>
        // private static List<AufTragC> GetDataForAufTragCExcel(DateTime day)
        // {
        //     var connectionString = DbfFileDataReader.BuildConnectionString(_options);
        //     var connection = new SqlConnection(connectionString);
        //     var command = new SqlCommand($"select * from {_options.Table} where DATUM >= @yesterday");
        //     command.Parameters.AddWithValue("@yesterday", day);
        //     command.Connection = connection;
        //     connection.Open();
        //     using SqlDataReader dr = command.ExecuteReader();
        //     var list = dr.MapToList<AufTragC>();
        //     Console.WriteLine(list.Count);
        //     return list;
        // }




        #endregion
    }
}
