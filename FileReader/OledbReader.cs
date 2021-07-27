using System;
using System.Data.OleDb;
using System.Linq;
using FileReader.Extensions;
using FileReader.Models;
using Microsoft.Extensions.Configuration;

namespace FileReader
{
    public static class OledbReader
    {
        private static IConfiguration _configuration = AppConfiguration.ReadConfigurationFromAppSettings();
        private static string _connection = _configuration["ConnectionString"];

        public static void ReadOledbFile()
        {
            try
            {
                var yesterday = DateTime.Now.AddDays(-30);
#pragma warning disable CA1416 // Validate platform compatibility
                using OleDbConnection connection = new OleDbConnection(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;" +
                    @"Data Source=D:\OCS\;" +
                    @"Extended Properties=dBASE III;");
#pragma warning restore CA1416 // Validate platform compatibility

                connection.Open();
                //var sql = "SELECT * FROM RBUCH where DATUM > #2/16/2019 12:00:00 AM#";
                var sql = "SELECT * FROM RBUCH1";
                using OleDbCommand command = new OleDbCommand(sql , connection);
                using OleDbDataReader dr = command.ExecuteReader();
                var list = dr.MapToList<DbfRecord>();
                var filteredList = list.Where(x => x.DATUM >= yesterday).ToList();
                Console.WriteLine(list.Count);
                //var schemaTable = dr.GetSchemaTable(); // Get Metadata of the current table.
                //var dt = new DataTable();
                //foreach (DataRow row in schemaTable.Rows) // Copy the schema to your datatable object
                //{
                //    string colName = row.Field<string>("ColumnName");
                //    Type t = row.Field<Type>("DataType");
                //    dt.Columns.Add(colName, t);
                //}
                //bool hasNextRow = false;
                //do
                //{
                //    if (!hasNextRow)
                //    {
                //        // We have to do this in order to peek through the next row. If we do not have next row, then we will have to commit the current changes
                //        hasNextRow = dr.Read();
                //    }
                //    if (!hasNextRow) break; // Break if there is no row.
                //    var newRow = dt.NewRow();
                //    foreach (DataColumn col in dt.Columns) // Copy current row
                //    {
                //        newRow[col.ColumnName] = dr[col.ColumnName];
                //    }
                //    dt.Rows.Add(newRow);
                //    dt.AcceptChanges();
                //    hasNextRow = dr.Read();
                //    if (dt.Rows.Count >= 1000 || !hasNextRow) // When the data exceeds thousands rows Or when there are no further rows, insert the data into sql and clear the memory. 
                //    {
                //        foreach (DataRow row in dt.Rows) // Adjust the values
                //        {
                //            for (int i = 0; i < dt.Columns.Count; i++)
                //            {
                //                if (dt.Columns[i].DataType == typeof(string))
                //                    row[i] = Regex.Replace(row[i].ToString(), "[#$%^*@!~?]", "");
                //            }
                //        }
                //        dt.AcceptChanges();
                //        //// Bulk Copy to SQL Server
                //        //using (SqlBulkCopy bulkCopy = new SqlBulkCopy(_connection)
                //        //{
                //        //    DestinationTableName = "RBUCH"
                //        //})
                //        //{
                //        //    // code for bulk insert
                //        //    bulkCopy.WriteToServer(dt);
                //        //}

                //        dt.Rows.Clear(); // Clear the memory so new rows can be read.
                //        dt.AcceptChanges();
                //    }
                //} while (hasNextRow);

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

        }

        

        
    }
}