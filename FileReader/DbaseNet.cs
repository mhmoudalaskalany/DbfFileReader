using System;
using dBASE.NET;

namespace FileReader
{
    public static class DbaseNet
    {
        private static void ReadDbfFile()
        {
            var dbf = new Dbf();
            dbf.Read("D:\\OCS\\RBUCH.DBF");
            Console.WriteLine(dbf.Records.Count);
        }
    }
}