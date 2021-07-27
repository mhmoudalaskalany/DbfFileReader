using System;
using System.Data.SqlClient;
using System.Reflection;

namespace FileReader.Extensions
{
    public static class SqlBulkCopyExtension
    {
        const String RowsCopiedFieldName = "_rowsCopied";
        static FieldInfo _rowsCopiedField = null;

        public static int RowsCopied(this SqlBulkCopy bulkCopy)
        {
            if (_rowsCopiedField == null) _rowsCopiedField = typeof(SqlBulkCopy).GetField(RowsCopiedFieldName, BindingFlags.NonPublic | BindingFlags.GetField
                | BindingFlags.Instance);            
            return (int)_rowsCopiedField.GetValue(bulkCopy);
        }
    }
}