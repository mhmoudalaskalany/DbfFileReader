using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;

namespace FileReader
{
    public static class DbExtension
    {
        public static List<T> MapToList<T>(this SqlDataReader dr) where T : new()
        {
            List<T> RetVal = null;
            var entity = typeof(T);
            var propDic = new Dictionary<string, PropertyInfo>();
            try
            {
                if (dr != null && dr.HasRows)
                {
                    RetVal = new List<T>();
                    var props = entity.GetProperties(BindingFlags.Instance | BindingFlags.Public);
                    propDic = props.ToDictionary(p => p.Name.ToUpper(), p => p);
                    while (dr.Read())
                    {
                        T newObject = new T();
                        for (int index = 0; index < dr.FieldCount; index++)
                        {
                            if (propDic.ContainsKey(dr.GetName(index).ToUpper()))
                            {
                                var info = propDic[dr.GetName(index).ToUpper()];
                                if (info != null && info.CanWrite)
                                {
                                    var val = dr.GetValue(index);
                                    info.SetValue(newObject, (val == DBNull.Value) ? null : val, null);
                                }
                            }
                        }
                        RetVal.Add(newObject);
                    }
                }

                return RetVal;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        public static List<T> MapToList<T>(this OleDbDataReader dr) where T : new()
        {
            List<T> RetVal = null;
            var entity = typeof(T);
            var propDic = new Dictionary<string, PropertyInfo>();
            try
            {
                if (dr != null && dr.HasRows)
                {
                    RetVal = new List<T>();
                    var props = entity.GetProperties(BindingFlags.Instance | BindingFlags.Public);
                    propDic = props.ToDictionary(p => p.Name.ToUpper(), p => p);
                    while (dr.Read())
                    {
                        T newObject = new T();
                        for (int index = 0; index < dr.FieldCount; index++)
                        {
                            if (propDic.ContainsKey(dr.GetName(index).ToUpper()))
                            {
                                var info = propDic[dr.GetName(index).ToUpper()];
                                if (info != null && info.CanWrite)
                                {
                                    var val = dr.GetValue(index);
                                    info.SetValue(newObject, (val == DBNull.Value) ? null : val, null);
                                }
                            }
                        }
                        RetVal.Add(newObject);
                    }
                }

                return RetVal;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}