using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using RSI.Data.Utility;


namespace RSI.Data
{
    public static class Extension
    {
        public static List<DatabaseParameter> ApplySecurityColumn<V>(this List<DatabaseParameter> source, V lstSecurityColumnInfo)
            where  V : List<SecurityColumnInfo>, new ()
        {
            if (lstSecurityColumnInfo != null && lstSecurityColumnInfo.Count != 0)
            {
                foreach (SecurityColumnInfo securityColumnInfo in lstSecurityColumnInfo)
                {
                    DatabaseParameter databaseParameter = source.SingleOrDefault(obj => obj.ParameterName == securityColumnInfo.FieldName);
                    if (databaseParameter != null)
                    {
                        EncryptUtility.EncryptType encryptType = (EncryptUtility.EncryptType)Enum.Parse(typeof(EncryptUtility.EncryptType), securityColumnInfo.EncryptType);
                        databaseParameter.DbType = System.Data.DbType.AnsiString;
                        databaseParameter.Value = EncryptUtility.Encrypt(encryptType, Convert.ToString(databaseParameter.Value), securityColumnInfo.EncryptKey, securityColumnInfo.EncryptIV);
                    }
                }
            }

            return source;
        }
    }
}
