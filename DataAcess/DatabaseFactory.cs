using System;
using System.Configuration;
using System.Data.Common;

namespace RSI.Data
{
    /// <summary>
    /// 資料庫連線工具
    /// </summary>
    public static class DatabaseFactory
    {
        /// <summary>
        /// 連線字串設定
        /// </summary>
        /// <param name="connectionName">連線字串名稱</param>
        /// <returns></returns>
        private static ConnectionStringSettings GetConnectionSettings(string connectionName)
        {
            ConnectionStringSettings connectionSettings;
            connectionSettings = ConfigurationManager.ConnectionStrings[connectionName];
            if (connectionSettings == null)
            {
                throw new ApplicationException("Invalid Database Name");
            }

            return connectionSettings;
        }

        /// <summary>
        /// 連線字串設定
        /// </summary>
        /// <returns></returns>
        private static ConnectionStringSettings GetConnectionSettings()
        {
            string connectionName = ConfigurationManager.AppSettings["DefaultConnectionName"];

            return GetConnectionSettings(connectionName);
        }

        /// <summary>
        /// 建立資料庫連線實體
        /// </summary>
        /// <param name="connectionName"></param>
        /// <returns>連線字串名稱</returns>
        public static Database GetDatabase(string connectionName)
        {
            Database database;
            ConnectionStringSettings connectionSettings;
            connectionSettings = GetConnectionSettings(connectionName);
            string dbProviderName = connectionSettings.ProviderName;

            switch (dbProviderName)
            {
                case "System.Data.SqlClient":
                    database = new SqlServerDatabase(connectionSettings);
                    break;
                case "Oracle.DataAccess.Client":
                    database = new OracleDatabase(connectionSettings);
                    break;
                default:
                    throw (new Exception("Not Supported DbProviderName[" + dbProviderName + "]"));
            }
            return database;
        }

        /// <summary>
        /// 建立資料庫連線實體
        /// </summary>
        /// <param name="connectionSetting"></param>
        /// <returns>連線字串名稱</returns>
        public static Database GetDatabase(ConnectionStringSettings connectionSetting)
        {
            Database database;
            string dbProviderName = connectionSetting.ProviderName;

            switch (dbProviderName)
            {
                case "System.Data.SqlClient":
                    database = new SqlServerDatabase(connectionSetting);
                    break;
                case "Oracle.DataAccess.Client":
                    database = new OracleDatabase(connectionSetting);
                    break;
                default:
                    throw (new Exception("Not Supported DbProviderName[" + dbProviderName + "]"));
            }
            return database;
        }

        /// <summary>
        /// 建立資料庫連線實體
        /// </summary>
        /// <returns></returns>
        public static Database GetDatabase()
        {
            Database database;
            ConnectionStringSettings connectionSettings;
            connectionSettings = GetConnectionSettings();
            string dbProviderName = connectionSettings.ProviderName;

            switch (dbProviderName)
            {
                case "System.Data.SqlClient":
                    database = new SqlServerDatabase(connectionSettings);
                    break;
                case "Oracle.DataAccess.Client":
                    database = new OracleDatabase(connectionSettings);
                    break;
                default:
                    throw (new Exception("Not Supported DbProviderName[" + dbProviderName + "]"));
            }
            return database;
        }

    }
}
