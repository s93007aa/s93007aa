using System;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;

namespace RSI.Data
{
    /// <summary>
    /// 資料庫實體
    /// </summary>
    public abstract class Database : IDisposable
    {
        private DatabaseType _databaseType;
        private ConnectionStringSettings _connectionSettings;

        /// <summary>
        /// 資料庫實體建構式
        /// </summary>
        /// <param name="connectionString">The connection string for the database.</param>
        protected Database(ConnectionStringSettings connectionSettings)
        {
            _connectionSettings = connectionSettings;
            switch (connectionSettings.ProviderName)
            {
                case "System.Data.SqlClient":
                    _databaseType = DatabaseType.SQLServer;
                    break;
                case "Oracle.DataAccess.Client":
                    _databaseType = DatabaseType.Oracle;
                    break;
            }
        }
        
        /// <summary>
        /// 資料庫類型
        /// </summary>
        public DatabaseType Type
        {
            get
            {
                return _databaseType;
            }
        }

        /// <summary>
        /// 資料庫連線
        /// </summary>
        protected DbConnection DbConnection { get; set; }

        /// <summary>
        /// 資料庫交易
        /// </summary>
        protected DbTransaction DbTransaction { get; set; }

        /// <summary>
        /// 開始資料庫交易
        /// </summary>
        /// <param name="connection"></param>
        /// <returns></returns>
        public abstract void BeginTransaction();

        /// <summary>
        /// 認可資料庫交易
        /// </summary>
        /// <param name="transaction"></param>
        public void CommitTransaction()
        {
            if (DbTransaction != null)
            {
                DbTransaction.Commit();
                DbTransaction.Dispose();
                DbTransaction = null;
            }
        }

        /// <summary>
        /// 復原資料庫交易
        /// </summary>
        /// <param name="transaction"></param>
        public void RollBackTransaction()
        {
            if (DbTransaction != null)
            {
                DbTransaction.Rollback();
                DbTransaction.Dispose();
                DbTransaction = null;
            }
        }

        /// <summary>
        /// 執行查詢
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public abstract DataTable ExecuteQuery(string sqlString, List<DatabaseParameter> parameterList, int timeout = 120);

        /// <summary>
        /// 執行資料庫查詢 (針對 Oracle RefCursor 參數型別, 請使用 DbType.Object) 
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public abstract DataTable ExecuteQueryNoCache(string sqlString, List<DatabaseParameter> parameterList);

        /// <summary>
        /// 針對資料庫執行 SQL 陳述式
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public abstract int ExecuteNonQuery(string sqlString, List<DatabaseParameter> parameterList);

        /// <summary>
        /// 針對資料庫執行新增指令,回傳 identity 欄位值
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <param name="identityFieldName"></param>
        /// <returns></returns>
        public abstract long ExecuteInsert(string sqlString, List<DatabaseParameter> parameterList, string identityFieldName);

        /// <summary>
        /// 執行查詢並傳回第一列第一個資料行
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public abstract object ExecuteScalar(string sqlString, List<DatabaseParameter> parameterList);

        /// <summary>
        /// 執行資料庫預存程序
        /// </summary>
        /// <param name="procedureName"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public abstract DataSet ExecuteProcedure(string procedureName, List<DatabaseParameter> parameterList);

        /// <summary>
        /// 大量新增資料
        /// </summary>
        /// <param name="destinationTableName"></param>
        /// <param name="dtData"></param>
        public abstract void BulkCopy(string destinationTableName, DataTable dtData);

        public void Dispose()
        {
            if (DbTransaction == null)
                this.DbConnection.Dispose();
        }

        public abstract DataTable GetDtStructure(string tblName);
        public abstract DataTable GetDtByID(string tblName, string ID);
        public abstract int AddData(DataTable obj);
        public abstract int UpdateData(DataTable obj);
        public abstract string GetID(string SeqName);
        public abstract int DeleteData(DataTable obj);
    }
}
