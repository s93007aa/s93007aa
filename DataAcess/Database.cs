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
    /// ��Ʈw����
    /// </summary>
    public abstract class Database : IDisposable
    {
        private DatabaseType _databaseType;
        private ConnectionStringSettings _connectionSettings;

        /// <summary>
        /// ��Ʈw����غc��
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
        /// ��Ʈw����
        /// </summary>
        public DatabaseType Type
        {
            get
            {
                return _databaseType;
            }
        }

        /// <summary>
        /// ��Ʈw�s�u
        /// </summary>
        protected DbConnection DbConnection { get; set; }

        /// <summary>
        /// ��Ʈw���
        /// </summary>
        protected DbTransaction DbTransaction { get; set; }

        /// <summary>
        /// �}�l��Ʈw���
        /// </summary>
        /// <param name="connection"></param>
        /// <returns></returns>
        public abstract void BeginTransaction();

        /// <summary>
        /// �{�i��Ʈw���
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
        /// �_���Ʈw���
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
        /// ����d��
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public abstract DataTable ExecuteQuery(string sqlString, List<DatabaseParameter> parameterList, int timeout = 120);

        /// <summary>
        /// �����Ʈw�d�� (�w�� Oracle RefCursor �Ѽƫ��O, �Шϥ� DbType.Object) 
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public abstract DataTable ExecuteQueryNoCache(string sqlString, List<DatabaseParameter> parameterList);

        /// <summary>
        /// �w���Ʈw���� SQL ���z��
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public abstract int ExecuteNonQuery(string sqlString, List<DatabaseParameter> parameterList);

        /// <summary>
        /// �w���Ʈw����s�W���O,�^�� identity ����
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <param name="identityFieldName"></param>
        /// <returns></returns>
        public abstract long ExecuteInsert(string sqlString, List<DatabaseParameter> parameterList, string identityFieldName);

        /// <summary>
        /// ����d�ߨöǦ^�Ĥ@�C�Ĥ@�Ӹ�Ʀ�
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public abstract object ExecuteScalar(string sqlString, List<DatabaseParameter> parameterList);

        /// <summary>
        /// �����Ʈw�w�s�{��
        /// </summary>
        /// <param name="procedureName"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public abstract DataSet ExecuteProcedure(string procedureName, List<DatabaseParameter> parameterList);

        /// <summary>
        /// �j�q�s�W���
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
