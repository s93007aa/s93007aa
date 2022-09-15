using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using Oracle.ManagedDataAccess.Client;
using System.Text;

namespace RSI.Data
{
    /// <summary>
    /// Oracle ��Ʈw����
    /// </summary>
    public class OracleDatabase : Database
    {
        /// <summary>
        /// Oracle ��Ʈw�s���u��غc��
        /// </summary>
        /// <param name="connectionSettings"></param>
        public OracleDatabase(ConnectionStringSettings connectionSettings)
            : base(connectionSettings)
        {
            DbConnection = new OracleConnection(connectionSettings.ConnectionString);
        }

        /// <summary>
        /// �}�l��Ʈw���
        /// </summary>
        /// <param name="connection"></param>
        /// <returns></returns>
        public override void BeginTransaction()
        {
            OracleConnection dbConn = (OracleConnection)DbConnection;
            if (dbConn.State == ConnectionState.Closed)
                dbConn.Open();
            if (DbTransaction == null)
                DbTransaction = dbConn.BeginTransaction();
        }

        /// <summary>
        /// ���o Oracle Command
        /// </summary>
        /// <returns></returns>
        public OracleCommand GetOracleCommd()
        {
            OracleCommand oracleCmd = new OracleCommand();
            oracleCmd.Connection = (OracleConnection)DbConnection;
            oracleCmd.BindByName = true;

            return oracleCmd;
        }

        /// <summary>
        /// �����Ʈw�d�� (�w�� Oracle RefCursor �Ѽƫ��O, �Шϥ� DbType.Onject) 
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public override DataTable ExecuteQuery(string sqlString, List<DatabaseParameter> parameterList, int timeout = 120)
        {
            List<OracleParameter> oracleParameterList = ConvertParameterList(parameterList);

            return ExecuteQuery(sqlString, oracleParameterList);
        }

        /// <summary>
        /// �����Ʈw�d�� (�w�� Oracle �S��Ѽƫ��O, �Ҧp CLOB, BLOB, BFILE)
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public DataTable ExecuteQuery(string sqlString, List<OracleParameter> parameterList)
        {
            DataTable dtResult = new DataTable();
            string sql = FormatSqlString(sqlString);
            OracleParameter[] oracleParameterList = new OracleParameter[parameterList.Count];
            parameterList.CopyTo(oracleParameterList, 0);
            OracleCommand cmd = new OracleCommand(sql, (OracleConnection)DbConnection);
            OracleDataReader oracleDataReader;

            cmd.AddToStatementCache = true;
            cmd.CommandType = CommandType.Text;
            cmd.BindByName = true;
            cmd.Parameters.AddRange(oracleParameterList);

            if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

            if (DbTransaction == null)
            {
                oracleDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            else
            {
                cmd.Transaction = (OracleTransaction)DbTransaction;
                oracleDataReader = cmd.ExecuteReader();
            }

            dtResult.Load(oracleDataReader, LoadOption.OverwriteChanges);

            if (DbTransaction == null) cmd.Connection.Close();

            cmd.Dispose();
            
            return dtResult;
        }

        /// <summary>
        /// �����Ʈw�d�� (�w�� Oracle RefCursor �Ѽƫ��O, �Шϥ� DbType.Object) 
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public override DataTable ExecuteQueryNoCache(string sqlString, List<DatabaseParameter> parameterList)
        {
            List<OracleParameter> oracleParameterList = ConvertParameterList(parameterList);

            return ExecuteQueryNoCache(sqlString, oracleParameterList);
        }

        /// <summary>
        /// �����Ʈw�d�� (�w�� Oracle �S��Ѽƫ��O, �Ҧp CLOB, BLOB, BFILE)
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public DataTable ExecuteQueryNoCache(string sqlString, List<OracleParameter> parameterList)
        {
            DataTable dtResult = new DataTable();
            string sql = FormatSqlString(sqlString);
            OracleParameter[] oracleParameterList = new OracleParameter[parameterList.Count];
            parameterList.CopyTo(oracleParameterList, 0);
            OracleCommand cmd = new OracleCommand(sql, (OracleConnection)DbConnection);
            OracleDataReader oracleDataReader;

            cmd.AddToStatementCache = false;
            cmd.CommandType = CommandType.Text;
            cmd.BindByName = true;
            cmd.Parameters.AddRange(oracleParameterList);

            if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

            if (DbTransaction == null)
            {
                oracleDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            else
            {
                cmd.Transaction = (OracleTransaction)DbTransaction;
                oracleDataReader = cmd.ExecuteReader();
            }

            dtResult.Load(oracleDataReader, LoadOption.OverwriteChanges);

            if (DbTransaction == null) cmd.Connection.Close();

            cmd.Dispose();

            return dtResult;
        }

        /// <summary>
        /// �w���Ʈw���� SQL ���z��
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public override int ExecuteNonQuery(string sqlString, List<DatabaseParameter> parameterList)
        {
            List<OracleParameter> oracleParameterList = ConvertParameterList(parameterList);

            return ExecuteNonQuery(sqlString, oracleParameterList);
        }

        /// <summary>
        /// �w���Ʈw���� SQL ���z��
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public int ExecuteNonQuery(string sqlString, List<OracleParameter> parameterList)
        {
            int affectedCount = 0;

            string sql = FormatSqlString(sqlString);
            OracleParameter[] oracleParameterList = new OracleParameter[parameterList.Count];
            parameterList.CopyTo(oracleParameterList, 0);
            OracleCommand cmd = new OracleCommand(sql, (OracleConnection)DbConnection);

            cmd.AddToStatementCache = true;
            cmd.CommandType = CommandType.Text;
            cmd.BindByName = true;
            cmd.Parameters.AddRange(oracleParameterList);

            if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

            if (DbTransaction != null) cmd.Transaction = (OracleTransaction)DbTransaction;

            affectedCount = cmd.ExecuteNonQuery();

            if (DbTransaction == null) cmd.Connection.Close();

            cmd.Dispose();

            return affectedCount;
        }

        /// <summary>
        /// �w���Ʈw����s�W���O,�^�� identity ����
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <param name="identityFieldName"></param>
        /// <returns></returns>
        public override Int64 ExecuteInsert(string sqlString, List<DatabaseParameter> parameterList, string identityFieldName)
        {
            Int64 scopeIdentity = -1;

            if (string.IsNullOrEmpty(identityFieldName))
            {
                ExecuteNonQuery(sqlString, parameterList);
            }
            else
            {
                List<OracleParameter> oracleParameterList = ConvertParameterList(parameterList);
                OracleParameter identity = new OracleParameter(identityFieldName, OracleDbType.Int64, ParameterDirection.Output);
                oracleParameterList.Add(identity);
                string sql = string.Format("{0} RETURNING {1} INTO :{1}", sqlString, identityFieldName);

                ExecuteNonQuery(sql, oracleParameterList);
                scopeIdentity = Int64.Parse(identity.Value.ToString());
            }

            return scopeIdentity;
        }

        /// <summary>
        /// ����d�ߨöǦ^�Ĥ@�C�Ĥ@�Ӹ�Ʀ�
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public override object ExecuteScalar(string sqlString, List<DatabaseParameter> parameterList)
        {
            List<OracleParameter> oracleParameterList = ConvertParameterList(parameterList);

            return ExecuteScalar(sqlString, oracleParameterList);
        }

        /// <summary>
        /// ����d�ߨöǦ^�Ĥ@�C�Ĥ@�Ӹ�Ʀ�
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public object ExecuteScalar(string sqlString, List<OracleParameter> parameterList)
        {
            object result;
            string sql = FormatSqlString(sqlString);
            OracleParameter[] oracleParameterList = new OracleParameter[parameterList.Count];
            parameterList.CopyTo(oracleParameterList, 0);
            OracleCommand cmd = new OracleCommand(sql, (OracleConnection)DbConnection);

            cmd.AddToStatementCache = false;
            cmd.CommandType = CommandType.Text;
            cmd.BindByName = true;
            cmd.Parameters.AddRange(oracleParameterList);

            if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

            if (DbTransaction == null)
            {
                result = cmd.ExecuteScalar();
            }
            else
            {
                cmd.Transaction = (OracleTransaction)DbTransaction;
                result = cmd.ExecuteScalar();
            }

            if (DbTransaction == null) cmd.Connection.Close();

            cmd.Dispose();

            return result;
        }

        /// <summary>
        /// �����Ʈw�w�s�{��
        /// </summary>
        /// <param name="procedureName"></param>
        /// <param name="parameterList"></param>
        public override DataSet ExecuteProcedure(string procedureName, List<DatabaseParameter> parameterList)
        {
            DataSet dsResult = new DataSet();

            return dsResult;
        }

        /// <summary>
        /// �j�q�s�W���
        /// </summary>
        /// <param name="destinationTableName"></param>
        /// <param name="dtData"></param>
        public override void BulkCopy(string destinationTableName, DataTable dtData)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// �N��Ʈw�Ѽƶ��X�ഫ�� Oracle �M�ݰѼƶ��X
        /// </summary>
        /// <param name="parameterCollection"></param>
        /// <returns></returns>
        private List<OracleParameter> ConvertParameterList(List<DatabaseParameter> parameterCollection)
        {
            List<OracleParameter> oracleParameterList = new List<OracleParameter>();

            foreach (DbParameter parameter in parameterCollection)
            {
                oracleParameterList.Add(new OracleParameter( parameter.ParameterName, ConvertDbType(parameter.DbType), parameter.Value, parameter.Direction));
            }

            return oracleParameterList;
        }

        /// <summary>
        /// �Ѽƫ��O�ഫ
        /// </summary>
        /// <param name="dbType"></param>
        /// <returns></returns>
        private OracleDbType ConvertDbType(DbType dbType)
        {
            OracleDbType oracleDbType = new OracleDbType();

            switch (dbType)
            {
                case DbType.String:
                    oracleDbType = OracleDbType.NVarchar2;
                    break;
                case DbType.AnsiString:
                    oracleDbType = OracleDbType.Varchar2;
                    break;
                case DbType.Decimal:
                    oracleDbType = OracleDbType.Decimal;
                    break;
                case DbType.Int16:
                    oracleDbType = OracleDbType.Int16;
                    break;
                case DbType.Int32:
                    oracleDbType = OracleDbType.Int32;
                    break;
                case DbType.Int64:
                    oracleDbType = OracleDbType.Int64;
                    break;
                case DbType.Single:
                    oracleDbType = OracleDbType.Single;
                    break;
                case DbType.Double:
                    oracleDbType = OracleDbType.Double;
                    break;
                case DbType.DateTime:
                    oracleDbType = OracleDbType.Date;
                    break;
                case DbType.Object:
                    oracleDbType = OracleDbType.RefCursor;
                    break;
                case DbType.Time:
                    oracleDbType = OracleDbType.TimeStamp;
                    break;
                default:
                    throw new Exception("Not Supported DbType !!");
            }

            return oracleDbType;
        }

        /// <summary>
        /// �N SQL �榡�Ʀ� Oracle �M�ݪ��榡
        /// </summary>
        /// <param name="sqlString"></param>
        /// <returns></returns>
        private string FormatSqlString(string sqlString)
        {
            return sqlString.Replace('@', ':');
        }

        /// <summary>
        /// �s�W�ɨϥ�
        /// </summary>
        /// <param name="tblName"></param>
        /// <returns></returns>
        public override DataTable GetDtStructure(string tblName)
        {
            DataTable dt = new DataTable();

            return dt;
        }

        /// <summary>
        /// ��s�ɨϥ�
        /// </summary>
        /// <param name="tblName"></param>
        /// <param name="ID"></param>
        /// <returns></returns>
        public override DataTable GetDtByID(string tblName, string ID)
        {
            DataTable dt = new DataTable();
            
            return dt;
        }
        public override int AddData(DataTable obj)
        {
            return 0;
        }

        public override int UpdateData(DataTable obj)
        {
            return 0;
        }

        public override string GetID(string SeqName)
        {
            return "";
        }

        public override int DeleteData(DataTable obj)
        {
            return 0;
        }
    }
}
