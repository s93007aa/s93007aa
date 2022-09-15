using System;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Text;
//using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;

namespace RSI.Data
{
    /// <summary>
    /// SQL Server 資料庫實體
    /// </summary>
    public class SqlServerDatabase : Database
    {
        public SqlServerDatabase(ConnectionStringSettings connectionSettings) 
            : base(connectionSettings)
        {
            DbConnection = new SqlConnection(connectionSettings.ConnectionString);
        }

        /// <summary>
        /// 開始資料庫交易
        /// </summary>
        /// <param name="connection"></param>
        /// <returns></returns>
        public override void BeginTransaction()
        {
            SqlConnection dbConn = (SqlConnection)DbConnection;
            if (dbConn.State == ConnectionState.Closed)
                dbConn.Open();
            if (DbTransaction == null)
                DbTransaction = ((SqlConnection)DbConnection).BeginTransaction();
        }

        /// <summary>
        /// 取得 SQL Server Command
        /// </summary>
        /// <returns></returns>
        public SqlCommand GetSqlCommd()
        {
            SqlCommand sqlCmd = new SqlCommand();
            sqlCmd.Connection = (SqlConnection)DbConnection;

            return sqlCmd;
        }

        /// <summary>
        /// 執行資料庫查詢
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public override DataTable ExecuteQuery(string sqlString, List<DatabaseParameter> parameterList, int timeout = 120)
        {
            List<SqlParameter> sqlParameterList = ConvertParameterList(parameterList);

            return ExecuteQuery(sqlString, sqlParameterList, timeout);
        }

        /// <summary>
        /// 執行資料庫查詢
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public DataTable ExecuteQuery(string sqlString, List<SqlParameter> parameterList, int timeout = 120)
        {
            DataTable dtResult = new DataTable();
            string sql = FormatSqlString(sqlString);
            SqlParameter[] sqlParameterList = new SqlParameter[parameterList.Count];
            parameterList.CopyTo(sqlParameterList, 0);
            SqlCommand cmd = new SqlCommand(sql, (SqlConnection)DbConnection);
            SqlDataReader sqlDataReader;

            cmd.CommandTimeout = timeout;
            cmd.CommandType = CommandType.Text;
            cmd.Parameters.AddRange(sqlParameterList);

            if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

            if (DbTransaction == null)
            {
                sqlDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            else
            {
                cmd.Transaction = (SqlTransaction)DbTransaction;
                sqlDataReader = cmd.ExecuteReader();
            }

            dtResult.Load(sqlDataReader, LoadOption.OverwriteChanges);

            if (DbTransaction == null) cmd.Connection.Close();

            cmd.Dispose();

            return dtResult;
        }

        /// <summary>
        /// 執行資料庫查詢 
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public override DataTable ExecuteQueryNoCache(string sqlString, List<DatabaseParameter> parameterList)
        {
            List<SqlParameter> sqlParameterList = ConvertParameterList(parameterList);

            return ExecuteQuery(sqlString, sqlParameterList);
        }

        /// <summary>
        /// 針對資料庫執行 SQL 陳述式
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public override int ExecuteNonQuery(string sqlString, List<DatabaseParameter> parameterList)
        {
            List<SqlParameter> sqlParameterList = ConvertParameterList(parameterList);

            return ExecuteNonQuery(sqlString, sqlParameterList);
        }

        /// <summary>
        /// 針對資料庫執行 SQL 陳述式
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public int ExecuteNonQuery(string sqlString, List<SqlParameter> parameterList)
        {
            int affedtedCount = 0;

            string sql = FormatSqlString(sqlString);
            SqlParameter[] sqlParameterList = new SqlParameter[parameterList.Count];
            parameterList.CopyTo(sqlParameterList, 0);
            SqlCommand cmd = new SqlCommand(sql, (SqlConnection)DbConnection);

            cmd.CommandType = CommandType.Text;
            cmd.Parameters.AddRange(sqlParameterList);

            if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

            if (DbTransaction != null) cmd.Transaction = (SqlTransaction)DbTransaction;

            affedtedCount = cmd.ExecuteNonQuery();

            if (DbTransaction == null) cmd.Connection.Close();

            cmd.Dispose();
            
            return affedtedCount;
        }

        /// <summary>
        /// 針對資料庫執行新增指令,回傳 identity 欄位值
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <param name="identityFieldName"></param>
        /// <returns></returns>
        public override long ExecuteInsert(string sqlString, List<DatabaseParameter> parameterList, string identityFieldName)
        {
            long scopeIdentity = -1;
            string sql = string.Format("{0}; SELECT SCOPE_IDENTITY()", FormatSqlString(sqlString));
            List<SqlParameter> sqlParameterList = ConvertParameterList(parameterList);

            SqlParameter[] sqlParameterArray = new SqlParameter[sqlParameterList.Count];
            sqlParameterList.CopyTo(sqlParameterArray, 0);
            SqlCommand cmd = new SqlCommand(sql, (SqlConnection)DbConnection);

            cmd.CommandType = CommandType.Text;
            cmd.Parameters.AddRange(sqlParameterArray);

            if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

            if (DbTransaction != null) cmd.Transaction = (SqlTransaction)DbTransaction;

            scopeIdentity = Convert.ToInt64(cmd.ExecuteScalar());

            if (DbTransaction == null) cmd.Connection.Close();

            cmd.Dispose();
            
            return scopeIdentity;
        }

        /// <summary>
        /// 執行查詢並傳回第一列第一個資料行
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public override object ExecuteScalar(string sqlString, List<DatabaseParameter> parameterList)
        {
            List<SqlParameter> sqlParameterList = ConvertParameterList(parameterList);

            return ExecuteScalar(sqlString, sqlParameterList);
        }

        /// <summary>
        /// 執行查詢並傳回第一列第一個資料行
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public object ExecuteScalar(string sqlString, List<SqlParameter> parameterList)
        {
            object result;

            string sql = FormatSqlString(sqlString);
            SqlParameter[] sqlParameterList = new SqlParameter[parameterList.Count];
            parameterList.CopyTo(sqlParameterList, 0);
            SqlCommand cmd = new SqlCommand(sql, (SqlConnection)DbConnection);

            cmd.CommandType = CommandType.Text;
            cmd.Parameters.AddRange(sqlParameterList);

            if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

            if (DbTransaction != null) cmd.Transaction = (SqlTransaction)DbTransaction;

            result = cmd.ExecuteScalar();

            if (DbTransaction == null) cmd.Connection.Close();

            cmd.Dispose();

            return result;
        }

        /// <summary>
        /// 執行資料庫預存程序
        /// </summary>
        /// <param name="procedureName"></param>
        /// <param name="parameterList"></param>
        /// <returns></returns>
        public override DataSet ExecuteProcedure(string procedureName, List<DatabaseParameter> parameterList)
        {
            DataSet dsResult = new DataSet();
            SqlParameter[] sqlParameterList = new SqlParameter[parameterList.Count];
            sqlParameterList = ConvertParameterList(parameterList).ToArray();
            SqlCommand cmd = new SqlCommand(procedureName, (SqlConnection)DbConnection);
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);

            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddRange(sqlParameterList);

            if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

            if (DbTransaction != null) cmd.Transaction = (SqlTransaction)DbTransaction;

            adapter.Fill(dsResult);

            if (DbTransaction == null) cmd.Connection.Close();

            cmd.Dispose();

            return dsResult;
        }

        /// <summary>
        /// 大量新增資料
        /// </summary>
        /// <param name="destinationTableName"></param>
        /// <param name="dtData"></param>
        public override void BulkCopy(string destinationTableName, DataTable dtData)
        {
            SqlBulkCopy sqlBulkCopy = new SqlBulkCopy((SqlConnection)DbConnection, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.UseInternalTransaction, null);
            sqlBulkCopy.DestinationTableName = destinationTableName;

            if (DbConnection.State != ConnectionState.Open) DbConnection.Open();
            sqlBulkCopy.WriteToServer(dtData);
            DbConnection.Close();
        }

        /// <summary>
        /// 將資料庫參數集合轉換成 SQL Server 專屬參數集合
        /// </summary>
        /// <param name="parameterCollection"></param>
        /// <returns></returns>
        private List<SqlParameter> ConvertParameterList(List<DatabaseParameter> parameterCollection)
        {
            List<SqlParameter> sqlParameterList = new List<SqlParameter>();

            foreach (DbParameter parameter in parameterCollection)
            {
                SqlParameter sqlParameter = new SqlParameter(parameter.ParameterName, ConvertDbType(parameter.DbType));
                sqlParameter.Direction = parameter.Direction;
                sqlParameter.Value = (parameter.Value == null) ? Convert.DBNull : parameter.Value;
                sqlParameterList.Add(sqlParameter);
            }

            return sqlParameterList;
        }

        /// <summary>
        /// 參數型別轉換
        /// </summary>
        /// <param name="dbType"></param>
        /// <returns></returns>
        private SqlDbType ConvertDbType(DbType dbType)
        {
            SqlDbType sqlDbType = new SqlDbType();

            switch (dbType)
            {
                case DbType.String:
                    sqlDbType = SqlDbType.NVarChar;
                    break;
                case DbType.AnsiString:
                    sqlDbType = SqlDbType.VarChar;
                    break;
                case DbType.Decimal:
                    sqlDbType = SqlDbType.Decimal;
                    break;
                case DbType.Int16:
                    sqlDbType = SqlDbType.SmallInt;
                    break;
                case DbType.Int32:
                    sqlDbType = SqlDbType.Int;
                    break;
                case DbType.Int64:
                    sqlDbType = SqlDbType.BigInt;
                    break;
                case DbType.Single:
                    sqlDbType = SqlDbType.Real;
                    break;
                case DbType.Double:
                    sqlDbType = SqlDbType.Float;
                    break;
                case DbType.DateTime:
                    sqlDbType = SqlDbType.DateTime;
                    break;
                case DbType.Time:
                    sqlDbType = SqlDbType.Timestamp;
                    break;
                default:
                    throw new Exception("Not Supported DbType !!");
            }

            return sqlDbType;
        }

        /// <summary>
        /// 將 SQL 格式化成 SQL Server 專屬的格式
        /// </summary>
        /// <param name="sqlString"></param>
        /// <returns></returns>
        private string FormatSqlString(string sqlString)
        {
            return sqlString.Replace(':', '@').Replace("||","+");
        }

        /// <summary>
        /// 通过DataEntitys增加数据
        /// </summary>
        /// <param name="obj">需要增加的数据实体</param>
        /// <param name="cmd">执行语句的Command,使用事务的时候使用</param>
        /// <returns>返回已增加数量</returns>
        public override int AddData(DataTable obj)
        {
            int n = 0;
            string sSQL = "";
            foreach (DataRow dr in obj.Rows)
            {
                List<SqlParameter> cmdParams = new List<SqlParameter>();
                sSQL = GetAddSQL(obj, ref cmdParams, dr);
                //if (blUpper) sSQL = sSQL.ToUpper();
                if (sSQL.Length != 0)
                {
                    //n += _db.ExecuteNonQuery(cmd, CommandType.Text, sSQL, cmdParams);
                    n += ExecuteNonQuery(sSQL, cmdParams);
                }
            }
            return n;
        }

        private string GetAddSQL(DataTable obj, ref List<SqlParameter> cmdParams, DataRow dr)
        {
            string sSQL, con1 = "", con2 = "", iKey = "ID";
            //string[] field = obj.Columns;
            sSQL = "insert into dbo." + obj.TableName + "(";
            for (int i = 0; i < obj.Columns.Count; i++)
            {
                DataColumn dc = obj.Columns[i];

                //if (dc.ColumnName != iKey)
                //{
                if (!Convert.IsDBNull(dr[dc]))
                {
                    con1 += dc.ColumnName + ",";
                    con2 += "@" + dc.ColumnName + ",";
                    if (obj.Columns[dc.ColumnName].DataType == typeof(System.String))
                    {
                        SqlParameter para = new SqlParameter("@" + dc.ColumnName, SqlDbType.NVarChar);
                        string val = dr[dc].ToString().ToUpper();//若是字串就全轉大寫
                        //val = System.Text.RegularExpressions.Regex.Replace(val, @"\s+", " ").Trim();//Replace space, tab, new line
                        //val = ChineseConverter.Convert(val, ChineseConversionDirection.SimplifiedToTraditional);
                        para.Value = val;
                        cmdParams.Add(para);
                        //cmdParams.Add(new SqlParameter("@" + dc.ColumnName, dr[dc].ToString().ToUpper()));//若是字串就全轉大寫
                        //cmdParams[i].SqlDbType = SqlDbType.NVarChar;
                    }
                    else
                        cmdParams.Add(new SqlParameter("@" + dc.ColumnName, dr[dc.ColumnName]));

                    if (obj.Columns[dc.ColumnName].DataType == typeof(System.Byte[]))
                        cmdParams[i].SqlDbType = SqlDbType.Image;
                }
                //}
                //else
                //{
                //    Int64 intId = GetID("dbo." + obj.TableName + "_SEQ");
                //    con1 += dc.ColumnName + ",";
                //    con2 += "@" + dc.ColumnName + ",";
                //    dr[dc.ColumnName] = intId;
                //    cmdParams.Add(new SqlParameter("@" + dc.ColumnName, dr[dc.ColumnName]));
                //}
            }
            con1 = con1.Substring(0, con1.Length - 1) + ")";
            con2 = " values(" + con2.Substring(0, con2.Length - 1) + ")";
            return sSQL + con1 + con2;
        }

        public override string GetID(string SeqName)
        {
            List<SqlParameter> cmdParams =new List<SqlParameter>();
            string[] seq = SeqName.Split('.');
            string SqlText = "exec dbo.Sequences '" + seq[0] + "','"+ seq[1] + "'";
            return ExecuteScalar(SqlText, cmdParams).ToString();
        }

        /// <summary>
        /// 新增時使用
        /// </summary>
        /// <param name="tblName"></param>
        /// <returns></returns>
        public override DataTable GetDtStructure(string tblName)
        {
            DataTable dt = new DataTable();

            List<DatabaseParameter> lstParameter = new List<DatabaseParameter>();
            string sSQL = "select * from " + tblName + " where 1=2";
            dt = ExecuteQuery(sSQL, lstParameter);
            dt.TableName = tblName;

            return dt;
        }

        /// <summary>
        /// 更新時使用
        /// </summary>
        /// <param name="tblName"></param>
        /// <param name="ID"></param>
        /// <returns></returns>
        public override DataTable GetDtByID(string tblName, string ID)
        {
            DataTable dt = new DataTable();

            List<DatabaseParameter> lstParameter = new List<DatabaseParameter>();
            string sSQL = "select * from " + tblName + " where ID=" + ID;
            dt = ExecuteQuery(sSQL, lstParameter);
            dt.TableName = tblName;

            return dt;
        }

        public override int DeleteData(DataTable obj)
        {
            int n = 0;
            string sSQL = "";
            List<SqlParameter> cmdParams = new List<SqlParameter>();
            foreach (DataRow dr in obj.Rows)
            {
                string id = dr["ID"].ToString();
                if (id.Length == 0) continue;

                sSQL = "delete " + obj.TableName + " where ID=" + id;
                if (sSQL.Length != 0)
                    n += ExecuteNonQuery(sSQL, cmdParams);
            }
            return n;
        }

        /// <summary>
        /// 通过DataEntitys修改数据
        /// </summary>
        /// <param name="obj">需要修改的数据实体</param>
        /// <param name="cmd">执行语句的Command,使用事务的时候使用</param>
        /// <returns>返回已修改数量</returns>
        public override int UpdateData(DataTable obj)
        {
            int n = 0;
            string sSQL = "";
            foreach (DataRow dr in obj.Rows)
            {
                List<SqlParameter> cmdParams = new List<SqlParameter>();
                sSQL = GetUpdateSQL(obj, ref cmdParams, dr);
                //if (blUpper) sSQL = sSQL.ToUpper();
                if (sSQL.Length != 0)
                    n += ExecuteNonQuery(sSQL, cmdParams);
            }
            return n;
        }

        private string GetUpdateSQL(DataTable obj, ref List<SqlParameter> cmdParams, DataRow dr)
        {
            string sSQL, con1 = "", con2 = "", iKey = "ID";
            sSQL = "update dbo." + obj.TableName + " set ";
            for (int i = 0; i < obj.Columns.Count; i++)
            {
                DataColumn dc = obj.Columns[i];
                Type ColumnsType = obj.Columns[dc.ColumnName].DataType;
                con1 += dc.ColumnName + "=@" + dc.ColumnName + ",";

                if (iKey != dc.ColumnName)
                {
                    if (!Convert.IsDBNull(dr[dc]))
                    {
                        if (CheckUpdateDBNull(ColumnsType, dc.ColumnName, dr))
                            cmdParams.Add(new SqlParameter("@" + dc.ColumnName, DBNull.Value));
                        else if (ColumnsType == typeof(System.String))
                        {
                            string val = dr[dc].ToString().ToUpper();//若是字串就全轉大寫
                            //val = System.Text.RegularExpressions.Regex.Replace(val, @"\s+", " ").Trim();//Replace space, tab, new line
                            //val = ChineseConverter.Convert(val, ChineseConversionDirection.SimplifiedToTraditional);
                            cmdParams.Add(new SqlParameter("@" + dc.ColumnName, val));
                        }
                        //else if (ColumnsType == typeof(System.Boolean))
                        //{
                        //    if (dr[dc].ToString() == "True") cmdParams.Add(new SqlParameter("@" + dc.ColumnName, "1"));
                        //    else cmdParams.Add(new SqlParameter("@" + dc.ColumnName, "0"));
                        //}
                        else
                            cmdParams.Add(new SqlParameter("@" + dc.ColumnName, dr[dc]));

                        if (ColumnsType == typeof(System.Byte[]))
                            cmdParams[cmdParams.Count - 1].SqlDbType = SqlDbType.Image;
                    }
                    else
                    {
                        //if (CheckUpdateDBNull(ColumnsType, dc.ColumnName, dr))
                        //    cmdParams.Add(new SqlParameter("@" + dc.ColumnName, DBNull.Value));
                        cmdParams.Add(new SqlParameter("@" + dc.ColumnName, DBNull.Value));
                    }
                }
                else
                {
                    con2 += " where " + dc + "=@" + dc;
                }
            }
            cmdParams.Add(new SqlParameter("@" +iKey,dr[iKey]));
            //cmdParams[obj.Columns.Count - 1] = new SqlParameter("@" + iKey, dr[iKey]);
            //cmdParams[field.Length - 1] = new SqlParameter(iKey,dr[iKey]);
            con1 = con1.Substring(0, con1.Length - 1);
            return sSQL + con1 + con2;
        }
        
        private bool CheckUpdateDBNull(Type obj, string field, DataRow dr)
        {
            if (obj == typeof(string) && dr[field].ToString() == "CHR(0)")
                return true;
            else if (obj == typeof(Int32) && Convert.ToInt32(dr[field]) == -1987654321)
                return true;
            else if (obj == typeof(DateTime) && Convert.ToDateTime(dr[field]).ToString("yyyy/MM/dd") == Convert.ToDateTime("2099/01/01").ToString("yyyy/MM/dd"))
                return true;
            else if (dr[field].ToString() == "9876543210")
                return true;

            return false;
        }
    }
}
