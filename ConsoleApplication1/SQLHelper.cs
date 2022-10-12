using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;

/// <summary>
/// SQLHelper 的摘要描述
/// </summary>
public class SQLHelper
{
    #region 公用屬性
    /// <summary>
    /// 取得或設定用來在資料來源連結字串。
    /// </summary>
    public String _Conn;
    /// <summary>
    /// 取得或設定用來在資料來源中選取資料錄的 SQL 陳述式或預存程序。
    /// </summary>
    public String _Cmd;
    /// <summary>
    /// 指定如何解譯命令字串。
    /// </summary>
    public CommandType _CmdType;
    /// <summary>
    /// 取得或設定 SQL 陳述式或預存程序 參數。
    /// </summary>
    public SqlParameterCollection _Params;
    /// <summary>
    /// 設定用來在資料來源中選取資料錄的 SQL 陳述式或預存程序集合。
    /// </summary>
    public System.Collections.Generic.List<System.Data.SqlClient.SqlCommand> _Cmds;
    public String _ErrMsg;

    #endregion
    #region 建構子
    public SQLHelper()
    {
        this._Conn = GetWebConfigConnectionString("DB");
        this._Cmd = string.Empty;
        this._CmdType = CommandType.Text;
        this._Params = (SqlParameterCollection)typeof(SqlParameterCollection).GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null, Type.EmptyTypes, null).Invoke(null);
        this._Cmds = new System.Collections.Generic.List<SqlCommand>();
        this._ErrMsg = string.Empty;
    }
    public SQLHelper(String CommandText)
        : this()
    {
        this._Cmd = CommandText;
    }
    #endregion
    #region 私有函式
    private void FillCommandParameters(SqlParameterCollection Parameters)
    {
        foreach (SqlParameter param in this._Params)
            Parameters.Add((SqlParameter)((ICloneable)param).Clone());
    }
    private static string GetWebConfigConnectionString(string name)
    {
        try
        {
            return System.Web.Configuration.WebConfigurationManager.ConnectionStrings[name].ConnectionString;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string Exception(Exception ex)
    {
        return ex.Message;
    }

    #endregion
    #region 公用涵式
    #region GetDataTable
    /// <summary>
    /// 執行 SQL 陳述式，並傳回產生的資料列。
    /// </summary>
    /// <returns></returns>
    public System.Data.DataTable GetDataTable()
    {
        System.Data.DataTable dt = new System.Data.DataTable();
        using (SqlDataAdapter da = new SqlDataAdapter(this._Cmd, this._Conn))
        {
            try
            {
                da.SelectCommand.CommandType = this._CmdType;
                FillCommandParameters(da.SelectCommand.Parameters);
                //da.FillSchema(dt, SchemaType.Source);
                da.Fill(dt);
                da.Dispose();
            }
            catch
            {
                throw;
            }
        }
        this._Params.Clear();
        return dt;
    }
    #endregion


    #region GetDataTable
    /// <summary>
    /// 執行 SQL 陳述式，並傳回產生的資料列。
    /// </summary>
    /// <returns></returns>
    public System.Data.DataTable GetDataTable(SqlConnection dbcn)
    {
        System.Data.DataTable dt = new System.Data.DataTable();
        using (SqlDataAdapter da = new SqlDataAdapter(this._Cmd, dbcn))
        {
            try
            {
                da.SelectCommand.CommandType = this._CmdType;
                FillCommandParameters(da.SelectCommand.Parameters);
                //da.FillSchema(dt, SchemaType.Source);
                da.Fill(dt);
                da.Dispose();
            }
            catch
            {
                throw;
            }
        }
        this._Params.Clear();
        return dt;
    }
    #endregion
    #region ExecuteNonQuery
    /// <summary>
    /// 執行 SQL 陳述式，並傳回受影響的資料列數。
    /// </summary>
    /// <returns></returns>
    public System.Int32 ExecuteNonQuery()
    {
        System.Int32 rtnVal = 0;
        SqlConnection conn = new SqlConnection(this._Conn);
        using (SqlCommand cmd = new SqlCommand(this._Cmd, conn))
        {
            cmd.CommandType = this._CmdType;
            this.FillCommandParameters(cmd.Parameters);
            cmd.Connection.Open();
            rtnVal = cmd.ExecuteNonQuery();
            cmd.Dispose();
        }
        conn.Close();
        return rtnVal;
    }


    public System.Int32 ExecuteNonQueryIDENTITY()
    {
        System.Int32 rtnVal = 0;
        SqlParameter IDParameter = new SqlParameter("@ID",SqlDbType.Int);
        IDParameter.Direction = ParameterDirection.Output;
        SqlConnection conn = new SqlConnection(this._Conn);
        using (SqlCommand cmd = new SqlCommand(this._Cmd, conn))
        {
            cmd.CommandType = this._CmdType;
            this.FillCommandParameters(cmd.Parameters);
            cmd.Parameters.Add(IDParameter);
            cmd.Connection.Open();
            rtnVal = cmd.ExecuteNonQuery();
            cmd.Dispose();
        }
        conn.Close();
        return (int)IDParameter.Value;
    }
    #endregion
    public SqlConnection getDbcn()
    {
        SqlConnection cn = new SqlConnection(this._Conn);
        cn.Open();
        return cn;
    }

    #region ExecuteNonQuery
    /// <summary>
    /// 執行 SQL 陳述式，並傳回受影響的資料列數。
    /// </summary>
    /// <returns></returns>
    public System.Int32 ExecuteNonQuery(SqlConnection dbcn)
    {
        System.Int32 rtnVal = 0;
        using (SqlCommand cmd = new SqlCommand(this._Cmd, dbcn))
        {
            cmd.CommandType = this._CmdType;

            this.FillCommandParameters(cmd.Parameters);
            rtnVal = cmd.ExecuteNonQuery();
            cmd.Dispose();
        }
        return rtnVal;
    }
    #endregion


    public System.Data.DataSet ExecuteDataSet()
    {
        DataSet ds = new DataSet();
        using (SqlConnection cn = getDbcn())
        {
            using (SqlCommand cm = new SqlCommand(this._Cmd, cn))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(cm))
                {
                    da.SelectCommand.CommandType = this._CmdType;
                    foreach (SqlParameter param in this._Params)
                        cm.Parameters.Add((SqlParameter)((ICloneable)param).Clone());
                    //da.FillSchema(dt, SchemaType.Source);
                    da.Fill(ds);
                    da.Dispose();
                }
            }
        }
        return ds;
    }

    #region Transaction
    /// <summary>
    /// 在 SQL Server 資料庫中產生的 Transact-SQL 交易。
    /// </summary>
    public void ExecuteTransaction()
    {
        using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(this._Cmd))
        {
            cmd.CommandType = this._CmdType;
            FillCommandParameters(cmd.Parameters);
            this._Cmds.Add(cmd);
        }
    }

    /// <summary>
    /// 認可資料庫交易。
    /// </summary>
    /// <returns></returns>
    public System.Boolean ExecuteTransCommit()
    {
        System.Boolean bRet = false;
        using (System.Data.SqlClient.SqlConnection conn = new System.Data.SqlClient.SqlConnection(this._Conn))
        {
            conn.Open();
            using (System.Data.SqlClient.SqlTransaction trans = conn.BeginTransaction())
            {
                try
                {
                    foreach (System.Data.SqlClient.SqlCommand cmd in this._Cmds)
                    {
                        //cmd.CommandTimeout = this._CommandTimeout;
                        cmd.Connection = conn;
                        cmd.Transaction = trans;
                        cmd.ExecuteNonQuery();
                    }
                    trans.Commit();
                    trans.Dispose();
                    conn.Close();
                    conn.Dispose();
                    bRet = true;
                }
                catch (Exception ex)
                {
                    this._ErrMsg = Exception(ex);
                    trans.Rollback();
                    trans.Dispose();
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        return bRet;
    }
    #endregion
    #endregion
}
