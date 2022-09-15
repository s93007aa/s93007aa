using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;
using System.Data;

namespace RSI.Data
{
    /// <summary>
    /// 資料庫命令參數
    /// </summary>
    public class DatabaseParameter: DbParameter
    {
        /// <summary>
        /// 資料庫命令參數建構式
        /// </summary>
        public DatabaseParameter(string parameterName, DbType dbType, object value)
        {
            ParameterName = parameterName;
            DbType = dbType;
            Direction = ParameterDirection.Input;
            Value = value;
        }
        
        /// <summary>
        /// 資料庫命令參數建構式
        /// </summary>
        public DatabaseParameter(string parameterName, DbType dbType, object value, ParameterDirection parameterDirection)
        {
            ParameterName = parameterName;
            DbType = dbType;
            Direction = parameterDirection;
            Value = value;
        }
        
        public override DbType DbType { get; set; }

        public override ParameterDirection Direction { get; set; }

        public override bool IsNullable { get; set; }

        public override string ParameterName { get; set; }

        public override int Size { get; set; }

        public override string SourceColumn { get; set; }

        public override bool SourceColumnNullMapping { get; set; }

        public override DataRowVersion SourceVersion { get; set; }

        public override object Value { get; set; }

        public override void ResetDbType()
        {
        }
    }
}
