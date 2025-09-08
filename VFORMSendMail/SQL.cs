using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace VFORMSendMail
{
    public class SelectSqlParamTable
    {
        public String ParameterName;
        public SqlDbType SqlDbType;
        public String Value;
    }

    public static class SqlParamTable
    {
        public static SelectSqlParamTable SetSelectSqlParamTable(String ParameterName, SqlDbType SqlDbType, String Value)
        {
            SelectSqlParamTable SetParamTable = new SelectSqlParamTable();
            SetParamTable.ParameterName = ParameterName;
            SetParamTable.SqlDbType = SqlDbType;
            SetParamTable.Value = Value;
            return SetParamTable;
        }
    }

    internal class CommonSQL
    {
        /// <summary>
        /// Select実行
        /// </summary>
        /// <param name="Sql">SQLクエリ</param>
        /// <param name="Param">SQLパラメータ</param>
        /// <returns></returns>
        public static DataTable SelectSql(String Sql, List<SelectSqlParamTable> Param = null)
        {
            var ConnectionString = ConfigurationManager.ConnectionStrings["sqlsvr"].ConnectionString;

            DataTable GetDataTable = new DataTable();

            try
            {
                using (SqlConnection Connection = new SqlConnection(ConnectionString))
                {
                    Connection.Open();
                    SqlCommand Command = new SqlCommand(Sql, Connection);

                    if (Param != null)
                    {
                        for (int i = 0; i < Param.Count; i++)
                        {
                            SqlParameter param = Command.CreateParameter();
                            param.ParameterName = Param[i].ParameterName;
                            param.SqlDbType = Param[i].SqlDbType;
                            param.Direction = ParameterDirection.Input;
                            param.Value = Param[i].Value;
                            Command.Parameters.Add(param);
                        }
                    }

                    var reader = Command.ExecuteReader();
                    GetDataTable.Load(reader);

                    Connection.Close();
                }
            }
            catch (SqlException ex)
            {
                WriteSqlErrorLog(ex, Param);
                throw;
            }

            return GetDataTable;
        }

        private static void WriteSqlErrorLog(SqlException ex, List<SelectSqlParamTable> Param)
        {
            StringBuilder objBld = new StringBuilder();
            objBld.Append(DateTime.Now.ToString());
            objBld.Append("\t");
            objBld.Append(ex.State);
            objBld.Append("\t");
            objBld.AppendLine(ex.Message);
            if (Param != null)
            {
                objBld.AppendLine("パラメータ");
                for (int i = 0; i <= Param.Count - 1; i++)
                {
                    objBld.Append(Param[i].ParameterName + "\t");
                    objBld.Append(Param[i].SqlDbType + "\t");
                    objBld.AppendLine(Param[i].Value + "\t");
                }
            }
            CommonLog.Instance.WriteLog(objBld.ToString());
        }
    }

    public static class SelectSQL
    {
        /// <summary>
        /// 申請の送付期限が現在日付の3日前ユーザーを取得する
        /// </summary>
        public static DataTable GetRequestDateLimitUser()
        {
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("  SELECT");
            sql.AppendLine("    tbl_user.user_name");
            sql.AppendLine("    ,tbl_user.mail");
            sql.AppendLine("    ,tbl_request.request_date");
            sql.AppendLine("    ,ISNULL(tbl_target_name.target, '') as target");
            sql.AppendLine("    ,ISNULL(tbl_camera_name.camera, '') as camera");
            sql.AppendLine("  FROM");
            sql.AppendLine("    tbl_request");
            sql.AppendLine("  INNER JOIN");
            sql.AppendLine("    tbl_user");
            sql.AppendLine("    ON tbl_request.user_id = tbl_user.user_id");
            sql.AppendLine("  LEFT JOIN");
            sql.AppendLine("    (");
            sql.AppendLine("      SELECT");
            sql.AppendLine("        STRING_AGG(CONVERT (NVARCHAR (MAX), CONCAT(TRIM(ISNULL(target_group_name,'')),CONVERT(NVARCHAR, target_no))), ',') WITHIN GROUP (ORDER BY request_id ASC) AS target");
            sql.AppendLine("        ,request_id");
            sql.AppendLine("      FROM ");
            sql.AppendLine("        tbl_request_target");
            sql.AppendLine("      GROUP BY");
            sql.AppendLine("        request_id");
            sql.AppendLine("    ) as tbl_target_name");
            sql.AppendLine("    ON tbl_request.request_id = tbl_target_name.request_id");
            sql.AppendLine("  LEFT JOIN");
            sql.AppendLine("    (");
            sql.AppendLine("      SELECT");
            sql.AppendLine("        STRING_AGG(CONVERT (NVARCHAR (MAX), camera_no), ',') WITHIN GROUP (ORDER BY request_id ASC) AS camera");
            sql.AppendLine("        ,request_id");
            sql.AppendLine("      FROM ");
            sql.AppendLine("        tbl_request_camera");
            sql.AppendLine("      GROUP BY");
            sql.AppendLine("        request_id");
            sql.AppendLine("    ) as tbl_camera_name");
            sql.AppendLine("    ON tbl_request.request_id = tbl_camera_name.request_id");
            sql.AppendLine("  WHERE");
            sql.AppendLine("    request_date = DATEADD(DAY, 3, CAST(GETDATE() AS DATE))");
            sql.AppendLine("    AND tbl_request.status = 0");

            DataTable GetSQLDataTable = CommonSQL.SelectSql(sql.ToString());

            return GetSQLDataTable;
        }

        /// <summary>
        /// 申請期限が1か月と2カ月前でターゲット交換数orカメラ構成数があるユーザーを取得する
        /// </summary>
        public static DataTable GetContentLimitUser()
        {
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("  SELECT ");
            sql.AppendLine("    contract_id");
            sql.AppendLine("    ,tbl_contract.branch_id");
            sql.AppendLine("    ,tbl_user.mail");
            sql.AppendLine("    ,contract_type_name");
            sql.AppendLine("    ,contract_count");
            sql.AppendLine("    ,request_limit_date");
            sql.AppendLine("    ,contract_type_name");
            sql.AppendLine("    ,(max_target_change * contract_count) - target_change_count as target_count");
            sql.AppendLine("    ,(max_camera_change * contract_count) - camera_change_count as camera_count");
            sql.AppendLine("  FROM ");
            sql.AppendLine("    tbl_contract");
            sql.AppendLine("  INNER JOIN");
            sql.AppendLine("  (");
            sql.AppendLine("     SELECT");
            sql.AppendLine("      STRING_AGG(CONVERT (NVARCHAR (MAX), mail), ',') WITHIN GROUP (ORDER BY branch_id ASC) AS mail");
            sql.AppendLine("      ,branch_id");
            sql.AppendLine("     FROM");
            sql.AppendLine("      tbl_user");
            sql.AppendLine("     GROUP BY");
            sql.AppendLine("      branch_id");
            sql.AppendLine("  ) as tbl_user");
            sql.AppendLine("  ON tbl_contract.branch_id = tbl_user.branch_id");
            sql.AppendLine("  INNER JOIN");
            sql.AppendLine("    mst_contract_type");
            sql.AppendLine("    ON tbl_contract.contract_pattern_id = mst_contract_type.contract_type_id");
            sql.AppendLine("  WHERE ");
            sql.AppendLine("    (request_limit_date = DATEADD(MONTH, 1, CAST(GETDATE() AS DATE))");
            sql.AppendLine("      Or request_limit_date = DATEADD(MONTH, 2, CAST(GETDATE() AS DATE)))");
            sql.AppendLine("    AND ((max_target_change * contract_count) - target_change_count > 0 ");
            sql.AppendLine("      Or (max_camera_change * contract_count) - camera_change_count > 0)");

            DataTable GetSQLDataTable = CommonSQL.SelectSql(sql.ToString());

            return GetSQLDataTable;
        }
    }
}
