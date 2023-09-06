using System;
using System.Data;
using System.Data.SqlClient;

namespace loCommon
{
    public class SqlConn : IDisposable
    {
        public string ConnStr { get; set; }

        public SqlConn()
        {
            ConnStr = KMIQ.Properties.Settings.Default.ConnectionString;
        }

        public DataTable SqlSelect(string sql, string customConnStr = null)
        {
            DataTable RtnTB = new DataTable();
            string conStr = ConnStr;
            if (customConnStr != null) conStr = customConnStr;

            if (!string.IsNullOrEmpty(sql))
            {
                if (!sql.EndsWith(";")) sql += ";";

                using (SqlConnection SqlCon = new SqlConnection(conStr))
                {
                    try
                    {
                        SqlCon.Open();

                        if (SqlCon.State == ConnectionState.Open)
                        {
                            using (SqlCommand SqlCmd = new SqlCommand(sql, SqlCon))
                            {
                                using (SqlDataAdapter SqlAdpt = new SqlDataAdapter(SqlCmd))
                                {
                                    try
                                    {
                                        SqlAdpt.Fill(RtnTB);
                                    }
                                    catch
                                    {
                                        RtnTB = new DataTable();
                                    }
                                    finally
                                    {
                                        SqlAdpt.Dispose();
                                        SqlCmd.Dispose();
                                        SqlCon.Close();
                                        SqlCon.Dispose();
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        return new DataTable();
                    }
                }

                return RtnTB;
            }
            else
            {
                return null;
            }
        }


        public bool SqlExcute(string sql, string customConnStr = null)
        {
            bool Rtn = false;
            string conStr = ConnStr;
            if (customConnStr != null) conStr = customConnStr;

            if (!string.IsNullOrEmpty(sql))
            {
                if (!sql.EndsWith(";")) sql += ";";

                using (SqlConnection SqlCon = new SqlConnection(conStr))
                {
                    try
                    {
                        SqlCon.Open();

                        if (SqlCon.State == ConnectionState.Open)
                        {
                            using (SqlCommand SqlCmd = new SqlCommand(sql, SqlCon))
                            {

                                try
                                {
                                    SqlCmd.ExecuteNonQuery();
                                    Rtn = true;
                                }
                                catch
                                {
                                    Rtn = false;
                                }
                                finally
                                {
                                    SqlCmd.Dispose();
                                    SqlCon.Close();
                                    SqlCon.Dispose();
                                }

                            }
                        }
                    }
                    catch (Exception)
                    {
                        Rtn = false;
                    }
                }

            }
            else
            {
                Rtn = false;
            }

            return Rtn;
        }


        public DataSet SqlSelectMultiResult(string sql, string customConnStr = null)
        {
            DataSet RtnTBSet = new DataSet();
            string conStr = ConnStr;
            if (customConnStr != null) conStr = customConnStr;

            if (!string.IsNullOrEmpty(sql))
            {
                if (!sql.EndsWith(";")) sql += ";";

                using (SqlConnection SqlCon = new SqlConnection(conStr))
                {
                    try
                    {
                        SqlCon.Open();

                        if (SqlCon.State == ConnectionState.Open)
                        {
                            using (SqlCommand SqlCmd = new SqlCommand(sql, SqlCon))
                            {
                                using (SqlDataAdapter SqlAdpt = new SqlDataAdapter(SqlCmd))
                                {
                                    try
                                    {
                                        SqlAdpt.Fill(RtnTBSet);
                                    }
                                    catch
                                    {
                                        RtnTBSet = new DataSet();
                                    }
                                    finally
                                    {
                                        SqlAdpt.Dispose();
                                        SqlCmd.Dispose();
                                        SqlCon.Close();
                                        SqlCon.Dispose();
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        return new DataSet();
                    }
                }

                return RtnTBSet;
            }
            else
            {
                return null;
            }
        }



        public string in2db(string inStr, string chgStr)
        {
            string returnVal = null;
            returnVal = inStr;

            if ((returnVal == null) | string.IsNullOrEmpty(returnVal))
            {
                returnVal = chgStr;
            }
            else
            {
                returnVal = returnVal.Replace("'", "''");
                //returnVal = returnVal.Replace(Chr(39), "''")
                //str = Replace(str,Chr(34),""") '["]
                //str = Replace(str,"<","<")
                //str = Replace(str,">",">")
                //str = Replace(str,"(","(")
                //str = Replace(str,")",")")

                //str = Replace(str,"#","#")
                //str = Replace(str,"&","&")
            }

            return returnVal;
        }


        public string ConvertDataTableToHTML(DataTable dt)
        {
            if (dt == null) return "";

            string html = "<table border=1>";
            //add header row
            html += "<tr>";
            for (int i = 0 ; i < dt.Columns.Count ; i++)
                html += "<td>" + dt.Columns[i].ColumnName + "</td>";
            html += "</tr>";
            //add rows
            for (int i = 0 ; i < dt.Rows.Count ; i++)
            {
                html += "<tr>";
                for (int j = 0 ; j < dt.Columns.Count ; j++)
                    html += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                html += "</tr>";
            }
            html += "</table>";
            return html;
        }






        #region IDisposable member
        public void Dispose()
        {
            ConnStr = "";
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

    }
}