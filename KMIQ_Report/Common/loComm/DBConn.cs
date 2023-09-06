using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
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
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        RtnTB = new DataTable();
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

        public DataTable SqlSelectOfCmd(SqlCommand cmd, string customConnStr = null)
        {
            DataTable RtnTB = new DataTable();
            string conStr = ConnStr;
            if (customConnStr != null) conStr = customConnStr;

            using (SqlConnection SqlCon = new SqlConnection(conStr))
            {
                try
                {
                    SqlCon.Open();

                    if (SqlCon.State == ConnectionState.Open)
                    {
                        cmd.Connection = SqlCon;
                        using (SqlDataAdapter SqlAdpt = new SqlDataAdapter(cmd))
                        {
                            try
                            {
                                SqlAdpt.Fill(RtnTB);
                            }
                            catch
                            {
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    RtnTB = new DataTable();
                }
            }

            return RtnTB;
        }

        public bool SqlExecuteOfCmd(SqlCommand cmd, string customConnStr = null)
        {
            bool Rtn = false;

            string conStr = ConnStr;
            if (customConnStr != null) conStr = customConnStr;

            using (SqlConnection SqlCon = new SqlConnection(conStr))
            {
                try
                {
                    SqlCon.Open();

                    if (SqlCon.State == ConnectionState.Open)
                    {
                        cmd.Connection = SqlCon;
                        try
                        {
                            cmd.ExecuteNonQuery();
                            Rtn = true;
                        }
                        catch
                        {
                            Rtn = false;
                        }
                        finally
                        {
                            cmd.Dispose();
                        }
                    }
                }
                catch (Exception)
                {
                    Rtn = false;
                }
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




    /*
    public class SqliteConn : IDisposable
    {
        private string _DbFile = "";
        public SqliteConn(string dbFile = "")
        {
            if (dbFile == "") dbFile = KMIQ.Globals.ThisWorkbook.Main.G_DB_FILE;
            _DbFile = dbFile;
        }

        public DataTable SqlSelect(string sql, string customDbFile = null)
        {
            DataTable RtnTB = new DataTable();
            string conStr = _DbFile;
            if (customDbFile != null) conStr = customDbFile;

            if (!string.IsNullOrEmpty(sql))
            {
                if (!sql.EndsWith(";")) sql += ";";

                using (SQLiteConnection SqlCon = new SQLiteConnection("Data source=" + conStr))
                {
                    try
                    {
                        SqlCon.Open();

                        if (SqlCon.State == ConnectionState.Open)
                        {
                            using (SQLiteCommand SqlCmd = new SQLiteCommand(sql, SqlCon))
                            {
                                using (SQLiteDataAdapter SqlAdpt = new SQLiteDataAdapter(SqlCmd))
                                {
                                    try
                                    {
                                        SqlAdpt.Fill(RtnTB);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("SqlSelect 오류 : " + ex.Message);
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
                    catch (Exception ex)
                    {
                        Console.WriteLine("[SqlSelect] on Error : " + ex.Message);
                    }
                }

                return RtnTB;
            }
            else
            {
                return RtnTB;
            }
        }


        public bool SqlExcute(string sql, string customDbFile = null)
        {
            bool Rtn = false;
            string conStr = _DbFile;
            if (customDbFile != null) conStr = customDbFile;

            if (!string.IsNullOrEmpty(sql))
            {
                if (!sql.EndsWith(";")) sql += ";";

                using (SQLiteConnection SqlCon = new SQLiteConnection("Data source=" + conStr))
                {
                    try
                    {
                        SqlCon.Open();

                        if (SqlCon.State == ConnectionState.Open)
                        {
                            using (SQLiteCommand SqlCmd = new SQLiteCommand(sql, SqlCon))
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






        public bool SqlBulkInsert(DataTable dtOutData, String TableNm, string customDbFile = null)
        {
            bool Rtn = false;

            string conStr = _DbFile;
            if (customDbFile != null) conStr = customDbFile;

            StringBuilder InsertCmdFields = new StringBuilder();
            StringBuilder InsertCommand = default(StringBuilder);


            if (dtOutData.Rows.Count > 0)
            {
                //***** Insert Command의 공통으로 쓰이는  테이블/필드 지정 문자열 생성


                InsertCmdFields.Append(" INSERT INTO " + TableNm + " (");
                foreach (DataColumn colNM in dtOutData.Columns)
                {
                    InsertCmdFields.Append(" [" + colNM.ColumnName + "],");
                }
                InsertCmdFields.Remove(InsertCmdFields.Length - 1, 1);
                InsertCmdFields.Append(" ) VALUES (");


                using (SQLiteConnection SqlCon = new SQLiteConnection("Data source=" + conStr))
                {
                    SqlCon.Open();

                    using (SQLiteCommand sqliteCmd = new SQLiteCommand(SqlCon))
                    {
                        using (SQLiteTransaction Trans = SqlCon.BeginTransaction())
                        {

                            try
                            {
                                foreach (DataRow dtRow in dtOutData.Rows)
                                {
                                    InsertCommand = new StringBuilder();

                                    InsertCommand.Append(InsertCmdFields.ToString());
                                    for (int colCnt = 0 ; colCnt <= dtOutData.Columns.Count - 1 ; colCnt++)
                                    {
                                        if (dtOutData.Columns[colCnt].DataType == Type.GetType("System.String"))
                                        {
                                            InsertCommand.Append("'" + dtRow[colCnt].ToString() + "',");
                                        }
                                        else
                                        {
                                            if (string.IsNullOrEmpty(dtRow[colCnt].ToString()))
                                            {
                                                InsertCommand.Append("null,");
                                            }
                                            else
                                            {
                                                InsertCommand.Append(dtRow[colCnt].ToString() + ",");
                                            }

                                        }

                                    }
                                    InsertCommand.Remove(InsertCommand.Length - 1, 1);
                                    //InsertCommand.Insert(0, InsertCmdFields.ToString)
                                    InsertCommand.Append(" );");

                                    sqliteCmd.CommandText = InsertCommand.ToString();
                                    sqliteCmd.ExecuteNonQuery();
                                }

                                Trans.Commit();
                                Rtn = true;
                            }
                            catch (Exception ex)
                            {
                                Trans.Rollback();
                                Rtn = false;
                            }
                            finally
                            {
                                Trans.Dispose();
                                sqliteCmd.Dispose();
                                SqlCon.Close();
                                SqlCon.Dispose();
                            }
                        }
                    }
                }

                return Rtn;
            }
            else
            {
                return false;
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

        public void Dispose()
        {
            _DbFile = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    */
}



