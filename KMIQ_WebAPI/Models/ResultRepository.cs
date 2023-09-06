using loCommon;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace KMIQ.Models
{
    public class ResultRepository : IRepository<Result>
    {
        public ResultRepository()
        { }

        public string Add(Result entity)
        {
            throw new NotImplementedException();
        }

        public string Delete(Result entity)
        {
            throw new NotImplementedException();
        }

        public Result FindById(int id)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<Result> GetAll()
        {
            throw new NotImplementedException();
        }

        public IEnumerable<Result> SelectById(int id)
        {
            throw new NotImplementedException();
        }

        public string Update(Result entity)
        {
            string rtn = "";

            /// 먼저 entity의 데이터 유효성을 검증한다.
            if (entity.ID == "") return "ID_IS_NOTHING";
            if (entity.TypeId == 0) return "TypeId_IS_NOTHING";
            if (entity.Token == "") return "Token_IS_NOTHING";
            if (entity.uName == "") return "Name_IS_NOTHING";
            if (entity.ResultStr == "") return "ResultStr_IS_NOTHING";

            using (SqlConnection SqlCon = new SqlConnection(Properties.Settings.Default.ConnectionString))
            {
                try
                {
                    SqlCon.Open();

                    if (SqlCon.State == ConnectionState.Open)
                    {
                        using (SqlTransaction transaction = SqlCon.BeginTransaction("DataUploadTransaction"))
                        {
                            using (SqlCommand SqlCmd = SqlCon.CreateCommand())
                            {
                                SqlCmd.Connection = SqlCon;
                                SqlCmd.Transaction = transaction;

                                SqlCmd.CommandType = CommandType.StoredProcedure;
                                SqlCmd.CommandText = "dbo.[USP_UPDATE_ONLINE_DATA]";

                                SqlCmd.Parameters.Add("@TYPEID", SqlDbType.Int);
                                SqlCmd.Parameters.Add("@ID", SqlDbType.NVarChar, 100);
                                SqlCmd.Parameters.Add("@TOKEN", SqlDbType.VarChar, 100);
                                SqlCmd.Parameters.Add("@RESULT", SqlDbType.VarChar, 500);
                                SqlCmd.Parameters.Add("@NAME", SqlDbType.NVarChar, 50);
                                SqlCmd.Parameters.Add("@BIRTH", SqlDbType.VarChar, 10);
                                SqlCmd.Parameters.Add("@EMAIL", SqlDbType.NVarChar, 50);
                                SqlCmd.Parameters.Add("@TEL", SqlDbType.VarChar, 15);
                                SqlCmd.Parameters.Add("@IS_AGREE", SqlDbType.Bit);

                                try
                                {
                                    SqlCmd.Parameters["@TYPEID"].Value = entity.TypeId;
                                    SqlCmd.Parameters["@ID"].Value = entity.ID;
                                    SqlCmd.Parameters["@TOKEN"].Value = entity.Token;
                                    SqlCmd.Parameters["@RESULT"].Value = entity.ResultStr;
                                    SqlCmd.Parameters["@NAME"].Value = entity.uName;
                                    if (entity.uBirth != "") SqlCmd.Parameters["@BIRTH"].Value = entity.uBirth;
                                    if (entity.uEmail != "") SqlCmd.Parameters["@EMAIL"].Value = entity.uEmail;
                                    if (entity.uTel != "") SqlCmd.Parameters["@TEL"].Value = entity.uTel;
                                    SqlCmd.Parameters["@IS_AGREE"].Value = 1;


                                    DataTable rtnTb = new DataTable();
                                    using (SqlDataAdapter SqlAdpt = new SqlDataAdapter(SqlCmd))
                                    {
                                        try
                                        {
                                            SqlAdpt.Fill(rtnTb);
                                        }
                                        catch (Exception ex)
                                        {
                                            throw ex;
                                        }
                                    }

                                    if (rtnTb.Rows.Count > 0)
                                    {
                                        rtn = rtnTb.Rows[0][0].ToString();
                                    }

                                    transaction.Commit();
                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();
                                    rtn = "FAILED : " + ex.Message;
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                }
            }

            return rtn;
        }
    }
}