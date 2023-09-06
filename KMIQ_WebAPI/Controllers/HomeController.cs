using loCommon;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace KMIQ.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Result(string id)
        {
            ViewBag.Token = "";
            if (id != null)
            {
                if (getReportData(id))
                {
                    ViewBag.Token = id.ToString();
                    ViewBag.RESULT_DT = this.RESULT_DT;
                    ViewBag.RESULT_DETAIL_MAIN = this.RESULT_DETAIL_MAIN;
                    ViewBag.RESULT_DETAIL_SUB = this.RESULT_DETAIL_SUB;
                    ViewBag.USER_INFO = this.USER_INFO;

                    return View(REPORT_TYPE_VIEW);
                }
                return Content("이 토큰으로 유효한 리포트 결과를 확인하지 못했습니다.");
            }

            return Content("리포트 요청을 위한 토큰이 필요합니다.");

        }

        public ActionResult ResultTest()
        {
            return View();
        }



        #region make report process

        DataTable RESULT_DT = new DataTable();
        DataSet   RESULT_DETAIL_MAIN = new DataSet();
        DataSet   RESULT_DETAIL_SUB = new DataSet();
        Dictionary<string, string> USER_INFO = new Dictionary<string, string>();
        string REPORT_TYPE_VIEW = "Result";

        private bool getReportData(string token)
        {
            bool rtn = false;

            using (SqlConn sql = new SqlConn())
            {
                string qry = "";

                string RESULT_MAIN = "";
                string RESULT_SUB = "";
                string TYPE_GRADE = "";
                string LEVEL_CODE = "";
                string SEX_CODE = "1";

                RESULT_DT = new DataTable();
                RESULT_DETAIL_MAIN = new DataSet();
                RESULT_DETAIL_SUB = new DataSet();
                USER_INFO = new Dictionary<string, string>();

                /// personal report  -- 일단 심화형, 일반형 할것 없이 이 대상자의 리절트를 받아둔다.
                try
                {
                    qry = string.Format(" EXEC [dbo].[USP_RPT_ONLINE_PERSONAL] '{0}'", token);

                    DataTable rtnTb = sql.SqlSelect(qry);

                    if (rtnTb.Rows.Count > 0)
                    {
                        RESULT_DT = rtnTb.Copy();

                        LEVEL_CODE = !DBNull.Value.Equals(RESULT_DT.Rows[0]["LEVEL"]) ? RESULT_DT.Rows[0]["LEVEL"].ToString() : "";
                        TYPE_GRADE = !DBNull.Value.Equals(RESULT_DT.Rows[0]["TYPE_GRADE"]) ? RESULT_DT.Rows[0]["TYPE_GRADE"].ToString() : "";

                        if (LEVEL_CODE != "2") REPORT_TYPE_VIEW = "ResultBasic";

                        try
                        {
                            string CODE_STR = string.Join(",", rtnTb.AsEnumerable().OrderBy(x => x["CODE_RANK"]).Select(x => string.Format("'{0}'", x["CODE"].ToString())).ToArray());
                            string SCORE_STR = string.Join(",", rtnTb.AsEnumerable().OrderBy(x => x["CODE_RANK"]).Select(x => x["SCORE"].ToString()).ToArray());
                            string CRANK_STR = string.Join(",", rtnTb.AsEnumerable().OrderBy(x => x["CODE_RANK"]).Select(x => DBNull.Value.Equals(x["CHOICE_RANK"]) ? "NULL" : x["CHOICE_RANK"].ToString()).ToArray());

                            qry = string.Format(" SELECT [dbo].[FN_GET_RESULT_MAIN] ({0}, {1}); SELECT [dbo].[FN_GET_RESULT_SUB] ({0}, {2}, {1})", CODE_STR, CRANK_STR, SCORE_STR);

                            DataSet ds = loFunctions.getDataSetFromDB(qry);
                            if (ds != null && ds.Tables.Count == 2)
                            {
                                RESULT_MAIN = DBNull.Value.Equals(ds.Tables[0].Rows[0][0]) ? "" : ds.Tables[0].Rows[0][0].ToString();
                                RESULT_SUB = DBNull.Value.Equals(ds.Tables[1].Rows[0][0]) ? "" : ds.Tables[1].Rows[0][0].ToString();
                            }


                            if (RESULT_MAIN != "" && LEVEL_CODE != "" && SEX_CODE != "")
                            {
                                qry = string.Format(" EXEC [dbo].[USP_RPT_PERSONAL_RESULT] '{0}', '{1}', '{2}' ", RESULT_MAIN, LEVEL_CODE, SEX_CODE);
                                RESULT_DETAIL_MAIN = loFunctions.getDataSetFromDB(qry);
                            }

                            if (RESULT_SUB != "" && LEVEL_CODE != "" && SEX_CODE != "")
                            {
                                qry = string.Format(" EXEC [dbo].[USP_RPT_PERSONAL_RESULT] '{0}', '{1}', '{2}' ", RESULT_SUB, LEVEL_CODE, SEX_CODE);
                                RESULT_DETAIL_SUB = loFunctions.getDataSetFromDB(qry);
                            }

                            using (WebApi webapi = new WebApi())
                            {
                                USER_INFO = webapi.getUserInfo(token);
                                if (USER_INFO.ContainsKey("sex"))
                                {
                                    USER_INFO["sex"] = USER_INFO["sex"] == "M" ? "남" : USER_INFO["sex"] == "F" ? "여" : "";
                                }
                            }


                            StringBuilder str = new StringBuilder();
                            foreach (DataRow r in RESULT_DT.Rows)
                            {
                                str.AppendLine(string.Format("{{ axis: '{0}', value: {1:N2} }},", r["CODE"], r["SCORE"]));
                            }

                            ViewBag.GraphData = str.ToString();


                            rtn = true;
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                        rtn = false;
                }
                catch (Exception ex)
                {
                    throw new Exception("mkReport_PersonalReport - USP_RPT_PERSONAL 오류 : " + ex.Message, ex);
                }
            }

            return rtn;
        }

        #endregion

    }
}
