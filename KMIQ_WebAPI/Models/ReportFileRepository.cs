using loCommon;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Http;

namespace KMIQ.Models
{
    public class ReportFileRepository
    {
        List<ReportFile> _dbContext;

        public ReportFileRepository()
        {

        }

        public ReportFile GetFileByToken(string id)
        {
            ReportFile rtn = new ReportFile();

            if (id != "")
            {
                string qry = string.Format(" EXEC [dbo].[USP_GET_REPORT_FILE] 1, '{0}'", id);

                using (SqlConn sql = new SqlConn())
                {
                    try
                    {
                        DataTable rstDt = sql.SqlSelect(qry);

                        if (rstDt.Rows.Count > 0)
                        {
                            rtn.returnMsg = rstDt.Rows[0]["RESULT"].ToString();

                            if (rtn.returnMsg == "OK")
                            {
                                rtn.isEnabled = true;
                                rtn.fileData = (byte[])rstDt.Rows[0]["FILESTREAM"];
                                rtn.fileName = rstDt.Rows[0]["FILENAME"].ToString();
                                rtn.fileName = rtn.fileName.Substring(rtn.fileName.IndexOf("___") + 3);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rtn.returnMsg = string.Format("System error - " + ex.Message);
                    }
                }
            }

            return rtn;
        }

    }
}