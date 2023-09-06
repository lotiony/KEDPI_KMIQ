using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace loCommon
{
    public class ExcelFileLoad
    {
        //***************************************************************************************
        // 
        //***** 엑셀파일을 OLEDB로 연결해서 DataTable형태로 리턴해준다.
        // 
        //***************************************************************************************
        public static DataTable ExcelToDataSet(string FileNM, string DataRange, string Msg, short Criteria_Range1 = -1, string Criteria1 = "", short Criteria_Range2 = -1, string Criteria2 = "", string addWhere = "", bool errIgnore = false)
	    {

		    string strProvider = "";

		    StringBuilder qq = new StringBuilder();
		    DataTable Result = new DataTable();


		    if (FileNM.Contains(".xlsx")) {
			    strProvider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileNM + ";Extended Properties=\"Excel 12.0;HDR=YES\"";
		    } else {
			    strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileNM + ";Extended Properties=\"Excel 8.0;HDR=YES\"";
		    }


		    try {
			    using (OleDbConnection oConn = new OleDbConnection(strProvider)) {
				    oConn.Open();

				    //If InStr(DataRange, "$'") > -1 Then
				    //    DataRange = DataRange.Replace(" ", "")
				    //End If

				    qq.Clear();
				    qq.AppendLine(" SELECT * FROM " + DataRange);
				    if (!string.IsNullOrEmpty(addWhere)) {
					    qq.AppendLine(" " + addWhere);
				    }

				    try {
					    OleDbDataAdapter oAdpt = new OleDbDataAdapter(qq.ToString(), oConn);

					    oAdpt.Fill(Result);

					    //***** 비교조건식을 확인해서 맞지 않으면 대충 에러 발생시킨다.
					    if (Criteria_Range1 >= 0)
						    if (Result.Columns[Criteria_Range1 - 1].ToString() != Criteria1)
                                Result = null;

					    if (Criteria_Range2 >= 0)
						    if (Result.Columns[Criteria_Range2 - 1].ToString() != Criteria2)
                                Result = null;


				    } catch {
					    if (!errIgnore) {
						    if (!string.IsNullOrEmpty(Msg)) {
							    MessageBox.Show(Msg);
						    } else {
							    MessageBox.Show("데이터 시트가 올바르지 않습니다.");
						    }
					    }
					    return null;
				    } finally {
					    oConn.Close();

				    }
			    }


		    } catch {
		    }

		    return Result;

	    }




        //***************************************************************************************
        // 
        //***** 엑셀파일을 OLEDB로 연결해서 DataTable형태로 리턴해준다. : 시트번호로 처리해준다.
        // 
        //***************************************************************************************
        public static DataTable ExcelToDataSet(string FileNM, int SheetIdx, string DataRange, string Msg, short Criteria_Range1 = -1, string Criteria1 = "", short Criteria_Range2 = -1, string Criteria2 = "", string addWhere = "", string[] addValH = null,
        string[] addValV = null)
	    {

		    string strProvider = "";
		    StringBuilder qq = new StringBuilder();
		    DataTable Result = new DataTable();


		    if (FileNM.Contains(".xlsx")) {
			    strProvider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileNM + ";Extended Properties=\"Excel 12.0;HDR=YES\"";
		    } else {
			    strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileNM + ";Extended Properties=\"Excel 8.0;HDR=YES\"";
		    }


		    try {
			    using (OleDbConnection oConn = new OleDbConnection(strProvider)) {
				    oConn.Open();

				    DataTable dt = oConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
				    string dTblName = dt.Rows[SheetIdx - 1]["TABLE_NAME"].ToString().Replace("'", "");

				    qq.Clear();
				    //qq.AppendLine(" SELECT * FROM " + DataRange)
				    qq.AppendLine(" SELECT * ");
				    if ((addValH != null)) {
					    for (byte iR = 0; iR <= addValH.Length - 1; iR++) {
						    qq.AppendLine(" ,'" + addValV[iR] + "' AS " + addValH[iR]);
					    }
				    }
				    //qq.AppendLine(" FROM [" + dTblName.Insert(dTblName.Length - 1, IIf(DataRange <> "", DataRange, "")).Replace("1", "1") + "]")
				    qq.AppendLine(" FROM [" + dTblName + (!string.IsNullOrEmpty(DataRange) ? DataRange : "") + "]");
				    if (!string.IsNullOrEmpty(addWhere)) {
					    qq.AppendLine(" " + addWhere);
				    }

				    try {
					    OleDbDataAdapter oAdpt = new OleDbDataAdapter(qq.ToString(), oConn);
					    oAdpt.Fill(Result);

					    //***** 비교조건식을 확인해서 맞지 않으면 대충 에러 발생시킨다.
                        if (Criteria_Range1 >= 0)
                            if (Result.Columns[Criteria_Range1 - 1].ToString() != Criteria1)
                                Result = null;

					    if (Criteria_Range2 >= 0)
						    if (Result.Columns[Criteria_Range2 - 1].ToString() != Criteria2)
                                Result = null;


				    } catch {
					    MessageBox.Show("데이터 시트가 올바르지 않습니다.");
					    return null;
				    } finally {
					    oConn.Close();

				    }
			    }


		    } catch {
		    }

		    return Result;

	    }


        //***************************************************************************************
        // 
        //***** 엑셀파일을 OLEDB로 연결해서 DataTable형태로 리턴해준다. (커스텀 쿼리 이용)
        // 
        //***************************************************************************************
        public static DataTable ExcelToDataSet(string FileNM, string qry, string Msg, bool errIgnore = false)
        {

            string strProvider = "";

            StringBuilder qq = new StringBuilder();
            DataTable Result = new DataTable();


            if (FileNM.Contains(".xlsx"))
            {
                strProvider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileNM + ";Extended Properties=\"Excel 12.0;HDR=YES\"";
            }
            else
            {
                strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileNM + ";Extended Properties=\"Excel 8.0;HDR=YES\"";
            }

            try
            {
                using (OleDbConnection oConn = new OleDbConnection(strProvider))
                {
                    oConn.Open();

                    try
                    {
                        OleDbDataAdapter oAdpt = new OleDbDataAdapter(qry, oConn);

                        oAdpt.Fill(Result);
                    }
                    catch (Exception ex)
                    {
                        if (!errIgnore)
                        {
                            if (!string.IsNullOrEmpty(Msg))
                            {
                                MessageBox.Show(Msg + Environment.NewLine + ex.Message);
                            }
                            else
                            {
                                MessageBox.Show("데이터 추출 쿼리가 올바르지 않습니다." + Environment.NewLine + ex.Message);
                            }
                        }
                        return null;
                    }
                    finally
                    {
                        oConn.Close();
                    }
                }
            }
            catch
            {
            }

            return Result;

        }

    }

}
