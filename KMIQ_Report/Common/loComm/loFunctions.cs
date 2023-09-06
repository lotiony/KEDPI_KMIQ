using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ADODB;
using System.ComponentModel;
//using Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Net;
using Microsoft.VisualBasic.Logging;
using System.IO;
using C1.C1Excel;

namespace loCommon
{
    public static class loFunctions
    {
        [DllImport("ole32.dll")]
        public static extern void CoUninitialize();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        ///******************************************************************************************************************************************************
        /// 오브젝트 리소스 해제
        ///******************************************************************************************************************************************************
        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
        }

        public static void ProcessKill(object Obj)
        {
            try
            {
                uint pID = 0;
                if (Obj != null)
                {
                    GetWindowThreadProcessId((IntPtr)((Microsoft.Office.Interop.Excel.Application)Obj).Hwnd, out pID);
                    Process.GetProcessById((int)pID).Kill();
                }
            }
           catch { }
        }


        public static string getMyIP()
        {
            string myIP = "0.0.0.0";

            try
            {
                using (WebClient wc = new WebClient())
                {
                    string RequestHtml = wc.DownloadString("Http://www.youhost.co.kr/ip.php");

                    //MessageBox.Show(RequestHtml)
                    string[] separ = new string[] { "접근하신 IP 주소는", " 입니다" };
                    string[] result = RequestHtml.Split(separ, StringSplitOptions.RemoveEmptyEntries);

                    myIP = result[1].Trim();
                    //myIP = RequestHtml.Substring("접근하신 IP 주소는 ".Length, RequestHtml.IndexOf(" 입니다.") - "접근하신 IP 주소는 ".Length).Trim();
                    //MessageBox.Show(USER_IP)
                }
            }
            catch { }
            return myIP;
        }



        //******************************************************************************************************************************************************
        // [연도 선택] (DBConn Ver) 콤보박스 셋팅. 쿼리로 연도값을 셋팅한다. Default : 현재년
        //******************************************************************************************************************************************************
        public static void SET_cb_Year(ComboBox obj)
        {
            obj.Items.Clear();

            for (int i = DateTime.Now.Year + 1 ; i >= 2013 ; i += -1)
            {
                obj.Items.Add(i.ToString() + "년");
            }
            obj.SelectedIndex = 1;
        }





        //******************************************************************************************************************************************************
        // 로그를 기록함
        //******************************************************************************************************************************************************
        private static Log log = new Log();
        public static void LogWrite(string log_text)
        {
            // 로그 기록 폴더        
            log.DefaultFileLogWriter.CustomLocation = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\RealSMART_Log";
            // 로그 파일 명(프로그램명_날짜)        
            log.DefaultFileLogWriter.BaseFileName = "ErrorLog_" + DateTime.Now.ToString("yyyy-MM-dd");
            // 로그 내용 기록        
            log.WriteEntry(String.Format(DateTime.Now.ToString(), "yyyy-MM-dd HH:mm:ss") + "  ===  " + log_text, TraceEventType.Information);
            // 로그 기록 닫기        
            log.DefaultFileLogWriter.Close();
        }



        //******************************************************************************************************************************************************
        // 경과시간 측정도구
        //******************************************************************************************************************************************************
        private static Stopwatch sw;
        private static Stopwatch sw2;
        public static void StartBenchmark()
        {
            sw = new Stopwatch();
            sw2 = new Stopwatch();
            sw.Start();
            sw2.Start();
        }

        public static string ElapsedBenchMark()
        {
            TimeSpan ts = sw.Elapsed;
            return String.Format("{0:00}:{1:00}:{2:00}.{3:000}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds);
        }
        public static string ElapsedBenchMark2()
        {
            TimeSpan ts = sw2.Elapsed;
            sw2.Restart();
            return String.Format("{0:00}:{1:00}:{2:00}.{3:000}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds);
        }

        public static void StopBenchmark()
        {
            try
            {
                if (sw != null)
                {
                    sw.Stop();
                    sw = null;
                    sw2.Stop();
                    sw2 = null;
                }
            }
            catch { }
        }


        //******************************************************************************************************************************************************
        // 저장할 사용자 파일명을 받아옴
        // loCommon.loFunctions.getOpenFileName("Xml Files (*.xml)|*.xml", "xml", "줄임말 파일을 선택해 주세요.");
        //******************************************************************************************************************************************************
        public static string getOpenFileName(string FilterStr = "", string DftExt = "", string Title = "", string initDirectory = "")
        {
            string OFileNM = "";

            //***** 저장시킬 파일명을 만들어 받아온다.
            System.Windows.Forms.OpenFileDialog oFO = new System.Windows.Forms.OpenFileDialog();

            var _with1 = oFO;
            _with1.Multiselect = false;
            _with1.Filter = FilterStr;
            _with1.Title = Title;
            _with1.InitialDirectory = initDirectory;

            if (oFO.ShowDialog() == DialogResult.OK)
            {
                OFileNM = oFO.FileName;
            }

            return OFileNM;
        }


        //******************************************************************************************************************************************************
        // 저장할 사용자 파일명을 받아옴
        //******************************************************************************************************************************************************
        public static string getSaveFileName(string FilterStr = "", string DftExt = "", string DftFileNM = "", string DftFolder = "")
        {
            string SFileNM = "";

            //***** 저장시킬 파일명을 만들어 받아온다.
            System.Windows.Forms.SaveFileDialog oFS = new System.Windows.Forms.SaveFileDialog();

            var _with2 = oFS;
            _with2.Title = "저장될 파일명을 입력해 주세요.";
            if (!string.IsNullOrEmpty(FilterStr))
                _with2.Filter = FilterStr;
            if (!string.IsNullOrEmpty(DftExt))
                _with2.DefaultExt = DftExt;
            if (!string.IsNullOrEmpty(DftFileNM))
                _with2.FileName = DftFileNM;
            if (!string.IsNullOrEmpty(DftFolder))
                _with2.InitialDirectory = DftFolder;
            _with2.RestoreDirectory = true;

            if (oFS.ShowDialog() == DialogResult.OK)
            {
                SFileNM = oFS.FileName;
            }

            return SFileNM;
        }

        /// <summary>
        /// 파일명에 잘못된 문자가 포함되었는지 체크해서 리턴
        /// </summary>
        public static bool ValidFileName(string FileNM)
        {
            if (FileNM.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) != -1)
            {
                MessageBox.Show(string.Format("다음 문자는 사용할 수 없습니다.{0}\\ / : * ? \" < > ", Environment.NewLine));
                return false;
            }
            return true;
        }


        /// <summary>
        /// 파일명을 바꿔줌. 단, 파일이 이미 존재할 경우 (1), (2) 와 같이 자동 넘버링 
        /// </summary>
        public static void RenameFileExt(string srcFileNM, ref string tgtFileNM)
        {
            int count = 1;
            string FileNameOnly = System.IO.Path.GetFileNameWithoutExtension(tgtFileNM);
            string Extension = System.IO.Path.GetExtension(tgtFileNM);
            string path = System.IO.Path.GetDirectoryName(tgtFileNM);
            string newFullPath = tgtFileNM;

            while (File.Exists(newFullPath))
            {
                count += 1;
                string tmpFileNM = string.Format("{0} ({1})", FileNameOnly, count);
                newFullPath = System.IO.Path.Combine(path, tmpFileNM + Extension);
            }

            try
            {
                File.Move(srcFileNM, newFullPath);
                tgtFileNM = newFullPath;
            }
            catch { }

        }



        //******************************************************************************************************************************************************
        // 엑셀파일 시트명 받아오기
        //******************************************************************************************************************************************************
        public static List<string> ListSheetInExcel(string filePath)
        {
            System.Data.OleDb.OleDbConnectionStringBuilder sbConnection = new System.Data.OleDb.OleDbConnectionStringBuilder();
            List<string> listSheet = new List<string>();
            String strExtendedProperties = String.Empty;
            sbConnection.DataSource = filePath;
            if (Path.GetExtension(filePath).Equals(".xls"))
            {
                //for 97-03 Excel file
                sbConnection.Provider = "Microsoft.Jet.OLEDB.4.0";
                //HDR=ColumnHeader,IMEX=InterMixed
                strExtendedProperties = "Excel 8.0;HDR=Yes;IMEX=1";
            }
            else if (Path.GetExtension(filePath).Equals(".xlsx"))
            {
                //for 2007 Excel file
                sbConnection.Provider = "Microsoft.ACE.OLEDB.12.0";
                strExtendedProperties = "Excel 12.0;HDR=Yes;IMEX=1";
            }
            sbConnection.Add("Extended Properties", strExtendedProperties);
            using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(sbConnection.ToString()))
            {
                conn.Open();
                DataTable dtSheet = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

                foreach (DataRow drSheet in dtSheet.Rows)
                {
                    if (drSheet["TABLE_TYPE"].ToString() == "TABLE" & (drSheet["TABLE_NAME"].ToString().EndsWith("$") | drSheet["TABLE_NAME"].ToString().EndsWith("$'") | !(drSheet["TABLE_NAME"].ToString().Contains("xlnm#"))))
                    {
                        //If drSheet("TABLE_NAME").ToString().Contains("$") Then
                        //checks whether row contains '_xlnm#_FilterDatabase' or sheet name(i.e. sheet name always ends with $ sign)
                        listSheet.Add(drSheet["TABLE_NAME"].ToString());
                    }
                }
            }
            return listSheet;
        }


        //******************************************************************************************************************************************************
        // 진행률 상태 갱신.
        //******************************************************************************************************************************************************
        public static void UpdateProgress(ProgressBar Prog, Label lblProg, int nVal, string Msg)
        {
            Prog.Visible = true;
            lblProg.Visible = true;

            Prog.Value = nVal;
            lblProg.Text = Msg;

            Prog.Refresh();
            Application.DoEvents();
        }




        //******************************************************************************************************************************************************
        // 셀렉트된 라디오박스 객체 리턴
        //******************************************************************************************************************************************************
        public static RadioButton getSelectedOne(params RadioButton[] rb)
        {
            RadioButton rtnrb = null;

            for (int i = 0 ; i <= rb.Length - 1 ; i++)
            {
                if (rb[i].Checked)
                {
                    rtnrb = rb[i];
                    break;
                }
            }
            return rtnrb;
        }

        //******************************************************************************************************************************************************
        // 셀렉트된 체크박스의 텍스트 목록 리턴
        //******************************************************************************************************************************************************
        public static List<string> getCheckedItem(params CheckBox[] chk)
        {
            List<string> checkedString = new List<string>();

            for (int i = 0 ; i <= chk.Length - 1 ; i++)
            {
                if (chk[i].Checked)
                {
                    checkedString.Add(chk[i].Text);
                }
            }
            return checkedString;
        }


        //******************************************************************************************************************************************************
        // AdminMode Check
        //******************************************************************************************************************************************************
        public static bool checkRunningAuthorizer()
        {
            Process[] oProcess = Process.GetProcessesByName("RealSMART_ApplicationAuthorizer");
            if (oProcess.Length > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        /// <summary>
        /// 업로드가 실패한 파일을 바탕화면으로 따로 빼준다.
        /// </summary>
        public static void copy_UploadFaild_File(string fileNM)
        {
            string targetDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), string.Format("uploadFailed_{0}", DateTime.Now.ToShortDateString()));
            if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);

            try
            {
                File.Copy(fileNM, Path.Combine(targetDir, Path.GetFileName(fileNM)));
            }
            catch (Exception)
            { }
        }


        #region DataTable <-> ADODB.Recordset Convert Functions

        /// *************************************************************************************************************************
        /// ADODB.Recordset 객체를 System.Data.DataTable 로 변환
        /// *************************************************************************************************************************
        public static DataTable ConvertToDataTable(ADODB.Recordset objRS)
        {
            System.Data.OleDb.OleDbDataAdapter objDA = new System.Data.OleDb.OleDbDataAdapter();
            DataTable objDT = new DataTable();
            objDA.Fill(objDT, objRS);
            return objDT;
        }

        /// *************************************************************************************************************************
        /// System.Data.DataTable를 ADODB.Recordset객체로 변환
        /// *************************************************************************************************************************
        public static ADOR.Recordset ConvertToRecordset(DataTable inTable)
        {
            ADOR.Recordset result = new ADOR.Recordset();
            result.CursorLocation = ADOR.CursorLocationEnum.adUseClient;

            ADOR.Fields resultFields = result.Fields;
            System.Data.DataColumnCollection inColumns = inTable.Columns;

            foreach (DataColumn inColumn in inColumns)
            {
                resultFields.Append(inColumn.ColumnName, TranslateType(inColumn.DataType), inColumn.MaxLength, (inColumn.AllowDBNull ? ADOR.FieldAttributeEnum.adFldIsNullable : ADOR.FieldAttributeEnum.adFldUnspecified), null);
            }

            result.Open(System.Reflection.Missing.Value, System.Reflection.Missing.Value, ADOR.CursorTypeEnum.adOpenStatic, ADOR.LockTypeEnum.adLockOptimistic, 0);

            foreach (DataRow dr in inTable.Rows)
            {
                result.AddNew(System.Reflection.Missing.Value, System.Reflection.Missing.Value);

                for (int columnIndex = 0 ; columnIndex <= inColumns.Count - 1 ; columnIndex++)
                {
                    resultFields[columnIndex].Value = dr[columnIndex];
                }
            }

            result.MoveFirst();
            return result;
        }


        private static ADOR.DataTypeEnum TranslateType(Type columnType)
        {
            switch (columnType.UnderlyingSystemType.ToString())
            {
                case "System.Boolean":

                    return ADOR.DataTypeEnum.adBoolean;
                case "System.Byte":

                    return ADOR.DataTypeEnum.adUnsignedTinyInt;
                case "System.Char":

                    return ADOR.DataTypeEnum.adChar;
                case "System.DateTime":

                    return ADOR.DataTypeEnum.adDate;
                case "System.Decimal":

                    return ADOR.DataTypeEnum.adCurrency;
                case "System.Double":

                    return ADOR.DataTypeEnum.adDouble;
                case "System.Int16":

                    return ADOR.DataTypeEnum.adSmallInt;
                case "System.Int32":

                    return ADOR.DataTypeEnum.adInteger;
                case "System.Int64":

                    return ADOR.DataTypeEnum.adBigInt;
                case "System.SByte":

                    return ADOR.DataTypeEnum.adTinyInt;
                case "System.Single":

                    return ADOR.DataTypeEnum.adSingle;
                case "System.UInt16":

                    return ADOR.DataTypeEnum.adUnsignedSmallInt;
                case "System.UInt32":

                    return ADOR.DataTypeEnum.adUnsignedInt;
                case "System.UInt64":

                    return ADOR.DataTypeEnum.adUnsignedBigInt;
                default:
                    return ADOR.DataTypeEnum.adVarChar;
            }
        }

        #endregion






        //******************************************************************************************************************************************************
        // 엑셀파일 시트명 받아오기 (C1Excel)
        //******************************************************************************************************************************************************
        public static List<string> ListSheetInC1Excel(string filePath)
        {
            List<string> listSheet = new List<string>();

            using (C1XLBook wb = new C1XLBook())
            {
                try
                {
                    if (Path.GetExtension(filePath).Equals(".xls"))
                    {
                        //for 97-03 Excel file
                        wb.Load(filePath, false);
                    }
                    else if (Path.GetExtension(filePath).Equals(".xlsx"))
                    {
                        //for 2007 Excel file
                        wb.Load(filePath, FileFormat.OpenXml, false);
                    }

                    if (wb.Sheets.Count > 0)
                        for (int i = 0 ; i < wb.Sheets.Count ; i++)
                        {
                            listSheet.Add(wb.Sheets[i].Name);
                        }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("엑셀파일을 불러올 수 없습니다. " + ex.Message);
                }
            }

            return listSheet;
        }

        //******************************************************************************************************************************************************
        // 엑셀파일 시트 데이터 받아오기 (C1Excel)
        //******************************************************************************************************************************************************
        public static DataTable C1ExcelSheetToDataTable(string filePath, string shtName, string FailedMessage = "", int startColIndex = 0, int startRowIndex = 0, int endColIndex = 0, int endRowIndex = 0, bool includeHeader = true)
        {
            DataTable rtnTbl = new DataTable();

            using (C1XLBook wb = new C1XLBook())
            {
                try
                {
                    if (Path.GetExtension(filePath).Equals(".xls"))
                    {
                        //for 97-03 Excel file
                        wb.Load(filePath, true);
                    }
                    else if (Path.GetExtension(filePath).Equals(".xlsx"))
                    {
                        //for 2007 Excel file
                        wb.Load(filePath, FileFormat.OpenXml, true);
                    }

                    if (wb.Sheets.Count > 0)
                    {
                        if (wb.Sheets[shtName] != null)
                        {
                            var sht = wb.Sheets[shtName];

                            for (int i = startColIndex ; i <= (endColIndex == 0 ? sht.Columns.Count -1 : endColIndex) ; i++)
                            {
                                if (includeHeader)
                                    rtnTbl.Columns.Add(sht[startRowIndex, i].Value != null ? sht[startRowIndex, i].Text : "");
                                else
                                    rtnTbl.Columns.Add(i.ToString());
                            }

                            for (int i = startRowIndex + (includeHeader ? 1 : 0) ; i <= (endRowIndex == 0 ? sht.Rows.Count -1 : endRowIndex) ; i++)
                            {
                                object[] rowValues = new object[rtnTbl.Columns.Count];
                                int colIdx = 0;
                                for (int j = startColIndex ; j <= (endColIndex == 0 ? sht.Columns.Count - 1 : endColIndex) ; j++)
                                {
                                    rowValues[colIdx] = sht[i, j].Value != null ? sht[i, j].Text : null;
                                    colIdx += 1;
                                }

                                if (rowValues.Count(x => x == null) != rowValues.Count())    // 모든 열 값이 공란인 행은 추가하지 않음.
                                    rtnTbl.Rows.Add(rowValues);
                            }
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(FailedMessage + Environment.NewLine + ex.Message);
                }
            }

            return rtnTbl;
        }



        public static double StdDev(IEnumerable<double> values)
        {
            double result = 0;
            try
            {
                if (values.Count() > 0)
                {
                    double mean = values.Average();
                    double sum = values.Sum(d => Math.Pow(d - mean, 2));
                    result = Math.Sqrt((sum) / (values.Count() - 1));
                }
            }
            catch { }
            return result;
        }

    }
}
