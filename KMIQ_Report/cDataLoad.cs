using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Interop;
using loCommon;
using System.ComponentModel;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;

namespace KMIQ
{
    public class cDataLoad
    {
        BackgroundWorker bg_Worker;
        public DataTable MainData { get; set; }
        public cDataLoad()
        {
            bg_Worker = new BackgroundWorker();
            bg_Worker.WorkerReportsProgress = false;
            bg_Worker.WorkerSupportsCancellation = false;
            bg_Worker.DoWork += Bg_Worker_DoWork;
            
        }

        private void Bg_Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            LoadData();
            UpdateDB();
        }


        #region public process


        public void ClearDBCheck()
        {
            if (MessageBox.Show(Globals.ThisWorkbook.Main.ClearDB_Confirm_Message, "Clear DB", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //ClearDB();
                if (MainData != null) MainData.Dispose();
                MainData = null;

                RibbonComboBox combo = (RibbonComboBox)Globals.Ribbons.Menu.cb_Student;
                combo.Items.Clear();
                combo.Text = Globals.ThisWorkbook.Main.PersonCombo_Selector;


                Globals.ThisWorkbook.Main.MakeReport.ClearAllReport();
                MessageBox.Show(Globals.ThisWorkbook.Main.ClearDB_Complete_Message);
            }
        }

        /// <summary>
        /// 파일을 직접 불러와서 리포트를 생성한다.
        /// </summary>
        public void GetOpenFile()
        {
            string dataFileName = loFunctions.getOpenFileName("엑셀 데이터 파일(*.xlsx)|*.xlsx", "xlsx", "데이터 파일을 선택하세요.", Environment.GetFolderPath(Environment.SpecialFolder.Desktop));

            if (dataFileName != "")
            {
                var shtList = loFunctions.ListSheetInC1Excel(dataFileName);

                if (shtList.Count > 0 && shtList[0] == "Data")
                {
                    MainData = loFunctions.C1ExcelSheetToDataTable(dataFileName, shtList[0], "데이터 파일 불러오기 실패!", 0, 2, 126, 0, false);

                    if (MainData.Rows.Count > 0)
                    {
                        MessageBox.Show("데이터를 불러왔습니다. 대상자를 선택하세요.");
                        
                    }
                }
                else
                {
                    MessageBox.Show("정상적인 Data파일로 인식하지 못했습니다. SMART OMR 프로그램을 통해 저장된 데이터 파일을 그대로 사용하셔야 합니다.");
                }
            }

        }

        #endregion

  

        #region private process

        List<UploadedInfo> CurrentFileList;
        List<UploadedInfo> UploadedFileList;
        DataTable AnswerTB;
        Form.frmProgress f_Progress;

        private void ShowProgress()
        {
            f_Progress = new Form.frmProgress();
            new WindowInteropHelper(f_Progress).Owner = Globals.ThisWorkbook.Main.G_WINDOW_HANDLE;
            f_Progress.updateProgress2(true, Globals.ThisWorkbook.Main.UpdateDB_During_Message);
            f_Progress.Show();
        }

        private void CloseProgress()
        {
            f_Progress.Close();
        }

        /// <summary>
        /// DB에 저장된 업데이트 정보와 현재 폴더 내의 파일 정보를 비교해 다른점이 있는지만 간단 체크한다.
        /// </summary>
        bool IsUpdatable()
        {
            bool rtn = true;

            //int iRow = 1;
            CurrentFileList = new List<UploadedInfo>();
            DirectoryInfo dr = new DirectoryInfo(Globals.ThisWorkbook.Main.G_DATA_FOLDER);
            if (dr.Exists)
            {
                foreach (FileInfo f in dr.GetFiles("*.*", SearchOption.AllDirectories))
                {
                    if (!f.Name.StartsWith("~$"))
                    {
                        CurrentFileList.Add(new UploadedInfo()
                        {
                            Order = CurrentFileList.Count + 1,
                            FullName = f.FullName,
                            Exam = f.DirectoryName.Replace(dr.FullName, "").Replace(@"\", ""),
                            FileName = f.Name,
                            Type = getDataType(f.Name),
                            LastDate = f.LastWriteTime,
                            Different = true
                        });
                        //Globals.Sheet9.Cells[iRow, 1].Value = f.FullName;
                        //Globals.Sheet9.Cells[iRow, 2].Value = f.Directory;
                        //Globals.Sheet9.Cells[iRow, 3].Value = f.DirectoryName.Replace(dr.FullName, "").Replace(@"\", "");
                        //Globals.Sheet9.Cells[iRow, 4].Value = f.Name;
                        //Globals.Sheet9.Cells[iRow, 5].Value = getDataType(f.Name);
                        //Globals.Sheet9.Cells[iRow, 6].Value = f.LastWriteTime;
                        //iRow++;
                    }

                }
            }

            UploadedFileList = getUploadedFileList();

            CurrentFileList.Sort((UploadedInfo x, UploadedInfo y) => x.Order.CompareTo(y.Order));
            UploadedFileList.Sort((UploadedInfo x, UploadedInfo y) => x.Order.CompareTo(y.Order));


            foreach (UploadedInfo f in CurrentFileList)
            {
                /// 기존의 파일목록과 파일명 / 변경시각을 비교해서  변화 여부를 체크한다.  
                /// Different = True : 달라짐(기존리스트에서 찾을수 없음. Any 리턴값은 false) 
                /// Different = False : 안달라짐 (기존리스트에 이미 있음. Any 리턴값은 true)
                f.Different = !UploadedFileList.Any(x => x.FileName.Equals(f.FileName) && x.LastDate.ToString().Equals(f.LastDate.ToString()));
            }

            /// 달라짐 값이 1개 이상 있으면
            rtn = CurrentFileList.Where(x => x.Different).Count() > 0 ? true : false;
            //rtn = !UploadedFileList.SequenceEqual(CurrentFileList);
            
            return rtn;
        }

        private List<UploadedInfo> getUploadedFileList()
        {
            List<UploadedInfo> rtn = new List<UploadedInfo>();

            //using (SqliteConn sqlite = new SqliteConn(Globals.ThisWorkbook.Main.G_DB_FILE))
            //{
            //    DataTable tbl = sqlite.SqlSelect("SELECT * FROM T_Uploaded");
            //    if (tbl.Rows.Count > 0)
            //    {
            //        foreach (DataRow r in tbl.Rows)
            //        {
            //            rtn.Add(new UploadedInfo()
            //            {
            //                Order = rtn.Count + 1,
            //                FullName = r["FullName"].ToString(),
            //                Exam = r["Exam"].ToString(),
            //                FileName = r["FileName"].ToString(),
            //                Type = getDataType(r["Type"].ToString()),
            //                LastDate = Convert.ToDateTime(r["LastDate"])
            //            });

            //        }
            //    }
            //}

            return rtn;
        }

        private DataType getDataType(string p)
        {
            DataType rtn = DataType.Not;
            string fname = p.ToLower();

            if (fname.Contains("data"))
                rtn = DataType.Data;
            else if (fname.Contains("정답"))
                rtn = DataType.Answer;
            //else if (fname.Contains("conv") || fname.Contains("table") || fname.Contains("cvt") || fname.Contains("convert"))
            //    rtn = DataType.Convert;
            else if (fname.Contains("커멘트"))
                rtn = DataType.Comment;
            else if (fname.Contains("코드"))
                rtn = DataType.Code;

            return rtn;
        }

        /// <summary>
        /// 폴더에서 데이터 불러와 업로드 처리한다.
        /// 변경된 데이터를 추적하는 액션 또한 여기 포함된다.
        /// </summary>
        void LoadData()
        {

        }


        /// <summary>
        /// 지정된 영역을 이용해 엑셀로부터 데이터테이블을 받는다. 
        /// </summary>
        private DataTable getDataTable(string fileName, List<string> shtNameList, C1ExcelArea area, string dataName, bool includeHeader = true)
        {
            DataTable rtn = new DataTable();

            var shts = shtNameList.Where(x => x.ToLower().Contains(area.shtName.ToLower())).ToList();
            string shtName = "";
            if (shts.Count > 0) shtName = shts[0];

            if (shtName != "")
            {
                rtn = loFunctions.C1ExcelSheetToDataTable(fileName, shtName, dataName + " 불러오기 실패", area.startColIndex, area.startRowIndex, area.endColIndex, area.endRowIndex, includeHeader);
            }

            return rtn;
        }


        /// <summary>
        /// 특정 시험에 달린 모든 데이터를 지워준다.
        /// </summary>
        private void ClearDB(string exam)
        {
            //string[] examTables = new string[] { "T_Student", "T_Answer", "T_Score", "T_RptBase", "T_Accumulate" };
            //using (SqliteConn sqlite = new SqliteConn(Globals.ThisWorkbook.Main.G_DB_FILE))
            //{
            //    foreach (string tblName in examTables)
            //        sqlite.SqlExcute(string.Format("DELETE  FROM {0} WHERE Exam = '{1}'", tblName, exam));
            //}
        }

        
        /// <summary>
        /// 각 테이블을 실제 db에 적용해 주는 프로시저
        /// </summary>
        private void ExecuteUpload(DataTable tbl, string tblName, string exam)
        {
            //using (SqliteConn sqlite = new SqliteConn(Globals.ThisWorkbook.Main.G_DB_FILE))
            //{
            //    sqlite.SqlBulkInsert(tbl, tblName);
            //}
        }



        /// <summary>
        /// 현재 업로드 된 파일 상태를 DB에 업데이트한다.
        /// </summary>
        private void UpdateDB()
        {
            //using (SqliteConn sqlite = new SqliteConn(Globals.ThisWorkbook.Main.G_DB_FILE))
            //{
            //    sqlite.SqlExcute("DELETE FROM T_Uploaded");

            //    if (CurrentFileList.Count > 0)
            //    {
            //        foreach (UploadedInfo f in CurrentFileList)
            //        {
            //            try
            //            {
            //                string qry = string.Format("INSERT INTO T_Uploaded VALUES ('{0}', '{1}', '{2}', '{3}', '{4}')", f.FullName, f.Exam, f.FileName, f.Type.ToString(), f.LastDate.ToString());
            //                sqlite.SqlExcute(qry);
            //            }
            //            catch (Exception ex)
            //            {
            //                loFunctions.LogWrite("UpdateDB : 파일목록 업로드 오류 - " + ex.Message);
            //                throw;
            //            }
            //        }
            //    }
            //}
        }

        #endregion



        #region 정/오 여부 판단, 채점

        public string getErrata(DataRow r)
        {
            string Errata = "";

            int Sequence = Convert.ToInt32(r["Sequence"].ToString());
            string Answer = r["Answer"].ToString();
            string Correct = r["CorrectAnswer"].ToString();
            string[] colArr;
            string[] ansArr;


            if (Correct != "")      // 정답이 있는 경우만 채점.
            {
                bool ox = false;            // 정오 여부
                bool shortAns = false;      // 주관식 여부

                /// 문항번호가 112~116, 147~154일 때 주관식으로 인식한다. 문항번호를 Main의 글로벌 상수값에서 체크하는 것으로 수정함.
                //if ((Sequence >= 68 && Sequence <= 72) || (Sequence >= 147 && Sequence <= 154)) shortAns = true;
                //shortAns = Globals.ThisWorkbook.Main.ShortAnswer.Contains(Sequence);

                //***** 문항번호와 Flag값들로 정답과 채점옵션을 찾아온다.
                try
                {
                    /// //////////////////////////////////////////////////////////////////////
                    /// 객관식 문항 채점
                    /// //////////////////////////////////////////////////////////////////////
                    if (!shortAns)      
                    {
                        switch (r["AnswerStyle"].ToString())
                        {

                            case "NONE":
                                ////문항당 단 하나의 정답만 있으며 응답도 단일 답이어야 함. 정답과 응답이 무조건 일치해야 맞음
                                ox = (Answer == Correct ? true : false);
                                break;

                            case "AND":
                                ////문항당 여러개의 정답이 있고 이 여러 개의 답이 모두 일치해야만 맞음.
                                //Dim q As IEnumerable(Of Integer) = From p As String In CollectAns.Replace(" ", "").Split(",").AsEnumerable
                                //                                    Join pp As String In QAnswer.Replace(" ", "").Split(",").AsEnumerable
                                //                                    On p(0) Equals pp(0)
                                //                                    Where pp(0) <> ""
                                //                                    Select p.Count
                                colArr = Correct.Replace(" ", "").Split(',');
                                ansArr = Answer.Replace(" ", "").Split(',');
                                Array.Sort(colArr);
                                Array.Sort(ansArr);
                                ox = (colArr.SequenceEqual(ansArr) ? true : false);
                                break;

                            case "OR":
                                ////문항당 여러개의 정답이 있고 여러개의 응답이 있음. 이 여러 개의 답 중에 응답이 포함되어 있으면 맞음.
                                ////정답갯수쪼개고 응답을 돌면서 비교.
                                ox = Correct.Replace(" ", "").Split(',').Union(Answer.Replace(" ", "").Split(',')).Distinct().SequenceEqual(Correct.Replace(" ", "").Split(','));
                                break;

                            case "ALL":
                                ////무조건 다맞음
                                ox = true;

                                break;

                            case "RANGE":
                                ////범위답  +++++++++++ 객관식에서는 RANGE를 쓰지 않으나, 프로그램의 기능상 구현되어야 하므로 포함시켰다.
                                ox = true;

                                break;
                        }
                    }
                    /// //////////////////////////////////////////////////////////////////////
                    /// 주관식 문항 채점
                    /// //////////////////////////////////////////////////////////////////////
                    else
                    {
                        switch (r["AnswerStyle"].ToString())
                        {

                            case "NONE":
                                ////문항당 단 하나의 정답만 있으며 응답도 단일 답이어야 함. 정답과 응답이 무조건 일치해야 맞음
                                ox = (Answer == Correct ? true : false);
                                break;

                            case "AND":
                                ////문항당 여러개의 정답이 있고 이 여러 개의 답이 모두 일치해야만 맞음.     ++++++++++ 주관식에서는 AND를 쓰지 않으나, 프로그램의 기능상 구현되어야 하므로 포함시켰다.
                                colArr = Correct.Replace(" ", "").Split(',');
                                ansArr = Answer.Replace(" ", "").Split(',');
                                Array.Sort(colArr);
                                Array.Sort(ansArr);
                                ox = (colArr.SequenceEqual(ansArr) ? true : false);
                                break;

                            case "OR":
                                ////문항당 여러개의 정답이 있고 여러개의 응답이 있음. 이 여러 개의 답 중에 응답이 포함되어 있으면 맞음.
                                ////정답갯수쪼개고 응답을 돌면서 비교. 정답과 응답에 분수가 있으면 모두 소수로 변경해서 처리한다.
                                colArr = Correct.Replace(" ", "").Split(',');
                                ansArr = Answer.Replace(" ", "").Split(',');
                                
                                List<decimal> colValList = ConvertFractionToDecimal(colArr);
                                List<decimal> ansValList = ConvertFractionToDecimal(ansArr);

                                ox = colValList.Union(ansValList).Distinct().SequenceEqual(colValList.Distinct());
                                break;

                            case "ALL":
                                ////무조건 다맞음
                                ox = true;

                                break;

                            case "RANGE":
                                ////정답의 범위에 속하면 맞음. '정답'지칭 알파벳은 a, x, A, X를 지원한다. 정답을 소문자로 변환하고 범위기호 추출 및 범위숫자 추출 후 decimal로 변환한다.
                                //// 응답도 decimal로 변환하고 범위에 대입해 처리한다.

                                colArr = Correct.ToLower().Replace("x", "a").Split('a');
                                ansArr = Answer.Replace(" ", "").Split(',');
                                string[] signList = new string[2];

                                // 범위기호와 범위숫자를 추출한다.
                                for (int i = 0 ; i < colArr.Length ; i++)
                                {
                                    if (colArr[i].Contains("<="))
                                    {
                                        signList[i] = "<=";
                                        colArr[i] = colArr[i].Replace("<=", "").Trim();
                                    }
                                    else if (colArr[i].Contains("<"))
                                    {
                                        signList[i] = "<";
                                        colArr[i] = colArr[i].Replace("<", "").Trim();
                                    } 
                                }

                                // 범위숫자를 decimal list로 만든다.
                                List<decimal> colRangeList = ConvertFractionToDecimal(colArr);
                                List<decimal> ansNumList = ConvertFractionToDecimal(ansArr);        // 무조건 0번 인덱스 (1개의 값)만 사용할 것이다.

                                ox = false;
                                // x < A < y , x < A <= y,  x < A ,  x <= A < y , x <= A <= y,  x <= A ,  A < y ,  A <= y , A = X 
                                if (signList[0] == "<")
                                {
                                    if (signList[1] == "<")
                                        ox = colRangeList[0] < ansNumList[0] && ansNumList[0] < colRangeList[1];
                                    else if (signList[1] == "<=")
                                        ox = colRangeList[0] < ansNumList[0] && ansNumList[0] <= colRangeList[1];
                                    else if (signList[1] == "")
                                        ox = colRangeList[0] < ansNumList[0];
                                }
                                else if (signList[0] == "<=")
                                {
                                    if (signList[1] == "<")
                                        ox = colRangeList[0] <= ansNumList[0] && ansNumList[0] < colRangeList[1];
                                    else if (signList[1] == "<=")
                                        ox = colRangeList[0] <= ansNumList[0] && ansNumList[0] <= colRangeList[1];
                                    else if (signList[1] == "")
                                        ox = colRangeList[0] <= ansNumList[0];
                                }
                                else if (signList[0] == "")
                                {
                                    if (signList[1] == "<")
                                        ox = ansNumList[0] < colRangeList[1];
                                    else if (signList[1] == "<=")
                                        ox = ansNumList[0] <= colRangeList[1];
                                    else if (signList[1] == "")
                                        ox = colRangeList[0] == ansNumList[0];
                                }

                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    ox = false;
                }
                Errata = ox.ToString();
            }
            else
            {
                Errata = "";
            }

            return Errata;
        }

        /// <summary>
        /// 분수-소수 및 숫자 변환 함수
        /// </summary>
        private List<decimal> ConvertFractionToDecimal(string[] colArr)
        {
            List<decimal> rtn = new List<decimal>();

            foreach (string c in colArr)
            {
                if (c.Contains("/"))
                {
                    if (c.Split('/').Length != 2)
                        rtn.Add(999999);        // 정상적인 분수가 아니라면 무조건 틀리도록 만든다.
                    else
                    {
                        try
                        {
                            rtn.Add(Math.Round(Convert.ToDecimal(c.Split('/')[0]) / Convert.ToDecimal(c.Split('/')[1]), 2));
                        }
                        catch (Exception ex)
                        {
                            rtn.Add(999999);        // 분수를 소수로 변환하는데 실패해도 무조건 틀리도록 만든다.
                            loFunctions.LogWrite("ConvertFractionToDecimal : 분수-소수 변환 실패 - " + ex.Message);
                        }
                    }
                }
                else
                {
                    try
                    {
                        if (c != "")
                            rtn.Add(Convert.ToDecimal(c));
                        else rtn.Add(999999);
                    }
                    catch (Exception ex)
                    {
                        rtn.Add(999999);        // 정답텍스트를 숫자로 변환하는데 실패해도 무조건 틀리도록 만든다.
                        loFunctions.LogWrite("ConvertFractionToDecimal : 정답-숫자 변환 실패 - " + ex.Message);
                    }
                }
            }

            return rtn;
        }


        private void ClearDB()
        {
            string databasePath = Globals.ThisWorkbook.Main.G_DB_FILE;
            string emptyPath = Globals.ThisWorkbook.Main.G_EMPTYDB_FILE;

            if (File.Exists(databasePath)) File.Copy(databasePath, string.Format("{0}_Backup_{1}", databasePath, DateTime.Now.ToString("yyyyMMdd_HHmmss")), true);
            if (File.Exists(emptyPath)) File.Copy(emptyPath, databasePath, true);
        }



        #endregion

        public class UploadedInfo
        {
            public int Order { get; set; }
            public string FullName { get; set; }
            public string Exam { get; set; }
            public string FileName { get; set; }
            public DataType Type { get; set; }
            public DateTime LastDate { get; set; }
            public bool Different { get; set; }
        }

        public enum DataType
        {
            Data,
            Answer,
            Convert,
            Essay,
            ClassCode,
            Code,
            Comment,
            Not
        }

    }
}
