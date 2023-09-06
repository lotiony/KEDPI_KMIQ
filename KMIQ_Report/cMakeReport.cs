using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using loCommon;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using XL = Microsoft.Office.Interop.Excel;

namespace KMIQ
{
    public class cMakeReport
    {
        #region const
        #endregion

        #region common variables
        public string SelectedExam = "";
        Form.frmProgress f_Progress;
        string qry = "";
        private Dictionary<string, string> personalReport;

        private Dictionary<string, string> commonData;


        DataTable defineTable;
        DataTable mainData { get { return Globals.ThisWorkbook.Main.DataLoad.MainData; } }
        DataSet resultData;
        List<string> strongArea;
        List<string> weakArea;
        #endregion


        public cMakeReport()
        {
            defineTable = getCommonData();
            defineTable.Columns.Add("Answer", typeof(string));
        }

        private DataTable getCommonData()
        {
            DataTable rtn = new DataTable();


            qry = "SELECT * FROM T_DATA_DEFINE";

            using (SqlConn sql = new SqlConn())
            {
                try
                {
                    rtn = sql.SqlSelect(qry);
                }
                catch (Exception ex)
                {
                    loFunctions.LogWrite("[getCommonData] on Error : " + ex.Message);
                }
            }

            return rtn;

        }

        #region public process

        /// <summary>
        /// 불러온 데이터에서 학생 목록을 만든다.
        /// </summary>
        internal void SetMenuPersonalCombo(object sender, RibbonBase ribbon)
        {

            RibbonComboBox combo = (RibbonComboBox)sender;
            combo.Items.Clear();

            RibbonDropDownItem item = ribbon.Factory.CreateRibbonDropDownItem();
            item.Label = Globals.ThisWorkbook.Main.PersonCombo_Selector;
            combo.Items.Add(item);

            personalReport = getMakePersonalList();

            if (personalReport.Count > 0)
                foreach (string key in personalReport.Keys)
                {
                    RibbonDropDownItem examitem = ribbon.Factory.CreateRibbonDropDownItem();
                    examitem.Label = personalReport[key];
                    combo.Items.Add(examitem);
                }
        }


        /// <summary>
        /// 학생이 선택되었을 때 개인성적표를 만들고 보여준다.
        /// </summary>
        internal void SelectedPersonalCombo(object sender)
        {
            RibbonComboBox combo = (RibbonComboBox)sender;

            if (combo.Text != Globals.ThisWorkbook.Main.PersonCombo_Selector)
            {
                makeReport_Personal(combo.Text);
            }
        }

        /// <summary>
        /// 전체 리포트 클리어
        /// </summary>
        internal void ClearAllReport()
        {
            clear_Result(true);

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }


        /// <summary>
        /// 프린트 출력창 열기
        /// </summary>
        internal void ShowPrint()
        {
            if (Globals.ThisWorkbook.Main.DataLoad.MainData == null || Globals.ThisWorkbook.Main.DataLoad.MainData.Rows.Count == 0) { MessageBox.Show(Globals.ThisWorkbook.Main.NoReport_Message); return; }

            personalReport = getMakePersonalList();

            Form.frmPrint fPrint = new Form.frmPrint();
            new WindowInteropHelper(fPrint).Owner = Globals.ThisWorkbook.Main.G_WINDOW_HANDLE;
            fPrint.ShowDialog();
        }


        #endregion




        #region private process

        private void ShowProgress()
        {
            f_Progress = new Form.frmProgress();
            new WindowInteropHelper(f_Progress).Owner = Globals.ThisWorkbook.Main.G_WINDOW_HANDLE;
            f_Progress.initializeProgress(100);
            f_Progress.Show();
        }

        private void CloseProgress()
        {
            f_Progress.Close();
        }

        private void MakeReport(string exam)
        {
            ShowProgress();

            SelectedExam = exam;

            makeReportProcess();

            //int i=0;
            //while (i <= 100)
            //{
            //    Thread.Sleep(40);
            //    f_Progress.updateProgress(i, string.Format("리포트 만들고 있습니다. {0}", i)); 
            //    System.Windows.Forms.Application.DoEvents();
            //    i++;
            //}
            CloseProgress();

            Globals.P1.Activate();
        }


        private Dictionary<string, string> getMakePersonalList()
        {
            Dictionary<string, string> rtn = new Dictionary<string, string>();

            if (mainData != null)
            {
                foreach (DataRow r in mainData.Rows)
                {
                    rtn.Add(string.Format("{0}{1}", r[0], DBNull.Value.Equals(r[1]) ? "" : r[1]), string.Format("{0}. {1} ({2})", r[0], DBNull.Value.Equals(r[1]) ? "" : r[1], DBNull.Value.Equals(r[3]) ? "" : r[3]));
                }
            }

            return rtn;
        }


        #region make report process

        public void makeReport_Personal(string selectItem)
        {
            clear_Result();

            string selectedNo = selectItem.Split('.')[0];
            int sNo = 0;
            if (Int32.TryParse(selectedNo, out sNo))
            {
                var sItem = mainData.AsEnumerable().Where(x => x[0].ToString().Equals(selectedNo)).ToList();

                if (sItem.Count > 0)
                {
                    DataTable ansData = defineTable.Copy();
                    List<object> ansList = sItem[0].ItemArray.ToList();

                    /// 응답 데이터를 defined data table에 채운다.
                    for (int i = 0 ; i < ansData.Rows.Count ; i++)
                    {
                        ansData.Rows[i]["Answer"] = ansList[i].ToString();
                    }

                    commonData = new Dictionary<string, string>();
                    commonData.Add("이름", ansData.AsEnumerable().Where(x => x["DATA_TYPE"].ToString().Equals("이름")).Select(x => x["Answer"].ToString()).ToList()[0].ToString());
                    commonData.Add("검사일", ansData.AsEnumerable().Where(x => x["DATA_TYPE"].ToString().Equals("검사일")).Select(x => x["Answer"].ToString()).ToList()[0].ToString());
                    commonData.Add("응답구분", ansData.AsEnumerable().Where(x => x["DATA_TYPE"].ToString().Equals("응답구분")).Select(x => x["Answer"].ToString()).ToList()[0].ToString());
                    commonData.Add("성명", getDisplayName(commonData["이름"], commonData["응답구분"]));

                    string ansStr = string.Join(",", ansData.AsEnumerable().Where(x => x["DATA_TYPE"].ToString().Equals("ANSWER")).Select(x => x["Answer"].ToString()).ToList());


                    qry = string.Format(" EXEC [dbo].[USP_RPT_RESULT] {0}, '{1}'", getTypeID(commonData["응답구분"]), ansStr);
                    using (SqlConn sql = new SqlConn())
                    {
                        try
                        {
                            resultData = sql.SqlSelectMultiResult(qry);

                        }
                        catch (Exception ex)
                        {
                            loFunctions.LogWrite("[makeReport_Personal - USP_RPT_RESULT] on Error : " + ex.Message);
                        }
                    }

                    if (resultData.Tables.Count == 3)
                    {
                        makeReportProcess();
                    }
                    else
                    {
                        MessageBox.Show("DB에서 결과를 제대로 처리하지 못했습니다. 응답자구분이나 문항 응답 데이터를 확인해 보시기 바랍니다.");
                    }
                }
            }
            else
            {
                MessageBox.Show("선택된 항목의 데이터가 없거나 잘못 되었습니다.");
            }
        }

        public void makeReport_Online(DataRow dataItem)
        {
            clear_Result();

            if (dataItem["RAW_DATA"].ToString() != "")
            {

                commonData = new Dictionary<string, string>();
                commonData.Add("이름", DBNull.Value.Equals(dataItem["NAME"]) ? dataItem["UID"].ToString() : dataItem["NAME"].ToString() );
                commonData.Add("검사일", Convert.ToDateTime(dataItem["IN_DT"]).ToShortDateString());
                commonData.Add("성명", getDisplayName(commonData["이름"], dataItem["TYPE_GRADE"].ToString()));

                string ansStr = dataItem["RAW_DATA"].ToString();


                    qry = string.Format(" EXEC [dbo].[USP_RPT_RESULT] {0}, '{1}'", dataItem["TYPEID"], ansStr);
                    using (SqlConn sql = new SqlConn())
                    {
                        try
                        {
                            resultData = sql.SqlSelectMultiResult(qry);

                        }
                        catch (Exception ex)
                        {
                            loFunctions.LogWrite("[makeReport_Personal - USP_RPT_RESULT] on Error : " + ex.Message);
                        }
                    }

                    if (resultData.Tables.Count == 3)
                    {
                        makeReportProcess();
                    }
                    else
                    {
                        MessageBox.Show("DB에서 결과를 제대로 처리하지 못했습니다. 응답자구분이나 문항 응답 데이터를 확인해 보시기 바랍니다.");
                    }
            }
            else
            {
                MessageBox.Show("선택된 항목의 데이터가 없거나 잘못 되었습니다.");
            }
        }

        /// <summary>
        /// 전체 리포트 메이킹 컨트롤 프로세스.
        /// </summary>
        private void makeReportProcess()
        {
            //f_Progress.updateProgress(1, getProgressStatusMsg(0));

            //makeReport_Personal();
            makeReport_P1();
            makeReport_P4();
            makeReport_P5();
            makeReport_P6();
            makeReport_P7();
            makeReport_P8();
            makeReport_P9();
            makeReport_P10();
            makeReport_P11();
            makeReport_P12();
            makeReport_P13();
            //f_Progress.updateProgress(99, getProgressStatusMsg(2));
        }


        private void makeReport_P1()
        {
            var rpt = Globals.P1;

            rpt.Shapes.Item("t_검사일").TextFrame2.TextRange.Text = commonData["검사일"];
            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["이름"];
        }

        private void makeReport_P4()
        {
            var rpt = Globals.P4;

            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["성명"];
            rpt.Shapes.Item("t_성명2").TextFrame2.TextRange.Text = commonData["성명"];
            rpt.Shapes.Item("t_성명3").TextFrame2.TextRange.Text = commonData["성명"];

            /// 강점과 약점 텍스트를 만든다.
            string cmtStrong = "";
            string cmtWeak = "";

            strongArea = resultData.Tables[0].AsEnumerable().Select(x => x["IQ_AREA"].ToString()).Take(3).ToList();
            weakArea = resultData.Tables[0].AsEnumerable().OrderBy(x => Convert.ToDecimal(x["SCR_PCT"])).ThenBy(x=> Convert.ToInt32(x["ORD"])).Select(x => x["IQ_AREA"].ToString()).Take(2).ToList();

            List<string> strongSubArea = new List<string>();
            List<string> weakSubArea = new List<string>();

            /// 강점지능의 각 서브영역에 대해 가장 강한 1개 (동점시 2개까지) 를 뽑아낸다.
            foreach (string iq_area in strongArea)
            {
                var subArea = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals(iq_area) && x["RNK"].ToString() == "1").Select(x => x["SUB_AREA"].ToString()).Take(2).ToList();
                if (subArea.Count > 0)
                    strongSubArea.Add(string.Join(",", subArea));
            }

            /// 약점지능의 각 서브영역에 대해 가장 약한 1개 (동점시 2개까지) 를 뽑아낸다.
            foreach (string iq_area in weakArea)
            {
                var subArea = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals(iq_area) && x["RNK_RVS"].ToString() == "1").Select(x => x["SUB_AREA"].ToString()).Take(2).ToList();
                if (subArea.Count > 0)
                    weakSubArea.Add(string.Join(",", subArea));
            }

            cmtStrong = string.Format("{0} 영역을 잘 하는 재능과 능력이 형성 발달 되어 있습니다.\r\n{1}에서는 {2} 능력이 높으며, {3}에서는 {4} 능력이 높습니다. {5}에서는 {6} 능력이 높게 형성발달 되어 있습니다.",
                                       string.Join(", ", strongArea), strongArea[0], strongSubArea[0], strongArea[1], strongSubArea[1], strongArea[2], strongSubArea[2]);

            cmtWeak = string.Format("{0} 영역에 대한 보완이 필요합니다.\r\n{1}에서는 {2}, {3}에서는 {4} 능력이 낮게 형성되어 어려움이 있을수 있습니다.",
                           string.Join(", ", weakArea), weakArea[0], weakSubArea[0], weakArea[1], weakSubArea[1]);


            rpt.Shapes.Item("t_강점").TextFrame2.TextRange.Text = cmtStrong;
            rpt.Shapes.Item("t_약점").TextFrame2.TextRange.Text = cmtWeak;
        }


        private void makeReport_P5()
        {
            var rpt = Globals.P5;

            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["성명"];
            rpt.Shapes.Item("t_성명2").TextFrame2.TextRange.Text = commonData["성명"];

            List<decimal> result = resultData.Tables[0].AsEnumerable().OrderBy(x => Convert.ToInt32(x["DSPORD"])).Select(x => Convert.ToDecimal(x["SCR_PCT"])).ToList();
            if (result.Count == 8)
            {
                for (int iRow = 0 ; iRow < 8 ; iRow++)
                {
                    rpt.Cells[iRow + 11, 33].Value = result[iRow];
                }
            }
        }


        private void makeReport_P6()
        {
            var rpt = Globals.P6;

            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["성명"];
            rpt.Shapes.Item("t_성명2").TextFrame2.TextRange.Text = commonData["성명"];

            string areaFileNM = "";

            for (int i = 0 ; i < 3 ; i++)
            {
                areaFileNM = Path.Combine(Globals.ThisWorkbook.Main.G_SYSTEM_FOLDER, strongArea[i] + ".png");

                if (File.Exists(areaFileNM))
                {
                    change_Picture(Globals.ThisWorkbook.Sheets["6P"], rpt.Shapes.Item("p_강점" + (i + 1).ToString()), areaFileNM);
                }
            }

            for (int i = 0 ; i < 2 ; i++)
            {
                areaFileNM = Path.Combine(Globals.ThisWorkbook.Main.G_SYSTEM_FOLDER, weakArea[i] + ".png");

                if (File.Exists(areaFileNM))
                {
                    change_Picture(Globals.ThisWorkbook.Sheets["6P"], rpt.Shapes.Item("p_약점" + (i + 1).ToString()), areaFileNM);
                }
            }

        }

        private void makeReport_P7()
        {
            var rpt = Globals.P7;
            int putRow = 28;
             
            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["성명"];
            rpt.Shapes.Item("t_성명2").TextFrame2.TextRange.Text = commonData["성명"];

            List<string[]> strongShort = resultData.Tables[0].AsEnumerable().Select(x => new string[] { x["IQ_SHORT"].ToString(), x["IQ_AREA"].ToString() }).Take(2).ToList();
            List<string[]> weakShort = resultData.Tables[0].AsEnumerable().OrderBy(x => Convert.ToDecimal(x["SCR_PCT"])).ThenBy(x => Convert.ToInt32(x["ORD"])).Select(x => new string[] { x["IQ_SHORT"].ToString(), x["IQ_AREA"].ToString() }).Take(2).ToList();


            string areaFileNM = "";

            for (int i = 0 ; i < 2 ; i++)
            {
                areaFileNM = Path.Combine(Globals.ThisWorkbook.Main.G_SYSTEM_FOLDER, strongShort[i][0] + ".png");

                if (File.Exists(areaFileNM))
                {
                    change_Picture(Globals.ThisWorkbook.Sheets["7P"], rpt.Shapes.Item("p_유형" + (i + 1).ToString()), areaFileNM);
                }

                var subResult = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals(strongShort[i][1])).OrderBy(x => Convert.ToInt32(x["ORD"]));
                foreach (DataRow r in subResult)
                {
                    rpt.Cells[putRow, 32] = r["SUB_AREA"].ToString();
                    rpt.Cells[putRow, 33] = r["SCR_PCT"];
                    putRow += 1;
                }

                putRow += 1;
            }

            for (int i = 0 ; i < 2 ; i++)
            {
                areaFileNM = Path.Combine(Globals.ThisWorkbook.Main.G_SYSTEM_FOLDER, weakShort[i][0] + ".png");

                if (File.Exists(areaFileNM))
                {
                    change_Picture(Globals.ThisWorkbook.Sheets["7P"], rpt.Shapes.Item("p_유형" + (i + 3).ToString()), areaFileNM);
                }

                var subResult = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals(weakShort[i][1])).OrderBy(x => Convert.ToInt32(x["ORD"]));
                foreach (DataRow r in subResult)
                {
                    rpt.Cells[putRow, 32] = r["SUB_AREA"].ToString();
                    rpt.Cells[putRow, 33] = r["SCR_PCT"];
                    putRow += 1;
                }

                putRow += 1;

            }


        }

        private void makeReport_P8()
        {
            var rpt = Globals.P8;

            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["성명"];

            int putR = 0, putC = 0;

            foreach (DataRow r in resultData.Tables[2].Rows)
            {
                putR = 0; putC = 0;

                switch (r["CLASSIFICATION"].ToString().ToLower())
                {
                    case "interest": putC = 33; break;
                    case "knowledge": putC = 34; break;
                    case "skill": putC = 35; break;
                }

                putR = Convert.ToInt32(r["DSPORD"]) + 11;

                if (putR > 0 && putC > 0)
                {
                    rpt.Cells[putR, putC].Value = r["SCR_PCT"];
                }
            }
        }


        private void makeReport_P9()
        {
            var rpt = Globals.P9;

            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["성명"];

            int putR = 11;

            var rst = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals("언어지능")).OrderBy(x => Convert.ToInt32(x["ORD"])).ToList();
            if (rst.Count > 0)
            {
                foreach (DataRow r in rst)
                {
                    rpt.Cells[putR, 32].Value = r["SUB_AREA"].ToString();
                    rpt.Cells[putR, 33].Value = r["SCR_PCT"];
                    putR++;
                }
            }


            putR = 24;

            rst = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals("논리수학지능")).OrderBy(x => Convert.ToInt32(x["ORD"])).ToList();
            if (rst.Count > 0)
            {
                foreach (DataRow r in rst)
                {
                    rpt.Cells[putR, 32].Value = r["SUB_AREA"].ToString();
                    rpt.Cells[putR, 33].Value = r["SCR_PCT"];
                    putR++;
                }
            }
        }

        private void makeReport_P10()
        {
            var rpt = Globals.P10;

            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["성명"];

            int putR = 11;

            var rst = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals("음악지능")).OrderBy(x => Convert.ToInt32(x["ORD"])).ToList();
            if (rst.Count > 0)
            {
                foreach (DataRow r in rst)
                {
                    rpt.Cells[putR, 32].Value = r["SUB_AREA"].ToString();
                    rpt.Cells[putR, 33].Value = r["SCR_PCT"];
                    putR++;
                }
            }


            putR = 24;

            rst = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals("신체운동지능")).OrderBy(x => Convert.ToInt32(x["ORD"])).ToList();
            if (rst.Count > 0)
            {
                foreach (DataRow r in rst)
                {
                    rpt.Cells[putR, 32].Value = r["SUB_AREA"].ToString();
                    rpt.Cells[putR, 33].Value = r["SCR_PCT"];
                    putR++;
                }
            }
        }


        private void makeReport_P11()
        {
            var rpt = Globals.P11;

            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["성명"];

            int putR = 11;

            var rst = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals("공간지능")).OrderBy(x => Convert.ToInt32(x["ORD"])).ToList();
            if (rst.Count > 0)
            {
                foreach (DataRow r in rst)
                {
                    rpt.Cells[putR, 32].Value = r["SUB_AREA"].ToString();
                    rpt.Cells[putR, 33].Value = r["SCR_PCT"];
                    putR++;
                }
            }


            putR = 24;

            rst = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals("자연지능")).OrderBy(x => Convert.ToInt32(x["ORD"])).ToList();
            if (rst.Count > 0)
            {
                foreach (DataRow r in rst)
                {
                    rpt.Cells[putR, 32].Value = r["SUB_AREA"].ToString();
                    rpt.Cells[putR, 33].Value = r["SCR_PCT"];
                    putR++;
                }
            }
        }


        private void makeReport_P12()
        {
            var rpt = Globals.P12;

            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["성명"];

            int putR = 11;

            var rst = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals("대인지능")).OrderBy(x => Convert.ToInt32(x["ORD"])).ToList();
            if (rst.Count > 0)
            {
                foreach (DataRow r in rst)
                {
                    rpt.Cells[putR, 32].Value = r["SUB_AREA"].ToString();
                    rpt.Cells[putR, 33].Value = r["SCR_PCT"];
                    putR++;
                }
            }


            putR = 25;

            rst = resultData.Tables[1].AsEnumerable().Where(x => x["IQ_AREA"].ToString().Equals("개인내지능")).OrderBy(x => Convert.ToInt32(x["ORD"])).ToList();
            if (rst.Count > 0)
            {
                foreach (DataRow r in rst)
                {
                    rpt.Cells[putR, 32].Value = r["SUB_AREA"].ToString();
                    rpt.Cells[putR, 33].Value = r["SCR_PCT"];
                    putR++;
                }
            }
        }


        private void makeReport_P13()
        {
            var rpt = Globals.P13;
            rpt.Shapes.Item("t_성명").TextFrame2.TextRange.Text = commonData["성명"];

            var rstOrder = resultData.Tables[0].AsEnumerable().OrderBy(x => Convert.ToInt32(x["DSPORD"])).ToList();

            int putR = 31;
            foreach (DataRow r in rstOrder)
            {
                rpt.Cells[putR, 33].Value = r["SCR_PCT"];
                rpt.Cells[putR - 14, 33].Value = r["SCR_PCT"];
                putR++;
            }
        }




        #region clear report

        public static void clear_Result(bool complete = false)
        {
            Globals.P1.Shapes.Item("t_검사일").TextFrame2.TextRange.Text = "";
            Globals.P1.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            
            if (complete)
            {
                //Globals.P1.Shapes.Item("t_검사단체").TextFrame2.TextRange.Text = "";
                //Globals.P1.Shapes.Item("t_검사기관").TextFrame2.TextRange.Text = "";
            }

            Globals.P4.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            Globals.P4.Shapes.Item("t_성명2").TextFrame2.TextRange.Text = "";
            Globals.P4.Shapes.Item("t_성명3").TextFrame2.TextRange.Text = "";
            Globals.P4.Shapes.Item("t_약점").TextFrame2.TextRange.Text = "";
            Globals.P4.Shapes.Item("t_강점").TextFrame2.TextRange.Text = "";

            Globals.P5.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            Globals.P5.Shapes.Item("t_성명2").TextFrame2.TextRange.Text = "";
            Globals.P5.Cells[11, 33].Resize[8, 1].ClearContents();

            Globals.P6.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            Globals.P6.Shapes.Item("t_성명2").TextFrame2.TextRange.Text = "";

            Globals.P7.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            Globals.P7.Shapes.Item("t_성명2").TextFrame2.TextRange.Text = "";
            Globals.P7.Cells[28, 32].Resize[19, 2].ClearContents();

            Globals.P8.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            Globals.P8.Cells[12, 33].Resize[8, 3].ClearContents();

            Globals.P9.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            Globals.P9.Cells[11, 33].Resize[30, 1].ClearContents();

            Globals.P10.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            Globals.P10.Cells[11, 33].Resize[30, 1].ClearContents();

            Globals.P11.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            Globals.P11.Cells[11, 33].Resize[30, 1].ClearContents();

            Globals.P12.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            Globals.P12.Cells[11, 33].Resize[30, 1].ClearContents();

            Globals.P13.Shapes.Item("t_성명").TextFrame2.TextRange.Text = "";
            Globals.P13.Cells[17, 33].Resize[25, 1].ClearContents();

        }



        #endregion






        private DataTable getReportTable(string p)
        {
            DataTable rtn = new DataTable();

            switch (p)
            {
                case "Summary":
                case "Summary2":
                    rtn.Columns.Add("No", typeof(Int32));
                    rtn.Columns.Add("학교코드", typeof(string));
                    rtn.Columns.Add("학교명", typeof(string));
                    rtn.Columns.Add("학번", typeof(string));
                    rtn.Columns.Add("이름", typeof(string));
                    rtn.Columns.Add("학년", typeof(string));
                    rtn.Columns.Add("총점_점수", typeof(decimal));
                    rtn.Columns.Add("총점_환산점수", typeof(decimal));
                    rtn.Columns.Add("총점_T점수", typeof(decimal));
                    rtn.Columns.Add("총점_학교석차", typeof(Int32));
                    rtn.Columns.Add("총점_전체석차", typeof(Int32));
                    rtn.Columns.Add("총점_석차백분위", typeof(decimal));

                    for (int iSbj = 1 ; iSbj <= 8 ; iSbj++)
                    {
                        rtn.Columns.Add(string.Format("과목{0}_점수", iSbj), typeof(decimal));
                        rtn.Columns.Add(string.Format("과목{0}_환산점수", iSbj), typeof(decimal));
                        rtn.Columns.Add(string.Format("과목{0}_학교석차", iSbj), typeof(Int32));
                        rtn.Columns.Add(string.Format("과목{0}_전체석차", iSbj), typeof(Int32));
                        rtn.Columns.Add(string.Format("과목{0}_석차백분위", iSbj), typeof(decimal));
                    }

                    break;
            }

            return rtn;
        }

        private string getProgressStatusMsg(int i)
        {
            string rtn = "";

            switch (i)
            {
                case 0: rtn = "리포트 작성중입니다. 잠시 기다려 주세요. [P1]";
                    break;
                case 1: rtn = "리포트 작성중입니다. 잠시 기다려 주세요. [성적분포도]";
                    break;
                case 2: rtn = "리포트 작성이 완료되었습니다.";
                    break;
                case 3: rtn = "리포트 작성중입니다. 잠시 기다려 주세요. [문항분석표]";
                    break;
                case 4: rtn = "리포트 작성중입니다. 잠시 기다려 주세요. [대학석차]";
                    break;

            }
            return rtn;
        }

        private string getAnswerCharacter(string p)
        {
            string rtn = "";
            switch (p)
            {
                case "1": rtn = "1"; break;
                case "2": rtn = "2"; break;
                case "3": rtn = "3"; break;
                case "4": rtn = "4"; break;
                case "5": rtn = "5"; break;
                case "6": rtn = "6"; break;
                case "7": rtn = "7"; break;
                case "8": rtn = "8"; break;
                case "9": rtn = "9"; break;
                case "10": rtn = "10"; break;
                case "11": rtn = "11"; break;
                case "12": rtn = "12"; break;
                case "13": rtn = "13"; break;
                case "14": rtn = "14"; break;
                case "15": rtn = "15"; break;
            }
            return rtn;
        }

        /// <summary>
        /// OMR 마킹으로부터 TYPEID를 받아온다. DB에 설정해 두었지만 귀찮으니까 그냥 하드코딩으로 처리해 리턴한다.
        /// </summary>
        private int getTypeID(string value)
        {
            int rtn = 0;

            switch (value)
            {
                case "1": rtn = 1; break;
                case "2": rtn = 2; break;
                case "3": rtn = 3; break;
                case "4": rtn = 3; break;
                case "5": rtn = 4; break;
            }

            return rtn;
        }


        private string getDisplayName(string v1, string v2)
        {
            string alias = "";

            switch (v2)
            {
                case "1":
                case "유아 INFANT":
                case "5":
                case "고3/성인 SENIOR":
                    alias = "님"; break;

                case "2":
                case "3":
                case "4":
                case "어린이 CHILD":
                case "아동/청소년 JUNIOR":
                    alias = "학생"; break;
            }

            return string.Format("{0} {1}", v1, alias);
        }

        private void change_Picture(XL.Worksheet sht, XL.Shape shp, string targetSrc)
        {
            string shpName = shp.Name;
            float l, t, w, h;
            l = shp.Left;
            t = shp.Top;
            w = shp.Width;
            h = shp.Height;
            shp.Delete();

            XL.Shape newShp = sht.Shapes.AddPicture(targetSrc, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, l, t, w, h);
            newShp.Name = shpName;
        }

        #endregion


        #endregion


    }
}
