using loCommon;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Interop;
using XL = Microsoft.Office.Interop.Excel;


namespace KMIQ
{
    public class cSaveReport
    {
        #region common variables

        public cMakeReport makeReport { get; set; }

        Form.frmProgress f_Progress;

        #endregion


        public cSaveReport(cMakeReport _makeReport)
        {
            makeReport = _makeReport;
        }


        #region public process

        public void SaveReport()
        {
            if (Globals.ThisWorkbook.Main.DataLoad.MainData == null) { MessageBox.Show(Globals.ThisWorkbook.Main.NoReport_Message); return; }


            ShowProgress();

            saveReportProcess();

            CloseProgress();

            Globals.P1.Activate();
            MessageBox.Show(Globals.ThisWorkbook.Main.SaveDB_Complete_Message);

        }

        #endregion




        #region private process

        private void ShowProgress()
        {
            f_Progress = new Form.frmProgress();
            new WindowInteropHelper(f_Progress).Owner = Globals.ThisWorkbook.Main.G_WINDOW_HANDLE;
            f_Progress.initializeProgress(16);
            f_Progress.Show();
        }

        private void CloseProgress()
        {
            f_Progress.Close();
        }




        #region save report process

        private bool printRunning = false;
        private bool success = false;
        private System.Windows.Forms.Timer tm = new System.Windows.Forms.Timer();
        public struct pArgu
        {
            public int prtCount;
            public List<string> selectReport { get; set; }

            public bool chkedExcel;
            public string SavePrintFolder { get { return Path.Combine(Globals.ThisWorkbook.Main.G_PRINT_FOLDER, Globals.ThisWorkbook.Main.MakeReport.SelectedExam); } }
            public string ExamName { get { return Globals.ThisWorkbook.Main.MakeReport.SelectedExam; } }
            public string SavePDFFolder { get { return Path.Combine(SavePrintFolder, "PDF"); } }
            public string SaveExcelFolder { get { return Path.Combine(SavePrintFolder, "Excel"); } }
            public string[] examInfo;
        }


        /// <summary>
        /// 전체 리포트 세이브 컨트롤 프로세스
        /// </summary>
        private void saveReportProcess()
        {
            //***** 출력용으로 전달할 매개변수 그룹 셋팅
            pArgu pArg = new pArgu();
            pArg.selectReport = new List<string>();
            pArg.chkedExcel = true;

            
            /// 리포트 카피 16개 + 각 학교별 저장에도 프로그레스를 업데이트 한다.
            f_Progress.initializeProgress(16);

            try
            {
                //***** 폴더 없으면 폴더 생성
                if (!Directory.Exists(pArg.SaveExcelFolder) & pArg.chkedExcel)
                    Directory.CreateDirectory(pArg.SaveExcelFolder);

                f_Progress.updateProgress(0, getProgressStatusMsg(0));
            }
            catch
            { }



            string rptFileNM = "_결과분석리포트.xlsx";
            string saveFileNM = Path.Combine(pArg.SaveExcelFolder, "전체" + rptFileNM);


            // 엑셀을 새로 만들고
            XL.Workbook xWB = null;
            XL.Application xApp = Globals.ThisWorkbook.Application;
            xApp.DisplayAlerts = false;
            xApp.ScreenUpdating = false;

            try
            {
                xWB = xApp.Workbooks.Add();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            try
            {
                string themeFile = xApp.Parent.Path.Replace("\\Office", "\\Document Themes ") + "\\Theme Colors\\Office 2007 - 2010.xml";
                //string themeFile = xApp.Parent.Path.Replace("\\Office", "\\Document Themes ") + "\\Theme Colors\\Blue.xml";
                //string themeFile2 = xApp.Parent.Path.Replace("\\Office", "\\Document Themes ") + "\\Theme Colors\\Flow.xml";

                if (File.Exists(themeFile))
                    xWB.Theme.ThemeColorScheme.Load(themeFile);
                //else if (File.Exists(themeFile2))
                //    xWB.Theme.ThemeColorScheme.Load(themeFile2);

                int newShtIndex = 1;
                List<XL.Worksheet> shts = new List<XL.Worksheet>();
                shts.Add(Globals.ThisWorkbook.Sheets[Globals.P1.Name]);

                for (byte dSht = 1 ; dSht <= xWB.Sheets.Count ; dSht++)
                {
                    try
                    {
                        xWB.Sheets["Sheet" + dSht.ToString()].delete();

                    }
                    catch (Exception ex)
                    { Console.WriteLine(ex.Message); }
                }
                //xlWB.Save()

                xWB.Sheets[1].Select();

                /// 종합(전체) 결과는 그대로 저장한다.
                saveFileNM = saveFileNM.Replace("\\\\", "\\");
                xWB.SaveAs(saveFileNM, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XL.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                loFunctions.LogWrite("save excel failed : " + ex.Message);
            }
            finally
            {
                xWB.Close(false);
                
            }

            try
            {
                xWB = null;
                //lf.releaseObject(xWB);
                //xApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            xApp.DisplayAlerts = true;
            xApp.ScreenUpdating = true;
        }

        #endregion



        private string getProgressStatusMsg(int i)
        {
            string rtn = "";

            switch (i)
            {
                case 0:
                    rtn = "리포트 저장을 시작합니다.";
                    break;
                case 1:
                    rtn = "리포트 저장중입니다. 잠시 기다려 주세요. [P1]";
                    break;
                case 2:
                    rtn = "리포트 저장중입니다. 잠시 기다려 주세요. [성적분포도-{0}]";
                    break;
                case 3:
                    rtn = "리포트 저장중입니다. 잠시 기다려 주세요. [문항분석표-{0}]";
                    break;
                case 4:
                    rtn = "리포트 저장중입니다. 잠시 기다려 주세요. [문항분석표-R형]";
                    break;
                case 5:
                    rtn = "리포트 저장중입니다. 잠시 기다려 주세요. [대학석차(분과별)]";
                    break;
                case 6:
                    rtn = "학교별 리포트 저장중입니다. 잠시 기다려 주세요. [{0}]";
                    break;

            }
            return rtn;
        }



        #endregion


    }
}
