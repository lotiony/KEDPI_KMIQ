using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using KMIQ;
using loCommon;
using XL = Microsoft.Office.Interop.Excel;

namespace KMIQ.Form
{
    /// <summary>
    /// frmPrint.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class frmPrint : Window
    {
        #region common variables

        DataTable rptList;
        DataTable stuList;
        BackgroundWorker bg_Print;

        const string REPORT_Result = "P1";
        const string REPORT_Distribution = "성적분포도";
        const string REPORT_QAnalysis = "문항분석표";
        //const string REPORT_TypeAnalysis = "Type Analysis";
        const string REPORT_Personal = "개인성적표";
        #endregion

        public frmPrint()
        {
            InitializeComponent();
            TitleArea.MouseLeftButtonDown += (o, e) => { try { DragMove(); } catch { } };

            form_Init();
        }

        

        #region form event handler

        private void Image_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            this.Close();
        }

        private void dg_Report_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (dg_Report.SelectedItems != null)
            {
                if (dg_Report.SelectedItems.Count > 1)
                {
                    foreach (DataRowView it in dg_Report.SelectedItems)
                    {
                        it.Row["IsSelected"] = !(bool)it.Row["IsSelected"];
                    }
                    Check_RptAllSelected();
                }
                else
                {
                    Check_RptAllSelected();
                }
            }
        }

        private void chk_All_Report_Click(object sender, RoutedEventArgs e)
        {
            rptList.AsEnumerable().ToList().ForEach(x => x["IsSelected"] = chk_All_Report.IsChecked);
            dg_Report.UpdateLayout();
        }

        private void Rpt_CheckBox_Click(object sender, RoutedEventArgs e)
        {
            Check_RptAllSelected();
        }

        private void dg_Student_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (dg_Student.SelectedItems != null)
            {
                if (dg_Student.SelectedItems.Count > 1)
                {
                    foreach (DataRowView it in dg_Student.SelectedItems)
                    {
                        it.Row["IsSelected"] = !(bool)it.Row["IsSelected"];
                    }
                    Check_StuAllSelected();
                }
                else
                {
                    Check_StuAllSelected();
                }
            }
        }

        private void chk_All_Student_Click(object sender, RoutedEventArgs e)
        {
            stuList.AsEnumerable().ToList().ForEach(x => x["IsSelected"] = chk_All_Student.IsChecked);
            dg_Student.UpdateLayout();
        }

        private void Stu_CheckBox_Click(object sender, RoutedEventArgs e)
        {
            Check_StuAllSelected();
        }

        private void tb_Search_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            Search();
        }

        private void btn_PrintSetup_Click(object sender, RoutedEventArgs e)
        {
            setPrintSetup();
        }

        private void btn_Print_Click(object sender, RoutedEventArgs e)
        {
            PrintProcess();
        }

        void bg_Print_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            printRunning = false;
            this.btn_Print.Content = "Print";
            Globals.P1.Select();

            if (success)
            {
                Progressbar_Update( Convert.ToInt32(progress.Maximum), "출력이 완료되었습니다.");
                System.Windows.Forms.MessageBox.Show("출력이 완료되었습니다.");
            }
            else
                Progressbar_Update(Convert.ToInt32(progress.Maximum), "출력 중 오류가 발생했습니다.");

            Progressbar_Visible(false);
        }

        void bg_Print_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Progressbar_Update(e.ProgressPercentage, (string)e.UserState);
        }

        void bg_Print_DoWork(object sender, DoWorkEventArgs e)
        {
            _PrintReport_Work(e.Argument);
        }

        private void chk_PDF_Checked(object sender, RoutedEventArgs e)
        {
            if ((bool)chk_PDF.IsChecked) chk_Excel.IsChecked = !chk_PDF.IsChecked;
        }

        private void chk_Excel_Checked(object sender, RoutedEventArgs e)
        {
            if ((bool)chk_Excel.IsChecked) chk_PDF.IsChecked = !chk_Excel.IsChecked;
        }

        #endregion



        #region private process

        private void form_Init()
        {
            Progressbar_Visible(false);

            /// 리포트 리스트 / 학생 리스트 셋팅
            rptList = new DataTable();
            stuList = new DataTable();
            
            rptList.Columns.Add("IsSelected", typeof(bool));
            rptList.Columns.Add("ReportType", typeof(string));

            rptList.Rows.Add(true, "개인성적표");
            //rptList.Rows.Add(true, Globals.P2.Name);

            stuList.Columns.Add("IsSelected", typeof(bool));
            stuList.Columns.Add("No", typeof(int));
            stuList.Columns.Add("성명", typeof(string));
            stuList.Columns.Add("생년월일", typeof(string));
            stuList.Columns.Add("검사일자", typeof(string));

            //for (int iRow = 0 ; iRow < Globals.ThisWorkbook.Main.MakeReport.rptCount ; iRow++)
            //{
            //    if (Globals.P1.Cells[Globals.ThisWorkbook.Main.Report1_Start_Row + iRow, 3].Value != "")
            //    {
            //        stuList.Rows.Add(true, iRow + 1, Globals.P1.Cells[Globals.ThisWorkbook.Main.Report1_Start_Row + iRow, 4].Value.ToString(), Globals.P1.Cells[Globals.ThisWorkbook.Main.Report1_Start_Row + iRow, 3].Value.ToString(), Globals.P1.Cells[Globals.ThisWorkbook.Main.Report1_Start_Row + iRow, 2].Value.ToString());
            //    }
            //}

            DataTable data = Globals.ThisWorkbook.Main.DataLoad.MainData;
            for (int iRow = 0 ; iRow < data.Rows.Count ; iRow++)
            {
                stuList.Rows.Add(true, data.Rows[iRow][0], data.Rows[iRow][1], data.Rows[iRow][3], data.Rows[iRow][5]);
            }

            dg_Report.DataContext = rptList.DefaultView;
            dg_Student.DataContext = stuList.DefaultView;
            chk_All_Report.IsChecked = true;
            chk_All_Student.IsChecked = true;

            /// 출력매수 컨트롤 셋팅
            cb_Copies.DataContext = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10,15,20,25,30,35,40,45,50,60,70,80,90,100,200 };
            cb_Copies.SelectedIndex = 0;

            /// 프로그레스 쪽 셋팅
            progress.Value = 0;
            progress.Maximum = 100;
            status.Text = "";

            /// 백그라운드워커 셋팅
            bg_Print = new BackgroundWorker();
            bg_Print.WorkerSupportsCancellation = true;
            bg_Print.WorkerReportsProgress = true;
            bg_Print.DoWork += bg_Print_DoWork;
            bg_Print.ProgressChanged += bg_Print_ProgressChanged;
            bg_Print.RunWorkerCompleted += bg_Print_RunWorkerCompleted;
        }

        private void Check_StuAllSelected()
        {
            if (stuList != null)
            {
                int selCount = 0;

                selCount = stuList.AsEnumerable().Where(x => (bool)x["IsSelected"]).Count();

                chk_All_Student.IsChecked = (selCount == dg_Student.Items.Count ? (bool)true : (bool)false);
            }
        }

        private void Check_RptAllSelected()
        {
            if (rptList != null)
            {
                int selCount = 0;

                selCount = rptList.AsEnumerable().Where(x => (bool)x["IsSelected"]).Count();

                chk_All_Report.IsChecked = (selCount == rptList.Rows.Count ? (bool)true : (bool)false);
                //= (selCount == rptList.Rows.Count ? (bool?)true : (selCount == 0 ? (bool?)false : (bool?)null));
            }
        }

        private void Search()
        {
            stuList.AsEnumerable().ToList().ForEach(x => x["IsSelected"] = false);

            stuList.DefaultView.RowFilter = string.Format("성명 Like '*{0}*' OR 생년월일 Like '*{0}*'", tb_Search.Text);
            dg_Student.DataContext = stuList.DefaultView;

            foreach (DataRowView it in dg_Student.Items)
            {
                it.Row["IsSelected"] = true;
            }
            Check_StuAllSelected();
        }

        private void setPrintSetup()
        {
            PrintDialog printdlg = new PrintDialog();
            printdlg.UseEXDialog = false;
            printdlg.ShowDialog();
            Properties.Settings.Default.Device_PrintName = printdlg.PrinterSettings.PrinterName;
            Properties.Settings.Default.Save();
        }



        /// 출력 관련 프로세스
        /// ------------------------------------------------------------------------------------
        #region Print Process

        private bool printRunning = false;
        private bool success = false;
        private System.Windows.Forms.Timer tm = new System.Windows.Forms.Timer();
        public struct pArgu
        {
            public int prtCount;
            public List<DataRowView> targetRows { get; set; }
            public List<string> selectReport { get; set; }

            public bool chkedPDF;
            public bool chkedExcel;
            public string SavePrintFolder { get { return Path.Combine(Globals.ThisWorkbook.Main.G_PRINT_FOLDER, DateTime.Now.ToString("yyyy-MM-dd")); } }
            public string ExamName { get { return Globals.ThisWorkbook.Main.MakeReport.SelectedExam; } }
            public string SavePDFFolder { get { return Path.Combine(SavePrintFolder, "PDF"); } }
            public string SaveExcelFolder { get { return Path.Combine(SavePrintFolder, "Excel"); } }
            public string SavePersonalFolder;
            public int setCopies;
            public string[] examInfo;
            public bool IsCheckedPersonalReport_1 { get; set; }
            public bool IsCheckedPersonalReport_2 { get; set; }
            public bool IsCheckedPersonalReport_4 { get; set; }
        }

        ///// 파일 출력용 디렉토리 선택 프로시저
        //private void setOutputFolder()
        //{
        //    System.Windows.Forms.FolderBrowserDialog oFO = new System.Windows.Forms.FolderBrowserDialog();

        //    var _with1 = oFO;
        //    _with1.Description = "파일이 저장될 폴더를 선택해 주세요.";
        //    _with1.ShowNewFolderButton = true;

        //    if (oFO.ShowDialog() == DialogResult.OK)
        //    {
        //        Settings.Default.US_PRINT_FOLDER = oFO.SelectedPath;
        //        Settings.Default.Save();

        //        tb_SaveFolder.Text = oFO.SelectedPath + "\\";
        //    }
        //}


        /// 출력 제어 프로세스
        private void PrintProcess()
        {
            if (!printRunning)
            {
                printRunning = true;
                this.btn_Print.Content = "Stop";

                //***** 출력용으로 전달할 매개변수 그룹 셋팅
                pArgu pArg = new pArgu();
                pArg.targetRows = new List<DataRowView>();
                pArg.selectReport = new List<string>();
                pArg.chkedPDF = (bool)this.chk_PDF.IsChecked;
                pArg.chkedExcel = (bool)this.chk_Excel.IsChecked;
                pArg.setCopies = Convert.ToInt32(cb_Copies.Text);
                //pArg.IsCheckedPersonalReport_1 = (bool)this.chk_personal_1.IsChecked;
                //pArg.IsCheckedPersonalReport_2 = (bool)this.chk_personal_2.IsChecked;
                //pArg.IsCheckedPersonalReport_4 = (bool)this.chk_personal_4.IsChecked;


                // 선택된 학생 목록 DataRow 리스트
                foreach (DataRowView sR in dg_Student.Items)
                    if ((bool)sR["IsSelected"])
                        pArg.targetRows.Add(sR);


                // 선택된 리포트 목록 string 리스트. 각 리포트 선택상태에 따라 프린트 카운트 개수도 과목개수 포함해서 증가시켜준다.
                //foreach (DataRow ss in rptList.Rows)
                //    if ((bool)ss["IsSelected"])
                //    {
                //        pArg.selectReport.Add(ss["ReportType"].ToString());

                //        switch (ss["ReportType"].ToString())
                //        {
                //            case REPORT_Result:
                //            case REPORT_Distribution:
                //            //case REPORT_TypeAnalysis:
                //            case REPORT_QAnalysis:
                //                pArg.prtCount += 1;
                //                break;

                //            //case REPORT_QAnalysis:
                //            //    pArg.prtCount += Globals.ThisWorkbook.Main.SubjectCombo_Section.Split('|').Length;
                //            //    break;
                //        }
                //    }

                
                pArg.prtCount += pArg.targetRows.Count;

                progress.Maximum = pArg.prtCount;


                /// 개인리포트가 포함되어 있으면 상황에 맞게 하위폴더명 만들어준다
                if (pArg.targetRows.Count > 0)
                {
                    //pArg.SavePersonalFolder = string.Format("개인 리포트");
                    pArg.SavePersonalFolder = "";
                }

                /// ***** 리포트 생성 전 선택된 항목을 점검
                //if (pArg.selectReport.Count == 0 && pArg.prtCount == 0)     // 선택된 항목 체크
                if (pArg.prtCount == 0)     // 선택된 항목 체크
                {
                    System.Windows.Forms.MessageBox.Show("출력할 리포트가 하나도 선택되지 않았습니다.");
                    PrintCancel();
                    return;
                }

                try
                {
                    //***** 폴더 없으면 폴더 생성
                    if (!Directory.Exists(pArg.SavePDFFolder) & pArg.chkedPDF)
                        Directory.CreateDirectory(pArg.SavePDFFolder);

                    //***** 폴더 없으면 폴더 생성
                    if (!Directory.Exists(pArg.SaveExcelFolder) & pArg.chkedExcel)
                        Directory.CreateDirectory(pArg.SaveExcelFolder);

                    //if (pArg.SavePersonalFolder != "") // 개별폴더도 생성
                    //    if (pArg.chkedPDF) Directory.CreateDirectory(Path.Combine(pArg.SavePDFFolder, pArg.SavePersonalFolder));
                    //    else if (pArg.chkedExcel) Directory.CreateDirectory(Path.Combine(pArg.SaveExcelFolder, pArg.SavePersonalFolder));

                    Progressbar_Visible(true);
                    Progressbar_Update(0, "리포트 출력을 시작합니다.");
                }
                catch
                { }

                Control.CheckForIllegalCrossThreadCalls = false;
                bg_Print.RunWorkerAsync(pArg);

            }
            else PrintCancel();
        }


        private void PrintCancel()
        {
            this.btn_Print.Content = "Canceling...";
            this.Dispatcher.Invoke((ThreadStart)(() => { }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            printRunning = false;
            this.btn_Print.IsEnabled = true;
            bg_Print.CancelAsync();
            Progressbar_Visible(false);
            success = false;
        }


        /// <summary>
        ///  백그라운드 출력 프로세스
        /// </summary>
        private void _PrintReport_Work(object Aset)
        {
            pArgu ArguSet = (pArgu)Aset;

            try
            {

                string SaveFileNM = "";
                string pCaption = "";
                int PrintedPageNum = 1;


                ///// 선택 리포트 목록을 순회하면서 출력
                //for (int i = 0 ; i < ArguSet.selectReport.Count ; i++)
                //{
                //    /// 취소여부를 확인하고
                //    if (!bg_Print.CancellationPending)
                //    {
                //        XL.Worksheet sht;
                //        sht = Globals.ThisWorkbook.Sheets[ArguSet.selectReport[i]];

                //        if (sht != null)
                //        {
                //            /// 공통 리포트에 대한 리포트별 파일명 지정
                //            switch (ArguSet.selectReport[i])
                //            {
                //                case REPORT_Result:
                //                case REPORT_Distribution:
                //                //case REPORT_TypeAnalysis:
                //                case REPORT_QAnalysis:
                //                    SaveFileNM = ArguSet.selectReport[i];
                //                    pCaption = string.Format("출력중입니다... [{0}]  ({1} / {2})", ArguSet.selectReport[i], PrintedPageNum, ArguSet.prtCount);
                //                    bg_Print.ReportProgress(PrintedPageNum, pCaption);
                //                    printOut(ArguSet, sht, "", SaveFileNM, ref PrintedPageNum);
                //                    break;

                //                //case REPORT_Distribution:
                //                //    foreach (string sbj in Globals.ThisWorkbook.Main.SubjectCombo_Subject.Split('|'))
                //                //    {
                //                //        if (bg_Print.CancellationPending) break;

                //                //        Globals.ThisWorkbook.Main.MakeReport.Selected_Subject = sbj;
                //                //        Globals.ThisWorkbook.Main.MakeReport.makeReport_Distribution();

                //                //        SaveFileNM = string.Format("{0} ({1})", ArguSet.selectReport[i], sbj); 
                //                //        pCaption = string.Format("Now printing... [{0}-{1}]  ({2} / {3})", ArguSet.selectReport[i], sbj, PrintedPageNum, ArguSet.prtCount);
                //                //        bg_Print.ReportProgress(PrintedPageNum, pCaption);
                //                //        printOut(ArguSet, sht, "", SaveFileNM, ref PrintedPageNum);
                //                //    }
                //                //    break;

                //                //case REPORT_QAnalysis:
                //                //    foreach (string sbj in Globals.ThisWorkbook.Main.SubjectCombo_Section.Split('|'))
                //                //    {
                //                //        if (bg_Print.CancellationPending) break;

                //                //        Globals.ThisWorkbook.Main.MakeReport.Selected_Section = sbj;
                //                //        Globals.ThisWorkbook.Main.MakeReport.makeReport_QAnalysis_Select();

                //                //        SaveFileNM = string.Format("{0} ({1})", ArguSet.selectReport[i], sbj); 
                //                //        pCaption = string.Format("Now printing... [{0}-{1}]  ({2} / {3})", ArguSet.selectReport[i], sbj, PrintedPageNum, ArguSet.prtCount);
                //                //        bg_Print.ReportProgress(PrintedPageNum, pCaption);
                //                //        printOut(ArguSet, sht, "", SaveFileNM, ref PrintedPageNum);
                //                //    }
                //                //    break;
                //            }
                //        }
                //    }
                //}

                /// 학생 목록을 순회하면서 출력
                for (int i = 0 ; i < ArguSet.targetRows.Count ; i++)
                {
                    DataRowView tR = ArguSet.targetRows[i];

                    /// 취소여부를 확인하고
                    if (!bg_Print.CancellationPending)
                    {
                        string ID = string.Format("{0}.{1} ({2})", tR["No"], DBNull.Value.Equals(tR["성명"]) ? "" : tR["성명"], DBNull.Value.Equals(tR["생년월일"]) ? "" : tR["생년월일"]);

                        Globals.ThisWorkbook.Main.MakeReport.makeReport_Personal(ID);

                        //if (!ArguSet.IsCheckedPersonalReport_2) Globals.Sheet4.Cells[72, 1].Resize[76, 1].EntireRow.Hidden = true;
                        //if (!ArguSet.IsCheckedPersonalReport_4) Globals.Sheet4.Cells[148, 1].Resize[63, 1].EntireRow.Hidden = true;

                        SaveFileNM = string.Format("{0}_{1}({2})", tR["검사일자"], DBNull.Value.Equals(tR["성명"]) ? "" : tR["성명"], DBNull.Value.Equals(tR["생년월일"]) ? "" : tR["생년월일"]);
                        pCaption = string.Format("출력중입니다... [{0}]  ({1} / {2})", ID, PrintedPageNum, ArguSet.prtCount);
                        bg_Print.ReportProgress(PrintedPageNum, pCaption);
                        printOut(ArguSet, ArguSet.SavePersonalFolder, SaveFileNM, ref PrintedPageNum);
                    }
                }
                success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                success = false;
            }
            finally
            {
                
            }
        }



        private void printOut(pArgu arg, string personalFolder, string saveFileNM, ref int PrintedPageNum)
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            if (arg.chkedPDF)
            {
            Retry1:
                try
                {
                    Globals.ThisWorkbook.Sheets[Globals.ThisWorkbook.Main.offlineOutputShts].Select();
                    Globals.ThisWorkbook.Application.ActiveSheet.ExportAsFixedFormat(type: XL.XlFixedFormatType.xlTypePDF, filename: Path.Combine(arg.SavePDFFolder, personalFolder, saveFileNM + ".pdf"), quality: XL.XlFixedFormatQuality.xlQualityStandard, includeDocProperties: true, ignorePrintAreas: false, openAfterPublish: false);
                    Globals.P1.Select();
                }
                catch (Exception ex)
                {
                    loFunctions.LogWrite(string.Format("{0}  pdf 프린트 실패. 재시도 합니다. {1}", saveFileNM, ex.Message));
                    Globals.P1.Select();
                    int waitTime = 0;
                    while (waitTime < 10)
                    {
                        Thread.Sleep(100);
                        System.Windows.Forms.Application.DoEvents();
                        waitTime += 1;
                    }
                    if (bg_Print.CancellationPending)
                    {
                        Globals.ThisWorkbook.Application.ScreenUpdating = true;
                        return;
                    }
                    goto Retry1;
                }
            }
            else if (arg.chkedExcel)
            {
                SaveXlReport(Path.Combine(arg.SaveExcelFolder, personalFolder, saveFileNM + ".xlsx"));
            }
            else
            {
                Globals.ThisWorkbook.Sheets[Globals.ThisWorkbook.Main.offlineOutputShts].Select();
                Globals.ThisWorkbook.Application.ActiveWindow.SelectedSheets.PrintOutEx( Copies: (short)arg.setCopies, ActivePrinter: Properties.Settings.Default.Device_PrintName);
                Globals.P1.Select();
            }

            Globals.ThisWorkbook.Application.ScreenUpdating = true;
            PrintedPageNum += 1;
        }

        public void Progressbar_Visible(bool Vis)
        {
            progress.Visibility = Vis ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            status.Visibility = Vis ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
        }
        public void Progressbar_Update(int PCT, string pCaption = "", int pCnt = 0, int pAllCnt = 0)
        {
            progress.Value = PCT;
            progress.UpdateLayout();

            status.Text = pCaption;
            System.Windows.Forms.Application.DoEvents();
        }


        private void SaveXlReport(string SaveFNM)
        {
            XL.Workbook xWB = null;
            XL.Application xApp = Globals.ThisWorkbook.Application;
            xApp.DisplayAlerts = false;

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


                Globals.P1.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P2.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P3.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P4.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P5.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P6.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P7.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P8.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P9.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P10.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P11.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P12.Copy(before: xWB.Sheets["Sheet1"]);
                Globals.P13.Copy(before: xWB.Sheets["Sheet1"]);

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
                SaveFNM = SaveFNM.Replace("\\\\", "\\");
                xWB.SaveAs(SaveFNM, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XL.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

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
        }

        #endregion





        #endregion












    }
}
