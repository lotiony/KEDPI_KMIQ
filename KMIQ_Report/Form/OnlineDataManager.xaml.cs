using loCommon;
using System;
using System.ComponentModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Linq;
using System.Collections.Generic;
using XL = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.IO;
using System.Data.SqlClient;

namespace KMIQ.Form
{
    /// <summary>
    /// OnlineDataManager.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class OnlineDataManager : Window
    {
        BackgroundWorker bg_Worker;
        DataTable onlineDataTable;
        DataTable cbYearDataTable;
        int selectcount = 0;
        private bool printRunning = false;
        private bool success = false;
        WebApi webApi;

        System.Windows.Forms.Timer tm_autoPublish;
        bool RunningAutoPublish = false;
        int tm_Second;
        const int tm_Gap = 30;

        public OnlineDataManager()
        {
            InitializeComponent();
            this.TitleArea.MouseLeftButtonDown += (o, e) => { try { DragMove(); } catch { } };

            form_Init();
        }


        #region form event handler

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try { webApi?.Dispose(); } catch { }
            this.Close();
        }

        private void OnlineDataList_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (OnlineDataList.SelectedItems != null)
            {
                if (OnlineDataList.SelectedItems.Count > 1)
                {
                    foreach (DataRowView it in OnlineDataList.SelectedItems)
                    {
                        it.Row["IsSelected"] = !(bool)it.Row["IsSelected"];
                    }
                    Check_AllSelected();
                }
                else
                {
                    Check_AllSelected();
                }
            }
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            getOnlineDataList();
        }

        private void cbYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            search();
        }

        private void tbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            search();
        }

        private void cbxAll_Click(object sender, RoutedEventArgs e)
        {
            onlineDataTable.AsEnumerable().ToList().ForEach(x => x["IsSelected"] = this.cbxAll.IsChecked);
            this.OnlineDataList.UpdateLayout();
        }

        private void cbCheck_Click(object sender, RoutedEventArgs e)
        {
            Check_AllSelected();
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            runPublish();
        }

        private void btnAutoPublish_Click(object sender, RoutedEventArgs e)
        {
            AutoPublish();
        }

        #endregion



        #region private process

        private void form_Init()
        {
            Progressbar_Visible(false);

            cbYearDataTable = new DataTable();
            cbYearDataTable.Columns.Add("Year", typeof(string));

            webApi = new WebApi();
            webApi.CompleteSendResult += WebApi_CompleteSendResult;

            /// 온라인 데이터 로드
            getOnlineDataList();

            bg_Worker = new BackgroundWorker();
            bg_Worker.WorkerReportsProgress = true;
            bg_Worker.WorkerSupportsCancellation = true;
            bg_Worker.ProgressChanged += Bg_Worker_ProgressChanged;
            bg_Worker.RunWorkerCompleted += Bg_Worker_RunWorkerCompleted;
            bg_Worker.DoWork += bg_Worker_DoWork;

            this.cbxAll.IsChecked = false;
            this.btnStart.Content = "리포트 발행";

            /// 프로그레스 쪽 셋팅
            progress.Value = 0;
            progress.Maximum = 100;
            status.Content = "";


            /// 자동발행 셋팅
            tm_autoPublish = new System.Windows.Forms.Timer();
            tm_autoPublish.Interval = 1000;
            tm_Second = tm_Gap;
            tm_autoPublish.Tick += Tm_autoPublish_Tick;

        }

        private void Check_AllSelected()
        {
            if (onlineDataTable != null)
            {
                int selCount = 0;

                selCount = onlineDataTable.AsEnumerable().Where(x => (bool)x["IsSelected"]).Count();

                this.cbxAll.IsChecked = (selCount == OnlineDataList.Items.Count ? (bool)true : (bool)false);
            }
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

            status.Content = pCaption;
            System.Windows.Forms.Application.DoEvents();
        }


        #endregion




        #region search process

        /// <summary>
        /// Connecting ms-sql db and Get All service form data. Getting the unioned common items and logined user personal items.
        /// All data is stored to datatable named ServiceFormTable.
        /// </summary>
        private void getOnlineDataList()
        {
            using (SqlConn sql = new SqlConn())
            {
                try
                {
                    string qry = string.Format(" EXEC [dbo].[USP_GET_ONLINEDATALIST]");

                    onlineDataTable = sql.SqlSelect(qry.ToString());
                    onlineDataTable.Columns.Add("IsSelected", typeof(Boolean));
                }
                catch (Exception ex)
                {
                    loFunctions.LogWrite("온라인 진단 데이터 로드 오류 : " + ex.Message);
                }
            }

            cbYearDataTable.Rows.Clear();
            cbYearDataTable.Rows.Add("전체");
            List<int> yearList = onlineDataTable.AsEnumerable().Select(x => Convert.ToDateTime(x["IN_DT"]).Year).Distinct().OrderBy(x => x).ToList<int>();

            if (yearList.Count > 0)
            {
                foreach (int year in yearList)
                {
                    cbYearDataTable.Rows.Add(year.ToString());
                }
            }
            this.cbYear.ItemsSource = cbYearDataTable.DefaultView;
            this.cbYear.DisplayMemberPath = "Year";
            this.cbYear.SelectedValuePath = "Year";
            this.cbYear.SelectedIndex = 0;

            search();
        }

        private void search()
        {
            onlineDataTable.AsEnumerable().ToList().ForEach(x => x["IsSelected"] = false);

            onlineDataTable.DefaultView.RowFilter = string.Format("(UID Like '*{0}*' OR NAME Like '*{0}*') {1}", this.tbSearch.Text, (this.cbYear.SelectedValue != null && this.cbYear.SelectedValue.ToString() != "전체") ? string.Format("AND (IN_DT >= #1/1/{0}# AND IN_DT <= #12/31/{0}#)", cbYear.SelectedValue) : "" );

            foreach (DataRowView it in onlineDataTable.DefaultView)
            {
                if (!Convert.ToBoolean(it.Row["IS_PUBLISHED"]))
                    it.Row["IsSelected"] = true;
            }

            this.OnlineDataList.DataContext = onlineDataTable.DefaultView;
        }


        #endregion





        #region Auto publish process

        private void AutoPublish()
        {
            if (!RunningAutoPublish)
            {
                RunningAutoPublish = true;
                EnableAllControls(false);
                this.btnAutoPublish.Content = "자동발행 중지";
                AutoPublishStatusText.Text = "자동발행이 시작되었습니다.";
                tm_Second = 0;
                tm_autoPublish.Start();
            }
            else
            {
                RunningAutoPublish = false;
                EnableAllControls(true);
                this.btnAutoPublish.Content = "자동발행 시작";
                AutoPublishStatusText.Text = "자동발행이 중지되었습니다.";
                tm_autoPublish.Stop();
                //ReportPrint();
            }
        }


        private void Tm_autoPublish_Tick(object sender, EventArgs e)
        {
            if (RunningAutoPublish)
            {
                AutoPublishStatusText.Text = string.Format("자동발행 실행중입니다. {0}초 후 갱신됩니다.", tm_Second);
                tm_Second--;

                if (tm_Second == -1)
                {
                    tm_Second = tm_Gap;
                    getOnlineDataList();
                    FixReportUpdateState();

                    if (checkAutoPublishTarget() > 0)
                    {
                        tm_autoPublish.Stop();
                        try
                        {
                            AutoPublishStatusText.Text = string.Format("리포트를 발행하고 있습니다. 잠시 기다려 주세요...");
                            this.Dispatcher.Invoke((ThreadStart)(() => { }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

                            ReportPrint();

                            while (bg_Worker.IsBusy)
                            {
                                Thread.Sleep(100);
                                this.Dispatcher.Invoke((ThreadStart)(() => { }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                                //Application.Current.Dispatcher.Invoke((ThreadStart)(() => { }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                            }
                        }
                        catch { }
                        tm_autoPublish.Start();
                    }
                }
            }
            else
            {
                tm_autoPublish.Stop();
            }
        }


        private int checkAutoPublishTarget()
        {
            int rtn = 0;

            if (onlineDataTable.DefaultView.Count > 0)
            {
                foreach (DataRowView it in onlineDataTable.DefaultView)
                {
                    if (!Convert.ToBoolean(it.Row["IS_PUBLISHED"]))
                    {
                        it.Row["IsSelected"] = true;
                        rtn++;
                    }
                }
            }
            return rtn;
        }

        
        private void EnableAllControls(bool _enabled)
        {
            this.btnStart.IsEnabled = _enabled;
            this.cbxAll.IsEnabled = _enabled;
            this.btnRefreshData.IsEnabled = _enabled;
            this.cbYear.IsEnabled = _enabled;
            this.tbSearch.IsEnabled = _enabled;
            this.OnlineDataList.IsEnabled = _enabled;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }


        #endregion



        #region Manual publish process

        private void runPublish()
        {
            ReportPrint();
        }

        /// <summary>
        /// printing SmartMarksheet paper : DB print. 
        /// </summary>
        private void ReportPrint()
        {
            if (bg_Worker.IsBusy)
            {
                PrintCancel();
            }
            else
            {
                this.printRunning = true;
                this.btnStart.Content = "발행 중지";

                //***** 출력용으로 전달할 매개변수 그룹 셋팅
                frmPrint.pArgu pArg = new frmPrint.pArgu();
                pArg.targetRows = new List<DataRowView>();
                pArg.selectReport = new List<string>();
                pArg.chkedPDF = true;
                pArg.setCopies = 1;

                // 선택된 데이터 목록 DataRow 리스트
                foreach (DataRowView sR in this.OnlineDataList.Items)
                    if ((bool)sR["IsSelected"])
                        pArg.targetRows.Insert(0, sR);


                pArg.prtCount += pArg.targetRows.Count;
                progress.Maximum = pArg.prtCount;

                /// ***** 리포트 생성 전 선택된 항목을 점검
                //if (pArg.selectReport.Count == 0 && pArg.prtCount == 0)     // 선택된 항목 체크
                if (pArg.prtCount == 0)     // 선택된 항목 체크
                {
                    System.Windows.Forms.MessageBox.Show("발행할 데이터가 선택되지 않았습니다.");
                    PrintCancel();
                    return;
                }


                Progressbar_Visible(true);
                Progressbar_Update(0, "리포트 출력을 시작합니다.");

                bg_Worker.RunWorkerAsync(pArg);
            }
        }

        private void PrintCancel()
        {
            this.btnStart.Content = "Canceling...";
            this.Dispatcher.Invoke((ThreadStart)(() => { }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

            printRunning = false;
            this.btnStart.IsEnabled = true;
            bg_Worker.CancelAsync();
            Progressbar_Visible(false);
            success = false;
        }

        void bg_Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            _PrintReport_Work(e.Argument);
        }

        /// <summary>
        ///  백그라운드 출력 프로세스
        /// </summary>
        private void _PrintReport_Work(object Aset)
        {
            frmPrint.pArgu ArguSet = (frmPrint.pArgu)Aset;

            try
            {
                string SaveFileNM = "";
                string pCaption = "";
                int PrintedPageNum = 1;

                /// 학생 목록을 순회하면서 출력
                for (int i = 0 ; i < ArguSet.targetRows.Count ; i++)
                {
                    DataRowView tR = ArguSet.targetRows[i];

                    /// 취소여부를 확인하고
                    if (!bg_Worker.CancellationPending)
                    {
                        string ID = string.Format("{0}({1})_{2}", DBNull.Value.Equals(tR["NAME"]) ? tR["UID"].ToString() : tR["NAME"].ToString(), tR["UID"].ToString(), Convert.ToDateTime(tR["IN_DT"]).ToString("yyyyMMdd") );

                        Globals.ThisWorkbook.Main.MakeReport.makeReport_Online(tR.Row);

                        SaveFileNM = string.Format("{0}_결과리포트", ID);
                        pCaption = string.Format("발행 및 게시중입니다... [{0}]  ({1} / {2})", ID, PrintedPageNum, ArguSet.prtCount);
                        bg_Worker.ReportProgress(PrintedPageNum, pCaption);
                        printOut(ArguSet, tR.Row, Globals.ThisWorkbook.Main.G_TMP_FOLDER, SaveFileNM, ref PrintedPageNum);

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

        private void printOut(frmPrint.pArgu arg, DataRow tR, string personalFolder, string saveFileNM, ref int PrintedPageNum)
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            if (arg.chkedPDF)
            {
                Retry1:
                try
                {
                    string pdfFileName = Path.Combine(personalFolder, saveFileNM + ".pdf");
                    Globals.ThisWorkbook.ExportAsFixedFormat(type: XL.XlFixedFormatType.xlTypePDF, filename: pdfFileName, quality: XL.XlFixedFormatQuality.xlQualityStandard, includeDocProperties: true, ignorePrintAreas: false, openAfterPublish: false);
                    Globals.P1.Select();

                    UpdatePublishState(tR, pdfFileName);
                }
                catch (Exception ex)
                {
                    loFunctions.LogWrite(string.Format("{0}  pdf 프린트 실패. 재시도 합니다. {1}", saveFileNM, ex.Message));
                    int waitTime = 0;
                    while (waitTime < 10)
                    {
                        Thread.Sleep(100);
                        System.Windows.Forms.Application.DoEvents();
                        waitTime += 1;
                    }
                    if (bg_Worker.CancellationPending)
                    {
                        Globals.ThisWorkbook.Application.ScreenUpdating = true;
                        return;
                    }
                    goto Retry1;
                }
            }

            Globals.ThisWorkbook.Application.ScreenUpdating = true;
            PrintedPageNum += 1;
        }

        private void Bg_Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            printRunning = false;
            Globals.P1.Select();

            if (success)
            {
                Progressbar_Update(Convert.ToInt32(progress.Maximum), "리포트 발행 및 게시가 완료되었습니다.");

                if (!RunningAutoPublish)
                    System.Windows.Forms.MessageBox.Show("리포트 발행 및 게시가 완료되었습니다.");
            }
            else
                Progressbar_Update(Convert.ToInt32(progress.Maximum), "리포트 발행 중 오류가 발생했습니다.");


            this.btnStart.Content = "리포트 발행";
            Progressbar_Visible(false);
        }

        private void Bg_Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Progressbar_Update(e.ProgressPercentage, (string)e.UserState);
            this.Dispatcher.Invoke((ThreadStart)(() => { }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }



        #region Update publish state

        private void UpdatePublishState(DataRow item, string pdfFileName)
        {
            item["IS_PUBLISHED"] = false;
            item["DSP_PUB"] = "X";

            /// DB의 해당 ODID에도 발행여부를 갱신하고 pdf파일을 업로드 해준다.
            if (File.Exists(pdfFileName))
            {
                using (SqlConnection SqlCon = new SqlConnection(Properties.Settings.Default.ConnectionString))
                {
                    try
                    {
                        SqlCon.Open();

                        if (SqlCon.State == ConnectionState.Open)
                        {
                            using (SqlCommand SqlCmd = SqlCon.CreateCommand())
                            {
                                SqlCmd.Connection = SqlCon;
                                SqlCmd.CommandType = CommandType.StoredProcedure;

                                SqlCmd.CommandText = "dbo.[USP_P_PUBLISH_FILE_UPLOAD]";

                                SqlCmd.Parameters.Add("@ODID", SqlDbType.Int);
                                SqlCmd.Parameters.Add("@FILENAME", SqlDbType.NVarChar, 300);
                                SqlCmd.Parameters.Add("@FILESTREAM", SqlDbType.VarBinary, Int32.MaxValue - 1);

                                try
                                {
                                    byte[] fileBytes = File.ReadAllBytes(pdfFileName);

                                    SqlCmd.Parameters["@ODID"].Value = item["ODID"];
                                    SqlCmd.Parameters["@FILENAME"].Value = Path.GetFileName(pdfFileName);
                                    SqlCmd.Parameters["@FILESTREAM"].Value = fileBytes;

                                    SqlCmd.ExecuteNonQuery();

                                    item["IS_PUBLISHED"] = true;
                                    item["DSP_PUB"] = "O";

                                    webApi.SendResult(item["TOKEN"].ToString());
                                }
                                catch (Exception ex)
                                {
                                    loFunctions.LogWrite("Data Uploading error : " + ex.Message);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        loFunctions.LogWrite("[UpdatePublishState - USP_P_PUBLISH_FILE_UPLOAD] on Error : " + ex.Message);
                    }
                }


                try
                {
                    File.Delete(pdfFileName);
                }
                catch (Exception)
                {
                }
            }
        }

        /// <summary>
        /// api에서 SendResult의 결과가 날아오면 해당 Guid값으로 갱신결과를 업데이트 해준다.
        /// </summary>
        private void WebApi_CompleteSendResult(string guid, bool success)
        {
            if (success)
            {
                using (SqlConn sql = new SqlConn())
                {
                    try
                    {
                        string qry = $" UPDATE T_ONLINE_DATA SET IS_REPORT_UPDATE = 1 WHERE TOKEN = '{guid.ToUpper()}'";
                        sql.SqlExcute(qry);

                        DataRow item = onlineDataTable.AsEnumerable().Where(x => x["TOKEN"].ToString() == guid.ToUpper()).FirstOrDefault();
                        if (item != null)
                        {
                            item["DSP_UPDATE"] = "O";
                        }
                    }
                    catch (Exception ex)
                    {
                        loFunctions.LogWrite("웹 갱신결과 업데이트 오류 : " + ex.Message);
                    }
                }

            }
        }

        /// <summary>
        /// IS_PUBLISH가 1인데 IS_REPORT_UPDATE가 0인 항목들 : 왜인지 webApi에 SendResult가 제대로 작동 안한 놈들 모두 찾아서 일괄 갱신 처리한다.
        /// </summary>
        private void FixReportUpdateState()
        {
            var fixTargets = onlineDataTable.AsEnumerable().Where(x => x["DSP_PUB"].ToString() == "O" && x["DSP_UPDATE"].ToString() == "X").ToList();

            if (fixTargets.Count > 0)
            {
                fixTargets.ForEach(x=> {
                    webApi.SendResult(x["TOKEN"].ToString());
                    Console.WriteLine($"웹리포트상태 강제갱신 : {x["TOKEN"].ToString()}");
                });
            }
        }


        #endregion

        #endregion


    }

}
