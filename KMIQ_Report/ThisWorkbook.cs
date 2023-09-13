using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.IO;

namespace KMIQ
{
    public partial class ThisWorkbook
    {
        public Main Main;

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            InitializeProgram();
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                Globals.ThisWorkbook.Application.DisplayAlerts = true;
            }
            catch { }
        }

        private void ThisWorkbook_BeforeClose(ref bool Cancel)
        {
            if (Directory.Exists(Globals.ThisWorkbook.Main.G_TMP_FOLDER)) Directory.Delete(Globals.ThisWorkbook.Main.G_TMP_FOLDER, true);
            Globals.ThisWorkbook.Close(false);
        }

        private void ThisWorkbook_BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            //MessageBox.Show("리포트 저장을 허용하지 않습니다.");
            //Cancel = true;
        }

        /// <summary>
        /// 시트 선택에 따라 리본메뉴 컨트롤.
        /// </summary>
        private void ThisWorkbook_SheetActivate(object Sh)
        {
            Microsoft.Office.Interop.Excel.Worksheet sht = (Microsoft.Office.Interop.Excel.Worksheet)Sh;
        }


        #region VSTO 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InternalStartup()
        {
            this.BeforeClose += new Microsoft.Office.Interop.Excel.WorkbookEvents_BeforeCloseEventHandler(this.ThisWorkbook_BeforeClose);
            this.SheetActivate += new Microsoft.Office.Interop.Excel.WorkbookEvents_SheetActivateEventHandler(this.ThisWorkbook_SheetActivate);
            this.Startup += new System.EventHandler(this.ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(this.ThisWorkbook_Shutdown);

        }

        #endregion



        #region Sheet Process
        private void InitializeProgram()
        {
            Globals.ThisWorkbook.Application.ActiveWindow.DisplayWorkbookTabs = true;
            Globals.ThisWorkbook.Application.DisplayAlerts = false;
            Globals.P1.Activate();
            Globals.ThisWorkbook.Application.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized;

            Main = new Main();


            //if (!run_Auth())
            //{
            //    Globals.ThisWorkbook.Application.Quit();
            //    return;
            //}
        }

        private void beforeReleaseProgram()
        {
            Globals.ThisWorkbook.Application.ActiveWindow.DisplayWorkbookTabs = true;
        }


        public Microsoft.Office.Interop.Excel.Worksheet getSheetByName(string ShtName)
        {
            foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in this.Worksheets)
            {
                if (worksheet.Name == ShtName)
                    return worksheet;
            }
            throw new ArgumentException();
        }

        #endregion





        #region auth process

        //static bool run_Auth()
        //{
        //    if (!loCommon.loFunctions.checkRunningAuthorizer())
        //    {
        //        ///******인증을 처리한다.   --> 갈무리는 키가 없어도 사용이 가능하다(로그인 처리해서). 단, 판독은 무조건 키가 꼽혀야만 가능하도록 한다.
        //        {
        //            if (!loAuthCheck.loAuthCheck.MEGALOCK_Auth.Lock_Check())
        //            {
        //                MessageBox.Show("허용된 USB 인증키가 발견되지 않았습니다." + Environment.NewLine + "프로그램을 실행할 수 없습니다.", "인증 실패 - ⓒRealSMART");
        //                return false;
        //            }

        //            try
        //            {
        //                ///***** DB체크해서 정품인증을 한다.
        //                //System.Threading.Thread th1 = new System.Threading.Thread(_thrd_DBCheck);
        //                //th1.Start();
        //                if (_thrd_DBCheck())
        //                {
        //                    return true;
        //                    ///***** 아이피인증을 한다.
        //                    //return _thrd_IPCheck();
        //                }

        //                //System.Threading.Thread th2 = new System.Threading.Thread(_thrd_IPCheck);
        //                //th2.Start();

        //                return false;
        //            }
        //            catch
        //            { return false; }
        //        }
        //    }
        //    else
        //    {
        //        return true;
        //    }


        //}


        //public static bool run_UsbCheck()
        //{
        //    if (!loCommon.loFunctions.checkRunningAuthorizer())
        //    {
        //        ///****** 키가 꼽혀있는지 체크한다.
        //        {
        //            if (!loAuthCheck.loAuthCheck.MEGALOCK_Auth.Lock_Check())
        //            {
        //                MessageBox.Show("판독은 USB인증키가 꽂혀있는 상태에서만 사용 가능합니다.", "USB 인증키 찾기 실패 - ⓒRealSMART");
        //                return false;
        //            }
        //            else
        //            {
        //                return true;
        //            }
        //        }
        //    }
        //    else
        //    {
        //        return true;
        //    }
        //}

        //static bool _thrd_DBCheck()
        //{
        //    int AuthCount = 0;
        //    if (loAuthCheck.loAuthCheck.DB_AUTH.DB_Check(Properties.Settings.Default.AuthID, ref AuthCount) == false)
        //        return false;
        //    else
        //    {
        //        return true;
        //    }
        //}
        //static bool _thrd_IPCheck()
        //{
        //    return loAuthCheck.loAuthCheck.DB_AUTH.IP_Check(Properties.Settings.Default.AuthID);
        //}


        #endregion




    }
}
