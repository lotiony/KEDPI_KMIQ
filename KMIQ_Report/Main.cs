using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Deployment.Application;
using System.IO;

namespace KMIQ
{
    public class Main
    {
        #region global public variables

        public cDataLoad DataLoad;
        public cMakeReport MakeReport;
        public cSaveReport SaveReport;
        public string G_DATA_FOLDER { get { return (ApplicationDeployment.IsNetworkDeployed ? Path.Combine(Globals.ThisWorkbook.Path, "Data") : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data")); } }
        public string G_PRINT_FOLDER { get { return (ApplicationDeployment.IsNetworkDeployed ? Path.Combine(Globals.ThisWorkbook.Path, "Print") : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Print")); } }
        public string G_SYSTEM_FOLDER { get { return (ApplicationDeployment.IsNetworkDeployed ? Path.Combine(Globals.ThisWorkbook.Path, "SystemFiles") : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemFiles")); } }
        public string G_TMP_FOLDER { get { return (ApplicationDeployment.IsNetworkDeployed ? Path.Combine(Globals.ThisWorkbook.Path, "Temp") : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp")); } }
        public string G_DB_FILE { get { return Path.Combine(G_SYSTEM_FOLDER, "db.sdb"); } }
        public string G_EMPTYDB_FILE { get { return Path.Combine(G_SYSTEM_FOLDER, "empty.t"); } }
        public IntPtr G_WINDOW_HANDLE { get { return (IntPtr)Globals.ThisWorkbook.Application.Hwnd; } }

        #endregion

        #region global public const
        //public readonly string ExamCombo_Selector = "시험을 선택하세요.";
        public readonly string PersonCombo_Selector = "대상을 선택하세요.";
        public readonly string NoSelect_Message = "시험을 선택하세요.";
        public readonly string NoReport_Message = "불러온 데이터가 없습니다. 데이터를 먼저 불러오세요.";
        public readonly string Data_Name_Name = "Name";
        public readonly string Data_ID_Name = "학번";
        public readonly string Data_Class_Name = "학교명";

        public readonly int Data_Name_Index = 1;        // column header를 가져오지 못할경우 index로 해당 데이터를 찾는다.
        public readonly int Data_ID_Index = 5;
        public readonly int Data_Class_Index = 4;
        
        //public readonly string Data_Essay_Name = "Essay";
        public readonly int Data_Mark_Start_Index = 6;     // F
        public readonly int Data_Mark_End_Index = 155;     // EY

        public readonly C1ExcelArea Data_Area = new C1ExcelArea() { shtName = "data", startColIndex = 0, startRowIndex = 2, endColIndex = 155, endRowIndex = 0 };

        //public readonly string ClearDB_Confirm_Message = "데이터베이스를 초기화하시겠습니까?\r\n현재 상태의 DB는 백업됩니다.";
        //public readonly string ClearDB_Complete_Message = "데이터베이스 초기화가 완료되었습니다.";
        public readonly string ClearDB_Confirm_Message = "현재 불러온 데이터 파일을 언로드 하고 리포트를 클리어 합니다.";
        public readonly string ClearDB_Complete_Message = "리포트와 데이터가 초기화 되었습니다.";

        public readonly string UpdateDB_Complete_Message = "데이터 업로드가 완료되었습니다.";
        public readonly string UpdateDB_NoNeeds_Message = "업로드 할 데이터가 없습니다.";
        public readonly string UpdateDB_During_Message = "데이터 업로드 중입니다. 잠시 기다려 주세요.";
        public readonly string SaveDB_Complete_Message = "리포트 저장이 완료되었습니다.";

        public readonly string[] offlineOutputShts = { "1P", "2P", "3P", "4P", "5P", "6P", "7P", "8P", "9P", "10P", "11P", "12P", "13P" };

        #endregion


        public Main()
        {
            DataLoad = new cDataLoad();
            MakeReport = new cMakeReport();
            SaveReport = new cSaveReport(MakeReport);
            Console.WriteLine(Path.Combine(Globals.ThisWorkbook.Path, "Data"));

            if (!Directory.Exists(G_TMP_FOLDER)) Directory.CreateDirectory(G_TMP_FOLDER);
        }

    }


    public struct C1ExcelArea
    {
        public string shtName;
        public int startColIndex;
        public int startRowIndex;
        public int endColIndex;
        public int endRowIndex;
        
        public C1ExcelArea(string _shtName, int _startColIndex, int _startRowIndex, int _endColIndex, int _endRowIndex)
        {
            shtName = _shtName;
            startColIndex = _startColIndex;
            startRowIndex = _startRowIndex;
            endColIndex = _endColIndex;
            endRowIndex = _endRowIndex;
        }
    }

}
