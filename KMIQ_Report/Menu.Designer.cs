namespace KMIQ
{
    partial class Menu : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Menu()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_LoadData = this.Factory.CreateRibbonButton();
            this.btn_ClearDB = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.cb_Student = this.Factory.CreateRibbonComboBox();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btn_PrintOut = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btn_Online_Publish = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.btn_About = this.Factory.CreateRibbonButton();
            this.tab2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group1);
            this.tab2.Groups.Add(this.group2);
            this.tab2.Groups.Add(this.group3);
            this.tab2.Groups.Add(this.group4);
            this.tab2.Groups.Add(this.group5);
            this.tab2.Label = "SMART REPORT";
            this.tab2.Name = "tab2";
            this.tab2.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabHome");
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_LoadData);
            this.group1.Items.Add(this.btn_ClearDB);
            this.group1.Label = "데이터 불러오기";
            this.group1.Name = "group1";
            // 
            // btn_LoadData
            // 
            this.btn_LoadData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_LoadData.Label = "불러오기";
            this.btn_LoadData.Name = "btn_LoadData";
            this.btn_LoadData.OfficeImageId = "ServerRestoreSqlDatabase";
            this.btn_LoadData.ShowImage = true;
            this.btn_LoadData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_LoadData_Click);
            // 
            // btn_ClearDB
            // 
            this.btn_ClearDB.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_ClearDB.Image = global::KMIQ.Properties.Resources.attention_100;
            this.btn_ClearDB.Label = "Clear";
            this.btn_ClearDB.Name = "btn_ClearDB";
            this.btn_ClearDB.ShowImage = true;
            this.btn_ClearDB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ClearDB_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.label2);
            this.group2.Items.Add(this.cb_Student);
            this.group2.Label = "리포트 보기";
            this.group2.Name = "group2";
            // 
            // label2
            // 
            this.label2.Label = "▒ 대상을 선택하세요 ▒";
            this.label2.Name = "label2";
            // 
            // cb_Student
            // 
            this.cb_Student.Label = "Select Person";
            this.cb_Student.Name = "cb_Student";
            this.cb_Student.ShowItemImage = false;
            this.cb_Student.ShowLabel = false;
            this.cb_Student.SizeString = "wwwwwwwwwwwwww";
            this.cb_Student.Text = null;
            this.cb_Student.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cb_Student_ItemsLoading);
            this.cb_Student.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cb_Student_TextChanged);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btn_PrintOut);
            this.group3.Label = "리포트 출력";
            this.group3.Name = "group3";
            // 
            // btn_PrintOut
            // 
            this.btn_PrintOut.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_PrintOut.Label = "리포트출력";
            this.btn_PrintOut.Name = "btn_PrintOut";
            this.btn_PrintOut.OfficeImageId = "PrintOptionsMenu";
            this.btn_PrintOut.ShowImage = true;
            this.btn_PrintOut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_PrintOut_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btn_Online_Publish);
            this.group4.Label = "온라인 검사";
            this.group4.Name = "group4";
            // 
            // btn_Online_Publish
            // 
            this.btn_Online_Publish.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Online_Publish.Label = "온라인검사 리포트발행";
            this.btn_Online_Publish.Name = "btn_Online_Publish";
            this.btn_Online_Publish.OfficeImageId = "ViewWebLayoutView";
            this.btn_Online_Publish.ShowImage = true;
            this.btn_Online_Publish.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Online_Publish_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.btn_About);
            this.group5.Label = "About";
            this.group5.Name = "group5";
            // 
            // btn_About
            // 
            this.btn_About.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_About.Image = global::KMIQ.Properties.Resources.ip_icon_04_Info;
            this.btn_About.Label = "프로그램 정보";
            this.btn_About.Name = "btn_About";
            this.btn_About.ShowImage = true;
            this.btn_About.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_About_Click);
            // 
            // Menu
            // 
            this.Name = "Menu";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Menu_Load);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_LoadData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ClearDB;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cb_Student;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_PrintOut;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_About;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Online_Publish;
    }

    partial class ThisRibbonCollection
    {
        internal Menu Menu
        {
            get { return this.GetRibbon<Menu>(); }
        }
    }
}
