namespace hisDBurl
{
    partial class hisDBurlGUI : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public hisDBurlGUI()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.logi = this.Factory.CreateRibbonGroup();
            this.setAllText = this.Factory.CreateRibbonButton();
            this.httpsGetPost = this.Factory.CreateRibbonButton();
            this.ahcmsG = this.Factory.CreateRibbonGroup();
            this.ahcmsGenUrl = this.Factory.CreateRibbonButton();
            this.ndapG = this.Factory.CreateRibbonGroup();
            this.ndapGenUrl = this.Factory.CreateRibbonButton();
            this.ahtwhG = this.Factory.CreateRibbonGroup();
            this.ahtwhGenUrl = this.Factory.CreateRibbonButton();
            this.hoacc = this.Factory.CreateRibbonEditBox();
            this.hopsw = this.Factory.CreateRibbonEditBox();
            this.hoDownload = this.Factory.CreateRibbonButton();
            this.hoCookie = this.Factory.CreateRibbonEditBox();
            this.hosa = this.Factory.CreateRibbonButton();
            this.debugG = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.logi.SuspendLayout();
            this.ahcmsG.SuspendLayout();
            this.ndapG.SuspendLayout();
            this.ahtwhG.SuspendLayout();
            this.debugG.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.logi);
            this.tab1.Groups.Add(this.ahcmsG);
            this.tab1.Groups.Add(this.ndapG);
            this.tab1.Groups.Add(this.ahtwhG);
            this.tab1.Groups.Add(this.debugG);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // logi
            // 
            this.logi.Items.Add(this.setAllText);
            this.logi.Items.Add(this.httpsGetPost);
            this.logi.Label = "庶務";
            this.logi.Name = "logi";
            // 
            // setAllText
            // 
            this.setAllText.Label = "文字化";
            this.setAllText.Name = "setAllText";
            this.setAllText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.setAllText_Click);
            // 
            // httpsGetPost
            // 
            this.httpsGetPost.Label = "取得影像連結";
            this.httpsGetPost.Name = "httpsGetPost";
            this.httpsGetPost.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // ahcmsG
            // 
            this.ahcmsG.Items.Add(this.ahcmsGenUrl);
            this.ahcmsG.Items.Add(this.hoacc);
            this.ahcmsG.Items.Add(this.hopsw);
            this.ahcmsG.Items.Add(this.hoDownload);
            this.ahcmsG.Items.Add(this.hoCookie);
            this.ahcmsG.Items.Add(this.hosa);
            this.ahcmsG.Label = "國史館檔案史料文物查詢系統";
            this.ahcmsG.Name = "ahcmsG";
            // 
            // ahcmsGenUrl
            // 
            this.ahcmsGenUrl.Label = "加超連結";
            this.ahcmsGenUrl.Name = "ahcmsGenUrl";
            this.ahcmsGenUrl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ahcmsGenUrl_Click);
            // 
            // ndapG
            // 
            this.ndapG.Items.Add(this.ndapGenUrl);
            this.ndapG.Label = "臺灣省議會史料總庫";
            this.ndapG.Name = "ndapG";
            // 
            // ndapGenUrl
            // 
            this.ndapGenUrl.Label = "加超連結";
            this.ndapGenUrl.Name = "ndapGenUrl";
            this.ndapGenUrl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ndapGenUrl_Click);
            // 
            // ahtwhG
            // 
            this.ahtwhG.Items.Add(this.ahtwhGenUrl);
            this.ahtwhG.Label = "國史館臺灣文獻館典藏管理系統";
            this.ahtwhG.Name = "ahtwhG";
            // 
            // ahtwhGenUrl
            // 
            this.ahtwhGenUrl.Label = "加超連結";
            this.ahtwhGenUrl.Name = "ahtwhGenUrl";
            this.ahtwhGenUrl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ahtwhGenUrl_Click);
            // 
            // hoacc
            // 
            this.hoacc.Label = "帳號";
            this.hoacc.Name = "hoacc";
            // 
            // hopsw
            // 
            this.hopsw.Label = "密碼";
            this.hopsw.Name = "hopsw";
            // 
            // hoDownload
            // 
            this.hoDownload.Label = "下載";
            this.hoDownload.Name = "hoDownload";
            this.hoDownload.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hoDownload_Click);
            // 
            // hoCookie
            // 
            this.hoCookie.Label = "cookie";
            this.hoCookie.Name = "hoCookie";
            // 
            // hosa
            // 
            this.hosa.Label = "檢視登入狀態";
            this.hosa.Name = "hosa";
            this.hosa.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hosa_Click);
            // 
            // debugG
            // 
            this.debugG.Items.Add(this.button1);
            this.debugG.Label = "開發人員專用勿動";
            this.debugG.Name = "debugG";
            // 
            // button1
            // 
            this.button1.Label = "button1";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // hisDBurlGUI
            // 
            this.Name = "hisDBurlGUI";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.logi.ResumeLayout(false);
            this.logi.PerformLayout();
            this.ahcmsG.ResumeLayout(false);
            this.ahcmsG.PerformLayout();
            this.ndapG.ResumeLayout(false);
            this.ndapG.PerformLayout();
            this.ahtwhG.ResumeLayout(false);
            this.ahtwhG.PerformLayout();
            this.debugG.ResumeLayout(false);
            this.debugG.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ahcmsG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ahcmsGenUrl;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup logi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton setAllText;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ndapG;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ahtwhG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ndapGenUrl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ahtwhGenUrl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton httpsGetPost;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox hoacc;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox hopsw;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton hoDownload;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox hoCookie;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton hosa;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup debugG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal hisDBurlGUI Ribbon1
        {
            get { return this.GetRibbon<hisDBurlGUI>(); }
        }
    }
}
