namespace SnakeGameExcel
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnInit = this.Factory.CreateRibbonButton();
            this.btnStart = this.Factory.CreateRibbonButton();
            this.btnSetup = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.btnRight = this.Factory.CreateRibbonButton();
            this.btnLeft = this.Factory.CreateRibbonButton();
            this.btnUp = this.Factory.CreateRibbonButton();
            this.btnDown = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "スネークゲームTR+";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnInit);
            this.group1.Items.Add(this.btnStart);
            this.group1.Label = "ゲーム操作";
            this.group1.Name = "group1";
            // 
            // btnInit
            // 
            this.btnInit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInit.Image = ((System.Drawing.Image)(resources.GetObject("btnInit.Image")));
            this.btnInit.Label = "ゲーム初期化";
            this.btnInit.Name = "btnInit";
            this.btnInit.ShowImage = true;
            this.btnInit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInit_Click);
            // 
            // btnStart
            // 
            this.btnStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStart.Image = ((System.Drawing.Image)(resources.GetObject("btnStart.Image")));
            this.btnStart.Label = "ゲーム開始";
            this.btnStart.Name = "btnStart";
            this.btnStart.ShowImage = true;
            this.btnStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStart_Click);
            // 
            // btnSetup
            // 
            this.btnSetup.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSetup.Image = ((System.Drawing.Image)(resources.GetObject("btnSetup.Image")));
            this.btnSetup.Label = "ゲーム設定";
            this.btnSetup.Name = "btnSetup";
            this.btnSetup.ShowImage = true;
            this.btnSetup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetup_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnLeft);
            this.group2.Items.Add(this.btnUp);
            this.group2.Items.Add(this.btnDown);
            this.group2.Items.Add(this.btnRight);
            this.group2.Label = "操作";
            this.group2.Name = "group2";
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnSetup);
            this.group3.Items.Add(this.btnHelp);
            this.group3.Label = "その他";
            this.group3.Name = "group3";
            // 
            // btnHelp
            // 
            this.btnHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHelp.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image")));
            this.btnHelp.Label = "ヘルプ";
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.ShowImage = true;
            this.btnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHelp_Click);
            // 
            // btnRight
            // 
            this.btnRight.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRight.Image = ((System.Drawing.Image)(resources.GetObject("btnRight.Image")));
            this.btnRight.Label = " ";
            this.btnRight.Name = "btnRight";
            this.btnRight.ShowImage = true;
            this.btnRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRight_Click);
            // 
            // btnLeft
            // 
            this.btnLeft.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLeft.Image = ((System.Drawing.Image)(resources.GetObject("btnLeft.Image")));
            this.btnLeft.Label = " ";
            this.btnLeft.Name = "btnLeft";
            this.btnLeft.ShowImage = true;
            this.btnLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLeft_Click);
            // 
            // btnUp
            // 
            this.btnUp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUp.Image = ((System.Drawing.Image)(resources.GetObject("btnUp.Image")));
            this.btnUp.Label = " ";
            this.btnUp.Name = "btnUp";
            this.btnUp.ShowImage = true;
            this.btnUp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUp_Click);
            // 
            // btnDown
            // 
            this.btnDown.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDown.Image = ((System.Drawing.Image)(resources.GetObject("btnDown.Image")));
            this.btnDown.Label = " ";
            this.btnDown.Name = "btnDown";
            this.btnDown.ShowImage = true;
            this.btnDown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDown_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
