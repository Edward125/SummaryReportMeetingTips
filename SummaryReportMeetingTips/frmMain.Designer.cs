namespace SummaryReportMeetingTips
{
    partial class frmMain
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tabRawData = new System.Windows.Forms.TabPage();
            this.btnImportData = new System.Windows.Forms.Button();
            this.btnQuery = new System.Windows.Forms.Button();
            this.comboRawDataType = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboSheetList = new System.Windows.Forms.ComboBox();
            this.btnAnalyzeFile = new System.Windows.Forms.Button();
            this.datagridRawData = new System.Windows.Forms.DataGridView();
            this.txtRawDataFile = new System.Windows.Forms.TextBox();
            this.lblRawData = new System.Windows.Forms.Label();
            this.tabReport = new System.Windows.Forms.TabPage();
            this.grbReportChildNode = new System.Windows.Forms.GroupBox();
            this.btnSaveReport = new System.Windows.Forms.Button();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.txtReportLastUpdateTime = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtReportNewTipsSaveTime = new System.Windows.Forms.TextBox();
            this.txtReportParentTotalTime = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtReportNewTipsOptimizePCT = new System.Windows.Forms.TextBox();
            this.txtReportNewTips = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtReportHaveTipsOptimizePCTTotal = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtReportHaveTipsOptimizePCT = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtReportNewTipsOptimizePCTTotal = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtReportHaveTipsSaveTime = new System.Windows.Forms.TextBox();
            this.txtReportAlreadyHaveTips = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtReportTotalTime = new System.Windows.Forms.TextBox();
            this.lblTotalWorkTime = new System.Windows.Forms.Label();
            this.grbParentnode = new System.Windows.Forms.GroupBox();
            this.txtParentType = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtReportSummary = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lstviewReport = new System.Windows.Forms.ListView();
            this.trviewReport = new System.Windows.Forms.TreeView();
            this.tabMeeting = new System.Windows.Forms.TabPage();
            this.lstviewMeeting = new System.Windows.Forms.ListView();
            this.trviewMeeting = new System.Windows.Forms.TreeView();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.tsslStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.tabMain.SuspendLayout();
            this.tabRawData.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.datagridRawData)).BeginInit();
            this.tabReport.SuspendLayout();
            this.grbReportChildNode.SuspendLayout();
            this.grbParentnode.SuspendLayout();
            this.tabMeeting.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tabRawData);
            this.tabMain.Controls.Add(this.tabReport);
            this.tabMain.Controls.Add(this.tabMeeting);
            this.tabMain.Location = new System.Drawing.Point(12, 22);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(1292, 564);
            this.tabMain.TabIndex = 0;
            this.tabMain.Selected += new System.Windows.Forms.TabControlEventHandler(this.tabMain_Selected);
            // 
            // tabRawData
            // 
            this.tabRawData.Controls.Add(this.btnImportData);
            this.tabRawData.Controls.Add(this.btnQuery);
            this.tabRawData.Controls.Add(this.comboRawDataType);
            this.tabRawData.Controls.Add(this.label1);
            this.tabRawData.Controls.Add(this.comboSheetList);
            this.tabRawData.Controls.Add(this.btnAnalyzeFile);
            this.tabRawData.Controls.Add(this.datagridRawData);
            this.tabRawData.Controls.Add(this.txtRawDataFile);
            this.tabRawData.Controls.Add(this.lblRawData);
            this.tabRawData.Location = new System.Drawing.Point(4, 23);
            this.tabRawData.Name = "tabRawData";
            this.tabRawData.Padding = new System.Windows.Forms.Padding(3);
            this.tabRawData.Size = new System.Drawing.Size(1284, 537);
            this.tabRawData.TabIndex = 0;
            this.tabRawData.Text = "RawData";
            this.tabRawData.UseVisualStyleBackColor = true;
            // 
            // btnImportData
            // 
            this.btnImportData.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnImportData.Location = new System.Drawing.Point(947, 11);
            this.btnImportData.Name = "btnImportData";
            this.btnImportData.Size = new System.Drawing.Size(86, 23);
            this.btnImportData.TabIndex = 2;
            this.btnImportData.Text = "Import Data";
            this.btnImportData.UseVisualStyleBackColor = true;
            this.btnImportData.Click += new System.EventHandler(this.btnImportData_Click);
            // 
            // btnQuery
            // 
            this.btnQuery.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnQuery.Location = new System.Drawing.Point(1157, 11);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(116, 23);
            this.btnQuery.TabIndex = 9;
            this.btnQuery.Text = "Query Raw Data";
            this.btnQuery.UseVisualStyleBackColor = true;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // comboRawDataType
            // 
            this.comboRawDataType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboRawDataType.FormattingEnabled = true;
            this.comboRawDataType.Items.AddRange(new object[] {
            "Report",
            "Meeting"});
            this.comboRawDataType.Location = new System.Drawing.Point(1057, 12);
            this.comboRawDataType.Name = "comboRawDataType";
            this.comboRawDataType.Size = new System.Drawing.Size(94, 22);
            this.comboRawDataType.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(1037, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(16, 19);
            this.label1.TabIndex = 7;
            this.label1.Text = "|";
            // 
            // comboSheetList
            // 
            this.comboSheetList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboSheetList.FormattingEnabled = true;
            this.comboSheetList.Location = new System.Drawing.Point(818, 11);
            this.comboSheetList.Name = "comboSheetList";
            this.comboSheetList.Size = new System.Drawing.Size(123, 22);
            this.comboSheetList.TabIndex = 6;
            this.comboSheetList.SelectedIndexChanged += new System.EventHandler(this.comboSheetList_SelectedIndexChanged);
            // 
            // btnAnalyzeFile
            // 
            this.btnAnalyzeFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnalyzeFile.Location = new System.Drawing.Point(726, 10);
            this.btnAnalyzeFile.Name = "btnAnalyzeFile";
            this.btnAnalyzeFile.Size = new System.Drawing.Size(86, 23);
            this.btnAnalyzeFile.TabIndex = 5;
            this.btnAnalyzeFile.Text = "Analyze File";
            this.btnAnalyzeFile.UseVisualStyleBackColor = true;
            this.btnAnalyzeFile.Click += new System.EventHandler(this.btnAnalyzeFile_Click);
            // 
            // datagridRawData
            // 
            this.datagridRawData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.datagridRawData.Location = new System.Drawing.Point(6, 38);
            this.datagridRawData.Name = "datagridRawData";
            this.datagridRawData.RowTemplate.Height = 23;
            this.datagridRawData.Size = new System.Drawing.Size(1272, 503);
            this.datagridRawData.TabIndex = 4;
            // 
            // txtRawDataFile
            // 
            this.txtRawDataFile.Location = new System.Drawing.Point(105, 10);
            this.txtRawDataFile.Name = "txtRawDataFile";
            this.txtRawDataFile.Size = new System.Drawing.Size(615, 22);
            this.txtRawDataFile.TabIndex = 1;
            this.txtRawDataFile.DoubleClick += new System.EventHandler(this.txtRawDataFile_DoubleClick);
            // 
            // lblRawData
            // 
            this.lblRawData.AutoSize = true;
            this.lblRawData.Location = new System.Drawing.Point(13, 13);
            this.lblRawData.Name = "lblRawData";
            this.lblRawData.Size = new System.Drawing.Size(86, 14);
            this.lblRawData.TabIndex = 0;
            this.lblRawData.Text = "Raw Data File:";
            // 
            // tabReport
            // 
            this.tabReport.Controls.Add(this.grbReportChildNode);
            this.tabReport.Controls.Add(this.grbParentnode);
            this.tabReport.Controls.Add(this.lstviewReport);
            this.tabReport.Controls.Add(this.trviewReport);
            this.tabReport.Location = new System.Drawing.Point(4, 23);
            this.tabReport.Name = "tabReport";
            this.tabReport.Padding = new System.Windows.Forms.Padding(3);
            this.tabReport.Size = new System.Drawing.Size(1284, 537);
            this.tabReport.TabIndex = 1;
            this.tabReport.Text = "Report";
            this.tabReport.UseVisualStyleBackColor = true;
            // 
            // grbReportChildNode
            // 
            this.grbReportChildNode.Controls.Add(this.btnSaveReport);
            this.grbReportChildNode.Controls.Add(this.label13);
            this.grbReportChildNode.Controls.Add(this.label12);
            this.grbReportChildNode.Controls.Add(this.txtReportLastUpdateTime);
            this.grbReportChildNode.Controls.Add(this.label11);
            this.grbReportChildNode.Controls.Add(this.txtReportNewTipsSaveTime);
            this.grbReportChildNode.Controls.Add(this.txtReportParentTotalTime);
            this.grbReportChildNode.Controls.Add(this.label4);
            this.grbReportChildNode.Controls.Add(this.txtReportNewTipsOptimizePCT);
            this.grbReportChildNode.Controls.Add(this.txtReportNewTips);
            this.grbReportChildNode.Controls.Add(this.label10);
            this.grbReportChildNode.Controls.Add(this.txtReportHaveTipsOptimizePCTTotal);
            this.grbReportChildNode.Controls.Add(this.label9);
            this.grbReportChildNode.Controls.Add(this.txtReportHaveTipsOptimizePCT);
            this.grbReportChildNode.Controls.Add(this.label8);
            this.grbReportChildNode.Controls.Add(this.label7);
            this.grbReportChildNode.Controls.Add(this.txtReportNewTipsOptimizePCTTotal);
            this.grbReportChildNode.Controls.Add(this.label6);
            this.grbReportChildNode.Controls.Add(this.txtReportHaveTipsSaveTime);
            this.grbReportChildNode.Controls.Add(this.txtReportAlreadyHaveTips);
            this.grbReportChildNode.Controls.Add(this.label5);
            this.grbReportChildNode.Controls.Add(this.txtReportTotalTime);
            this.grbReportChildNode.Controls.Add(this.lblTotalWorkTime);
            this.grbReportChildNode.Location = new System.Drawing.Point(475, 5);
            this.grbReportChildNode.Name = "grbReportChildNode";
            this.grbReportChildNode.Size = new System.Drawing.Size(801, 71);
            this.grbReportChildNode.TabIndex = 3;
            this.grbReportChildNode.TabStop = false;
            // 
            // btnSaveReport
            // 
            this.btnSaveReport.Location = new System.Drawing.Point(635, 37);
            this.btnSaveReport.Name = "btnSaveReport";
            this.btnSaveReport.Size = new System.Drawing.Size(156, 29);
            this.btnSaveReport.TabIndex = 25;
            this.btnSaveReport.Text = "Save";
            this.btnSaveReport.UseVisualStyleBackColor = true;
            this.btnSaveReport.Click += new System.EventHandler(this.btnSaveReport_Click);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(404, 44);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(43, 14);
            this.label13.TabIndex = 24;
            this.label13.Text = "改善比";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(504, 43);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(79, 14);
            this.label12.TabIndex = 23;
            this.label12.Text = "改善总工时比";
            // 
            // txtReportLastUpdateTime
            // 
            this.txtReportLastUpdateTime.Location = new System.Drawing.Point(713, 11);
            this.txtReportLastUpdateTime.Name = "txtReportLastUpdateTime";
            this.txtReportLastUpdateTime.ReadOnly = true;
            this.txtReportLastUpdateTime.Size = new System.Drawing.Size(78, 22);
            this.txtReportLastUpdateTime.TabIndex = 22;
            this.txtReportLastUpdateTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(261, 46);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(97, 14);
            this.label11.TabIndex = 21;
            this.label11.Text = "TipsSaveTime(h)";
            // 
            // txtReportNewTipsSaveTime
            // 
            this.txtReportNewTipsSaveTime.Location = new System.Drawing.Point(358, 41);
            this.txtReportNewTipsSaveTime.Name = "txtReportNewTipsSaveTime";
            this.txtReportNewTipsSaveTime.Size = new System.Drawing.Size(41, 22);
            this.txtReportNewTipsSaveTime.TabIndex = 20;
            this.txtReportNewTipsSaveTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtReportNewTipsSaveTime.TextChanged += new System.EventHandler(this.txtReportNewTipsSaveTime_TextChanged);
            this.txtReportNewTipsSaveTime.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNewTipsSaveTime_KeyPress);
            // 
            // txtReportParentTotalTime
            // 
            this.txtReportParentTotalTime.Location = new System.Drawing.Point(95, 42);
            this.txtReportParentTotalTime.Name = "txtReportParentTotalTime";
            this.txtReportParentTotalTime.ReadOnly = true;
            this.txtReportParentTotalTime.Size = new System.Drawing.Size(64, 22);
            this.txtReportParentTotalTime.TabIndex = 19;
            this.txtReportParentTotalTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(0, 18);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(94, 14);
            this.label4.TabIndex = 18;
            this.label4.Text = "Report总工时(h)";
            // 
            // txtReportNewTipsOptimizePCT
            // 
            this.txtReportNewTipsOptimizePCT.Location = new System.Drawing.Point(447, 40);
            this.txtReportNewTipsOptimizePCT.Name = "txtReportNewTipsOptimizePCT";
            this.txtReportNewTipsOptimizePCT.ReadOnly = true;
            this.txtReportNewTipsOptimizePCT.Size = new System.Drawing.Size(55, 22);
            this.txtReportNewTipsOptimizePCT.TabIndex = 17;
            this.txtReportNewTipsOptimizePCT.Text = "0";
            this.txtReportNewTipsOptimizePCT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtReportNewTips
            // 
            this.txtReportNewTips.Location = new System.Drawing.Point(218, 45);
            this.txtReportNewTips.Name = "txtReportNewTips";
            this.txtReportNewTips.Size = new System.Drawing.Size(36, 22);
            this.txtReportNewTips.TabIndex = 16;
            this.txtReportNewTips.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtReportNewTips.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNewTips_KeyPress);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(164, 46);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(54, 14);
            this.label10.TabIndex = 15;
            this.label10.Text = "新增Tips";
            // 
            // txtReportHaveTipsOptimizePCTTotal
            // 
            this.txtReportHaveTipsOptimizePCTTotal.Location = new System.Drawing.Point(583, 12);
            this.txtReportHaveTipsOptimizePCTTotal.Name = "txtReportHaveTipsOptimizePCTTotal";
            this.txtReportHaveTipsOptimizePCTTotal.ReadOnly = true;
            this.txtReportHaveTipsOptimizePCTTotal.Size = new System.Drawing.Size(41, 22);
            this.txtReportHaveTipsOptimizePCTTotal.TabIndex = 14;
            this.txtReportHaveTipsOptimizePCTTotal.Text = "0";
            this.txtReportHaveTipsOptimizePCTTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(504, 15);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(79, 14);
            this.label9.TabIndex = 13;
            this.label9.Text = "改善总工时比";
            // 
            // txtReportHaveTipsOptimizePCT
            // 
            this.txtReportHaveTipsOptimizePCT.Location = new System.Drawing.Point(447, 14);
            this.txtReportHaveTipsOptimizePCT.Name = "txtReportHaveTipsOptimizePCT";
            this.txtReportHaveTipsOptimizePCT.ReadOnly = true;
            this.txtReportHaveTipsOptimizePCT.Size = new System.Drawing.Size(54, 22);
            this.txtReportHaveTipsOptimizePCT.TabIndex = 12;
            this.txtReportHaveTipsOptimizePCT.Text = "0";
            this.txtReportHaveTipsOptimizePCT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(405, 17);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(43, 14);
            this.label8.TabIndex = 11;
            this.label8.Text = "改善比";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(632, 14);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(79, 14);
            this.label7.TabIndex = 10;
            this.label7.Text = "上次更新时间";
            // 
            // txtReportNewTipsOptimizePCTTotal
            // 
            this.txtReportNewTipsOptimizePCTTotal.Location = new System.Drawing.Point(583, 40);
            this.txtReportNewTipsOptimizePCTTotal.Name = "txtReportNewTipsOptimizePCTTotal";
            this.txtReportNewTipsOptimizePCTTotal.ReadOnly = true;
            this.txtReportNewTipsOptimizePCTTotal.Size = new System.Drawing.Size(41, 22);
            this.txtReportNewTipsOptimizePCTTotal.TabIndex = 9;
            this.txtReportNewTipsOptimizePCTTotal.Text = "0";
            this.txtReportNewTipsOptimizePCTTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(260, 18);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(97, 14);
            this.label6.TabIndex = 8;
            this.label6.Text = "TipsSaveTime(h)";
            // 
            // txtReportHaveTipsSaveTime
            // 
            this.txtReportHaveTipsSaveTime.Location = new System.Drawing.Point(357, 14);
            this.txtReportHaveTipsSaveTime.Name = "txtReportHaveTipsSaveTime";
            this.txtReportHaveTipsSaveTime.ReadOnly = true;
            this.txtReportHaveTipsSaveTime.Size = new System.Drawing.Size(41, 22);
            this.txtReportHaveTipsSaveTime.TabIndex = 7;
            this.txtReportHaveTipsSaveTime.Text = "0";
            this.txtReportHaveTipsSaveTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtReportAlreadyHaveTips
            // 
            this.txtReportAlreadyHaveTips.Location = new System.Drawing.Point(218, 14);
            this.txtReportAlreadyHaveTips.Name = "txtReportAlreadyHaveTips";
            this.txtReportAlreadyHaveTips.ReadOnly = true;
            this.txtReportAlreadyHaveTips.Size = new System.Drawing.Size(36, 22);
            this.txtReportAlreadyHaveTips.TabIndex = 6;
            this.txtReportAlreadyHaveTips.Text = "0";
            this.txtReportAlreadyHaveTips.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(163, 17);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(54, 14);
            this.label5.TabIndex = 5;
            this.label5.Text = "已有Tips";
            // 
            // txtReportTotalTime
            // 
            this.txtReportTotalTime.Location = new System.Drawing.Point(95, 16);
            this.txtReportTotalTime.Name = "txtReportTotalTime";
            this.txtReportTotalTime.ReadOnly = true;
            this.txtReportTotalTime.Size = new System.Drawing.Size(64, 22);
            this.txtReportTotalTime.TabIndex = 4;
            this.txtReportTotalTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // lblTotalWorkTime
            // 
            this.lblTotalWorkTime.AutoSize = true;
            this.lblTotalWorkTime.Location = new System.Drawing.Point(3, 47);
            this.lblTotalWorkTime.Name = "lblTotalWorkTime";
            this.lblTotalWorkTime.Size = new System.Drawing.Size(82, 14);
            this.lblTotalWorkTime.TabIndex = 4;
            this.lblTotalWorkTime.Text = "该项总工时(h)";
            // 
            // grbParentnode
            // 
            this.grbParentnode.Controls.Add(this.txtParentType);
            this.grbParentnode.Controls.Add(this.txtReportSummary);
            this.grbParentnode.Controls.Add(this.label2);
            this.grbParentnode.Controls.Add(this.label3);
            this.grbParentnode.Location = new System.Drawing.Point(6, 6);
            this.grbParentnode.Name = "grbParentnode";
            this.grbParentnode.Size = new System.Drawing.Size(463, 71);
            this.grbParentnode.TabIndex = 2;
            this.grbParentnode.TabStop = false;
            this.grbParentnode.Text = "Summary";
            // 
            // txtParentType
            // 
            this.txtParentType.Location = new System.Drawing.Point(67, 41);
            this.txtParentType.Name = "txtParentType";
            this.txtParentType.ReadOnly = true;
            this.txtParentType.Size = new System.Drawing.Size(390, 22);
            this.txtParentType.TabIndex = 3;
            this.txtParentType.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(1, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 14);
            this.label3.TabIndex = 2;
            this.label3.Text = "ParentType";
            // 
            // txtReportSummary
            // 
            this.txtReportSummary.Location = new System.Drawing.Point(67, 14);
            this.txtReportSummary.Name = "txtReportSummary";
            this.txtReportSummary.ReadOnly = true;
            this.txtReportSummary.Size = new System.Drawing.Size(390, 22);
            this.txtReportSummary.TabIndex = 1;
            this.txtReportSummary.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 14);
            this.label2.TabIndex = 0;
            this.label2.Text = "Summary";
            // 
            // lstviewReport
            // 
            this.lstviewReport.BackColor = System.Drawing.Color.White;
            this.lstviewReport.CheckBoxes = true;
            this.lstviewReport.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstviewReport.Location = new System.Drawing.Point(409, 82);
            this.lstviewReport.Name = "lstviewReport";
            this.lstviewReport.Size = new System.Drawing.Size(867, 459);
            this.lstviewReport.TabIndex = 1;
            this.lstviewReport.UseCompatibleStateImageBehavior = false;
            this.lstviewReport.View = System.Windows.Forms.View.Details;
            // 
            // trviewReport
            // 
            this.trviewReport.Location = new System.Drawing.Point(3, 83);
            this.trviewReport.Name = "trviewReport";
            this.trviewReport.Size = new System.Drawing.Size(400, 458);
            this.trviewReport.TabIndex = 0;
            this.trviewReport.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.trviewReport_AfterSelect);
            // 
            // tabMeeting
            // 
            this.tabMeeting.Controls.Add(this.lstviewMeeting);
            this.tabMeeting.Controls.Add(this.trviewMeeting);
            this.tabMeeting.Location = new System.Drawing.Point(4, 23);
            this.tabMeeting.Name = "tabMeeting";
            this.tabMeeting.Size = new System.Drawing.Size(1284, 537);
            this.tabMeeting.TabIndex = 2;
            this.tabMeeting.Text = "Meeting";
            this.tabMeeting.UseVisualStyleBackColor = true;
            // 
            // lstviewMeeting
            // 
            this.lstviewMeeting.BackColor = System.Drawing.Color.White;
            this.lstviewMeeting.CheckBoxes = true;
            this.lstviewMeeting.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstviewMeeting.Location = new System.Drawing.Point(409, 39);
            this.lstviewMeeting.Name = "lstviewMeeting";
            this.lstviewMeeting.Size = new System.Drawing.Size(872, 502);
            this.lstviewMeeting.TabIndex = 1;
            this.lstviewMeeting.UseCompatibleStateImageBehavior = false;
            this.lstviewMeeting.View = System.Windows.Forms.View.Details;
            // 
            // trviewMeeting
            // 
            this.trviewMeeting.Location = new System.Drawing.Point(3, 39);
            this.trviewMeeting.Name = "trviewMeeting";
            this.trviewMeeting.Size = new System.Drawing.Size(400, 502);
            this.trviewMeeting.TabIndex = 0;
            this.trviewMeeting.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.trviewMeeting_AfterSelect);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsslStatus});
            this.statusStrip1.Location = new System.Drawing.Point(0, 589);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1317, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // tsslStatus
            // 
            this.tsslStatus.Name = "tsslStatus";
            this.tsslStatus.Size = new System.Drawing.Size(0, 17);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1317, 611);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.tabMain);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmMain";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.tabMain.ResumeLayout(false);
            this.tabRawData.ResumeLayout(false);
            this.tabRawData.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.datagridRawData)).EndInit();
            this.tabReport.ResumeLayout(false);
            this.grbReportChildNode.ResumeLayout(false);
            this.grbReportChildNode.PerformLayout();
            this.grbParentnode.ResumeLayout(false);
            this.grbParentnode.PerformLayout();
            this.tabMeeting.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabMain;
        private System.Windows.Forms.TabPage tabRawData;
        private System.Windows.Forms.TabPage tabReport;
        private System.Windows.Forms.Label lblRawData;
        private System.Windows.Forms.Button btnImportData;
        private System.Windows.Forms.TextBox txtRawDataFile;
        private System.Windows.Forms.DataGridView datagridRawData;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.Button btnAnalyzeFile;
        private System.Windows.Forms.ComboBox comboSheetList;
        private System.Windows.Forms.TabPage tabMeeting;
        private System.Windows.Forms.TreeView trviewReport;
        private System.Windows.Forms.ListView lstviewReport;
        private System.Windows.Forms.TreeView trviewMeeting;
        private System.Windows.Forms.ListView lstviewMeeting;
        private System.Windows.Forms.ToolStripStatusLabel tsslStatus;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.ComboBox comboRawDataType;
        private System.Windows.Forms.GroupBox grbReportChildNode;
        private System.Windows.Forms.GroupBox grbParentnode;
        private System.Windows.Forms.TextBox txtReportHaveTipsOptimizePCTTotal;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtReportHaveTipsOptimizePCT;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtReportNewTipsOptimizePCTTotal;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtReportHaveTipsSaveTime;
        private System.Windows.Forms.TextBox txtReportAlreadyHaveTips;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtReportTotalTime;
        private System.Windows.Forms.Label lblTotalWorkTime;
        private System.Windows.Forms.TextBox txtParentType;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtReportSummary;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtReportNewTipsOptimizePCT;
        private System.Windows.Forms.TextBox txtReportNewTips;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtReportParentTotalTime;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtReportNewTipsSaveTime;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtReportLastUpdateTime;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button btnSaveReport;
    }
}

