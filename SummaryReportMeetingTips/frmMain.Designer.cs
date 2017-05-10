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
            this.grbChildNode = new System.Windows.Forms.GroupBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtNewTipsSaveTime = new System.Windows.Forms.TextBox();
            this.txtReportParentTotalTime = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtNewTipsOptimizePCT = new System.Windows.Forms.TextBox();
            this.txtNewTips = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtHaveTipsOptimizePCTTotal = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtHaveTipsOptimizePCT = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtNewTipsOptimizePCTTotal = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtHaveTipsSaveTime = new System.Windows.Forms.TextBox();
            this.txtAlreadyHaveTips = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtReportTotalTime = new System.Windows.Forms.TextBox();
            this.lblTotalWorkTime = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
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
            this.txtLastUpdateTime = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.btnSaveReport = new System.Windows.Forms.Button();
            this.tabMain.SuspendLayout();
            this.tabRawData.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.datagridRawData)).BeginInit();
            this.tabReport.SuspendLayout();
            this.grbChildNode.SuspendLayout();
            this.groupBox1.SuspendLayout();
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
            this.btnImportData.Location = new System.Drawing.Point(913, 11);
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
            this.btnQuery.Location = new System.Drawing.Point(1123, 11);
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
            this.comboRawDataType.Location = new System.Drawing.Point(1023, 12);
            this.comboRawDataType.Name = "comboRawDataType";
            this.comboRawDataType.Size = new System.Drawing.Size(94, 22);
            this.comboRawDataType.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(1003, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(16, 19);
            this.label1.TabIndex = 7;
            this.label1.Text = "|";
            // 
            // comboSheetList
            // 
            this.comboSheetList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboSheetList.FormattingEnabled = true;
            this.comboSheetList.Location = new System.Drawing.Point(784, 11);
            this.comboSheetList.Name = "comboSheetList";
            this.comboSheetList.Size = new System.Drawing.Size(123, 22);
            this.comboSheetList.TabIndex = 6;
            this.comboSheetList.SelectedIndexChanged += new System.EventHandler(this.comboSheetList_SelectedIndexChanged);
            // 
            // btnAnalyzeFile
            // 
            this.btnAnalyzeFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnalyzeFile.Location = new System.Drawing.Point(692, 10);
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
            this.datagridRawData.Size = new System.Drawing.Size(1233, 503);
            this.datagridRawData.TabIndex = 4;
            // 
            // txtRawDataFile
            // 
            this.txtRawDataFile.Location = new System.Drawing.Point(105, 10);
            this.txtRawDataFile.Name = "txtRawDataFile";
            this.txtRawDataFile.Size = new System.Drawing.Size(581, 22);
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
            this.tabReport.Controls.Add(this.grbChildNode);
            this.tabReport.Controls.Add(this.groupBox1);
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
            // grbChildNode
            // 
            this.grbChildNode.Controls.Add(this.btnSaveReport);
            this.grbChildNode.Controls.Add(this.label13);
            this.grbChildNode.Controls.Add(this.label12);
            this.grbChildNode.Controls.Add(this.txtLastUpdateTime);
            this.grbChildNode.Controls.Add(this.label11);
            this.grbChildNode.Controls.Add(this.txtNewTipsSaveTime);
            this.grbChildNode.Controls.Add(this.txtReportParentTotalTime);
            this.grbChildNode.Controls.Add(this.label4);
            this.grbChildNode.Controls.Add(this.txtNewTipsOptimizePCT);
            this.grbChildNode.Controls.Add(this.txtNewTips);
            this.grbChildNode.Controls.Add(this.label10);
            this.grbChildNode.Controls.Add(this.txtHaveTipsOptimizePCTTotal);
            this.grbChildNode.Controls.Add(this.label9);
            this.grbChildNode.Controls.Add(this.txtHaveTipsOptimizePCT);
            this.grbChildNode.Controls.Add(this.label8);
            this.grbChildNode.Controls.Add(this.label7);
            this.grbChildNode.Controls.Add(this.txtNewTipsOptimizePCTTotal);
            this.grbChildNode.Controls.Add(this.label6);
            this.grbChildNode.Controls.Add(this.txtHaveTipsSaveTime);
            this.grbChildNode.Controls.Add(this.txtAlreadyHaveTips);
            this.grbChildNode.Controls.Add(this.label5);
            this.grbChildNode.Controls.Add(this.txtReportTotalTime);
            this.grbChildNode.Controls.Add(this.lblTotalWorkTime);
            this.grbChildNode.Location = new System.Drawing.Point(475, 5);
            this.grbChildNode.Name = "grbChildNode";
            this.grbChildNode.Size = new System.Drawing.Size(801, 71);
            this.grbChildNode.TabIndex = 3;
            this.grbChildNode.TabStop = false;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(273, 46);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(97, 14);
            this.label11.TabIndex = 21;
            this.label11.Text = "TipsSaveTime(h)";
            // 
            // txtNewTipsSaveTime
            // 
            this.txtNewTipsSaveTime.Location = new System.Drawing.Point(373, 41);
            this.txtNewTipsSaveTime.Name = "txtNewTipsSaveTime";
            this.txtNewTipsSaveTime.Size = new System.Drawing.Size(41, 22);
            this.txtNewTipsSaveTime.TabIndex = 20;
            // 
            // txtReportParentTotalTime
            // 
            this.txtReportParentTotalTime.Location = new System.Drawing.Point(95, 42);
            this.txtReportParentTotalTime.Name = "txtReportParentTotalTime";
            this.txtReportParentTotalTime.Size = new System.Drawing.Size(64, 22);
            this.txtReportParentTotalTime.TabIndex = 19;
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
            // txtNewTipsOptimizePCT
            // 
            this.txtNewTipsOptimizePCT.Location = new System.Drawing.Point(462, 40);
            this.txtNewTipsOptimizePCT.Name = "txtNewTipsOptimizePCT";
            this.txtNewTipsOptimizePCT.Size = new System.Drawing.Size(40, 22);
            this.txtNewTipsOptimizePCT.TabIndex = 17;
            // 
            // txtNewTips
            // 
            this.txtNewTips.Location = new System.Drawing.Point(218, 45);
            this.txtNewTips.Name = "txtNewTips";
            this.txtNewTips.Size = new System.Drawing.Size(50, 22);
            this.txtNewTips.TabIndex = 16;
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
            // txtHaveTipsOptimizePCTTotal
            // 
            this.txtHaveTipsOptimizePCTTotal.Location = new System.Drawing.Point(583, 12);
            this.txtHaveTipsOptimizePCTTotal.Name = "txtHaveTipsOptimizePCTTotal";
            this.txtHaveTipsOptimizePCTTotal.Size = new System.Drawing.Size(41, 22);
            this.txtHaveTipsOptimizePCTTotal.TabIndex = 14;
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
            // txtHaveTipsOptimizePCT
            // 
            this.txtHaveTipsOptimizePCT.Location = new System.Drawing.Point(460, 14);
            this.txtHaveTipsOptimizePCT.Name = "txtHaveTipsOptimizePCT";
            this.txtHaveTipsOptimizePCT.Size = new System.Drawing.Size(41, 22);
            this.txtHaveTipsOptimizePCT.TabIndex = 12;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(418, 17);
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
            // txtNewTipsOptimizePCTTotal
            // 
            this.txtNewTipsOptimizePCTTotal.Location = new System.Drawing.Point(583, 40);
            this.txtNewTipsOptimizePCTTotal.Name = "txtNewTipsOptimizePCTTotal";
            this.txtNewTipsOptimizePCTTotal.Size = new System.Drawing.Size(41, 22);
            this.txtNewTipsOptimizePCTTotal.TabIndex = 9;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(272, 18);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(97, 14);
            this.label6.TabIndex = 8;
            this.label6.Text = "TipsSaveTime(h)";
            // 
            // txtHaveTipsSaveTime
            // 
            this.txtHaveTipsSaveTime.Location = new System.Drawing.Point(372, 14);
            this.txtHaveTipsSaveTime.Name = "txtHaveTipsSaveTime";
            this.txtHaveTipsSaveTime.Size = new System.Drawing.Size(41, 22);
            this.txtHaveTipsSaveTime.TabIndex = 7;
            // 
            // txtAlreadyHaveTips
            // 
            this.txtAlreadyHaveTips.Location = new System.Drawing.Point(218, 14);
            this.txtAlreadyHaveTips.Name = "txtAlreadyHaveTips";
            this.txtAlreadyHaveTips.Size = new System.Drawing.Size(50, 22);
            this.txtAlreadyHaveTips.TabIndex = 6;
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
            this.txtReportTotalTime.Size = new System.Drawing.Size(64, 22);
            this.txtReportTotalTime.TabIndex = 4;
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
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtParentType);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtReportSummary);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(6, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(463, 71);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Summary";
            // 
            // txtParentType
            // 
            this.txtParentType.Location = new System.Drawing.Point(78, 41);
            this.txtParentType.Name = "txtParentType";
            this.txtParentType.ReadOnly = true;
            this.txtParentType.Size = new System.Drawing.Size(379, 22);
            this.txtParentType.TabIndex = 3;
            this.txtParentType.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(69, 14);
            this.label3.TabIndex = 2;
            this.label3.Text = "Parent Type";
            // 
            // txtReportSummary
            // 
            this.txtReportSummary.Location = new System.Drawing.Point(106, 14);
            this.txtReportSummary.Name = "txtReportSummary";
            this.txtReportSummary.ReadOnly = true;
            this.txtReportSummary.Size = new System.Drawing.Size(351, 22);
            this.txtReportSummary.TabIndex = 1;
            this.txtReportSummary.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 14);
            this.label2.TabIndex = 0;
            this.label2.Text = "Report Summary";
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
            this.lstviewMeeting.Size = new System.Drawing.Size(830, 502);
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
            // txtLastUpdateTime
            // 
            this.txtLastUpdateTime.Location = new System.Drawing.Point(713, 11);
            this.txtLastUpdateTime.Name = "txtLastUpdateTime";
            this.txtLastUpdateTime.Size = new System.Drawing.Size(78, 22);
            this.txtLastUpdateTime.TabIndex = 22;
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
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(418, 44);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(43, 14);
            this.label13.TabIndex = 24;
            this.label13.Text = "改善比";
            // 
            // btnSaveReport
            // 
            this.btnSaveReport.Location = new System.Drawing.Point(635, 37);
            this.btnSaveReport.Name = "btnSaveReport";
            this.btnSaveReport.Size = new System.Drawing.Size(156, 29);
            this.btnSaveReport.TabIndex = 25;
            this.btnSaveReport.Text = "Save";
            this.btnSaveReport.UseVisualStyleBackColor = true;
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
            this.grbChildNode.ResumeLayout(false);
            this.grbChildNode.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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
        private System.Windows.Forms.GroupBox grbChildNode;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtHaveTipsOptimizePCTTotal;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtHaveTipsOptimizePCT;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtNewTipsOptimizePCTTotal;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtHaveTipsSaveTime;
        private System.Windows.Forms.TextBox txtAlreadyHaveTips;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtReportTotalTime;
        private System.Windows.Forms.Label lblTotalWorkTime;
        private System.Windows.Forms.TextBox txtParentType;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtReportSummary;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtNewTipsOptimizePCT;
        private System.Windows.Forms.TextBox txtNewTips;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtReportParentTotalTime;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtNewTipsSaveTime;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtLastUpdateTime;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button btnSaveReport;
    }
}

