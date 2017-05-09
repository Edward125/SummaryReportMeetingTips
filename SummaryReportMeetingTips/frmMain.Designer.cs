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
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tabRawData = new System.Windows.Forms.TabPage();
            this.comboSheetList = new System.Windows.Forms.ComboBox();
            this.btnAnalyzeFile = new System.Windows.Forms.Button();
            this.datagridRawData = new System.Windows.Forms.DataGridView();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnImportData = new System.Windows.Forms.Button();
            this.txtRawDataFile = new System.Windows.Forms.TextBox();
            this.lblRawData = new System.Windows.Forms.Label();
            this.tabReport = new System.Windows.Forms.TabPage();
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
            this.tabMeeting.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tabRawData);
            this.tabMain.Controls.Add(this.tabReport);
            this.tabMain.Controls.Add(this.tabMeeting);
            this.tabMain.Location = new System.Drawing.Point(12, 12);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(1253, 574);
            this.tabMain.TabIndex = 0;
            this.tabMain.SelectedIndexChanged += new System.EventHandler(this.tabMain_SelectedIndexChanged);
            this.tabMain.Selected += new System.Windows.Forms.TabControlEventHandler(this.tabMain_Selected);
            // 
            // tabRawData
            // 
            this.tabRawData.Controls.Add(this.comboSheetList);
            this.tabRawData.Controls.Add(this.btnAnalyzeFile);
            this.tabRawData.Controls.Add(this.datagridRawData);
            this.tabRawData.Controls.Add(this.textBox1);
            this.tabRawData.Controls.Add(this.btnImportData);
            this.tabRawData.Controls.Add(this.txtRawDataFile);
            this.tabRawData.Controls.Add(this.lblRawData);
            this.tabRawData.Location = new System.Drawing.Point(4, 23);
            this.tabRawData.Name = "tabRawData";
            this.tabRawData.Padding = new System.Windows.Forms.Padding(3);
            this.tabRawData.Size = new System.Drawing.Size(1245, 547);
            this.tabRawData.TabIndex = 0;
            this.tabRawData.Text = "RawData";
            this.tabRawData.UseVisualStyleBackColor = true;
            // 
            // comboSheetList
            // 
            this.comboSheetList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboSheetList.FormattingEnabled = true;
            this.comboSheetList.Location = new System.Drawing.Point(507, 11);
            this.comboSheetList.Name = "comboSheetList";
            this.comboSheetList.Size = new System.Drawing.Size(157, 22);
            this.comboSheetList.TabIndex = 6;
            this.comboSheetList.SelectedIndexChanged += new System.EventHandler(this.comboSheetList_SelectedIndexChanged);
            // 
            // btnAnalyzeFile
            // 
            this.btnAnalyzeFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnalyzeFile.Location = new System.Drawing.Point(415, 10);
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
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(762, 10);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(477, 22);
            this.textBox1.TabIndex = 3;
            // 
            // btnImportData
            // 
            this.btnImportData.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnImportData.Location = new System.Drawing.Point(670, 9);
            this.btnImportData.Name = "btnImportData";
            this.btnImportData.Size = new System.Drawing.Size(86, 23);
            this.btnImportData.TabIndex = 2;
            this.btnImportData.Text = "Import Data";
            this.btnImportData.UseVisualStyleBackColor = true;
            this.btnImportData.Click += new System.EventHandler(this.btnImportData_Click);
            // 
            // txtRawDataFile
            // 
            this.txtRawDataFile.Location = new System.Drawing.Point(105, 10);
            this.txtRawDataFile.Name = "txtRawDataFile";
            this.txtRawDataFile.Size = new System.Drawing.Size(304, 22);
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
            this.tabReport.Controls.Add(this.lstviewReport);
            this.tabReport.Controls.Add(this.trviewReport);
            this.tabReport.Location = new System.Drawing.Point(4, 23);
            this.tabReport.Name = "tabReport";
            this.tabReport.Padding = new System.Windows.Forms.Padding(3);
            this.tabReport.Size = new System.Drawing.Size(1245, 547);
            this.tabReport.TabIndex = 1;
            this.tabReport.Text = "Report";
            this.tabReport.UseVisualStyleBackColor = true;
            // 
            // lstviewReport
            // 
            this.lstviewReport.BackColor = System.Drawing.Color.White;
            this.lstviewReport.CheckBoxes = true;
            this.lstviewReport.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstviewReport.Location = new System.Drawing.Point(409, 39);
            this.lstviewReport.Name = "lstviewReport";
            this.lstviewReport.Size = new System.Drawing.Size(830, 502);
            this.lstviewReport.TabIndex = 1;
            this.lstviewReport.UseCompatibleStateImageBehavior = false;
            this.lstviewReport.View = System.Windows.Forms.View.Details;
            // 
            // trviewReport
            // 
            this.trviewReport.Location = new System.Drawing.Point(3, 39);
            this.trviewReport.Name = "trviewReport";
            this.trviewReport.Size = new System.Drawing.Size(400, 502);
            this.trviewReport.TabIndex = 0;
            this.trviewReport.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.trviewReport_AfterSelect);
            // 
            // tabMeeting
            // 
            this.tabMeeting.Controls.Add(this.lstviewMeeting);
            this.tabMeeting.Controls.Add(this.trviewMeeting);
            this.tabMeeting.Location = new System.Drawing.Point(4, 23);
            this.tabMeeting.Name = "tabMeeting";
            this.tabMeeting.Size = new System.Drawing.Size(1245, 547);
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
            this.statusStrip1.Size = new System.Drawing.Size(1277, 22);
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
            this.ClientSize = new System.Drawing.Size(1277, 611);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.tabMain);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "frmMain";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.tabMain.ResumeLayout(false);
            this.tabRawData.ResumeLayout(false);
            this.tabRawData.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.datagridRawData)).EndInit();
            this.tabReport.ResumeLayout(false);
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
        private System.Windows.Forms.TextBox textBox1;
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
    }
}

