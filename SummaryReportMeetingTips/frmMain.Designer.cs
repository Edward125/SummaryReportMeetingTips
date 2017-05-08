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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabRawData = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lblRawData = new System.Windows.Forms.Label();
            this.txtRawDataFile = new System.Windows.Forms.TextBox();
            this.btnImportData = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.tabRawData.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabRawData);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1253, 587);
            this.tabControl1.TabIndex = 0;
            // 
            // tabRawData
            // 
            this.tabRawData.Controls.Add(this.textBox1);
            this.tabRawData.Controls.Add(this.btnImportData);
            this.tabRawData.Controls.Add(this.txtRawDataFile);
            this.tabRawData.Controls.Add(this.lblRawData);
            this.tabRawData.Location = new System.Drawing.Point(4, 23);
            this.tabRawData.Name = "tabRawData";
            this.tabRawData.Padding = new System.Windows.Forms.Padding(3);
            this.tabRawData.Size = new System.Drawing.Size(1245, 560);
            this.tabRawData.TabIndex = 0;
            this.tabRawData.Text = "RawData";
            this.tabRawData.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 23);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1245, 503);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
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
            // txtRawDataFile
            // 
            this.txtRawDataFile.Location = new System.Drawing.Point(105, 10);
            this.txtRawDataFile.Name = "txtRawDataFile";
            this.txtRawDataFile.Size = new System.Drawing.Size(371, 22);
            this.txtRawDataFile.TabIndex = 1;
            this.txtRawDataFile.DoubleClick += new System.EventHandler(this.txtRawDataFile_DoubleClick);
            // 
            // btnImportData
            // 
            this.btnImportData.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnImportData.Location = new System.Drawing.Point(482, 9);
            this.btnImportData.Name = "btnImportData";
            this.btnImportData.Size = new System.Drawing.Size(86, 23);
            this.btnImportData.TabIndex = 2;
            this.btnImportData.Text = "Import Data";
            this.btnImportData.UseVisualStyleBackColor = true;
            this.btnImportData.Click += new System.EventHandler(this.btnImportData_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(574, 10);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(665, 22);
            this.textBox1.TabIndex = 3;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1277, 611);
            this.Controls.Add(this.tabControl1);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "frmMain";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabRawData.ResumeLayout(false);
            this.tabRawData.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabRawData;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label lblRawData;
        private System.Windows.Forms.Button btnImportData;
        private System.Windows.Forms.TextBox txtRawDataFile;
        private System.Windows.Forms.TextBox textBox1;
    }
}

