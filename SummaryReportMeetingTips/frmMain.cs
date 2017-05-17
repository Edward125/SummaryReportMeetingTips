using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Edward;
using System.IO;
using System.Data.OleDb;
using System.Data.SQLite;
using System.Diagnostics;
using System.Collections;

namespace SummaryReportMeetingTips
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new System.Drawing.Point(10, 10);
        }


        #region 窗体放大缩小

        private float X;
        private float Y;

        private void setTag(Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                con.Tag = con.Width + ":" + con.Height + ":" + con.Left + ":" + con.Top + ":" + con.Font.Size;
                if (con.Controls.Count > 0)
                    setTag(con);
            }
        }
        private void setControls(float newx, float newy, Control cons)
        {
            foreach (Control con in cons.Controls)
            {

                string[] mytag = con.Tag.ToString().Split(new char[] { ':' });
                float a = Convert.ToSingle(mytag[0]) * newx;
                con.Width = (int)a;
                a = Convert.ToSingle(mytag[1]) * newy;
                con.Height = (int)(a);
                a = Convert.ToSingle(mytag[2]) * newx;
                con.Left = (int)(a);
                a = Convert.ToSingle(mytag[3]) * newy;
                con.Top = (int)(a);
                Single currentSize = Convert.ToSingle(mytag[4]) * Math.Min(newx, newy);
                con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);
                if (con.Controls.Count > 0)
                {
                    setControls(newx, newy, con);
                }
            }

        }

        void Form1_Resize(object sender, EventArgs e)
        {
            float newx = (this.Width) / X;
            float newy = this.Height / Y;
            setControls(newx, newy, this);
            // this.Text = this.Width.ToString() + " " + this.Height.ToString();

        }

        #endregion


        private DataSet ds = new DataSet();
        private string lastParentNode = "";
        private string lastChildNode = "";
        //p.WorkType savelogworktype;

        private void frmMain_Load(object sender, EventArgs e)
        {
            this.Text = Application.ProductName + "<only for internal use>,ver:" + Application.ProductVersion + ",author:edward_song@yeah.net";
            //
            //窗体放大缩小
            this.Resize += new EventHandler(Form1_Resize);
            X = this.Width;
            Y = this.Height;
            setTag(this);
            Form1_Resize(new object(), new EventArgs());//x,y可在实例化时赋值,最后这句是新加的，在MDI时有用
            //
            txtRawDataFile.SetWatermark("Double Click here to select the raw data file(.xls,.xlsx)");

            if (!p.CheckFolder())
                Environment.Exit(0);
            if (!p.InitDataDB())
                Environment.Exit(0);
            tabMain.SelectedIndex = 1;
            setListview(lstviewMeeting, p.WorkType.Meeting);
            setListview(lstviewReport, p.WorkType.Report);
            //
            comboRawDataType.SelectedIndex = 0;



            if ((p.queryCount("SELECT COUNT(*) FROM t_reportrawdata") > 0) && (p.queryCount("SELECT COUNT(*) FROM t_meetingrawdata") > 0))
            {
                loadRawdata2DataGrid();
            }
            else
            {
                tsslStatus.ForeColor = Color.Red;
                tsslStatus.Text = "Raw data in database is incomplete,pls check...";
            }

            


        }

        private void txtRawDataFile_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            if (open.ShowDialog() == DialogResult.OK)
            {

                FileInfo fi = new FileInfo(open.FileName);
                if ((fi.Extension == ".xls") || (fi.Extension == ".xlsx"))
                {
                   txtRawDataFile.Text = open.FileName;
                }
                else
                {
                    MessageBox.Show("you select file is not excel file...", "File Not Excel", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

            }
        }


        #region DataSet


        static bool DataSetParse(string fileName, out DataSet ds)
        {
            // string connectionString = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", fileName);


            ////2003（Microsoft.Jet.Oledb.4.0）
            //string strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);
            ////2010（Microsoft.ACE.OLEDB.12.0）
            //string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);

            string connectionString = string.Empty;
            System.IO.FileInfo fi = new System.IO.FileInfo(fileName);
            //MessageBox.Show(fi.Extension);
            DataSet data = new DataSet();
            try
            {
                if (fi.Extension == ".xls")
                    connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", fileName);
                if (fi.Extension == ".xlsx")
                    connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", fileName);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                ds = data;
                return false;
            }






            foreach (var sheetName in GetExcelSheetNames(connectionString))
            {
                
                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    Console.WriteLine(sheetName);
                    var dataTable = new System.Data.DataTable(sheetName);
                    string query = string.Format("SELECT * FROM [{0}]", sheetName);
                    con.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                    adapter.Fill(dataTable);
                    data.Tables.Add(dataTable);

                }
            }

            ds = data;

            return true;

        }

        static bool DataSetParse(string fileName, out DataSet ds,ComboBox combo)
        {
            combo.Items.Clear();
            // string connectionString = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", fileName);


            ////2003（Microsoft.Jet.Oledb.4.0）
            //string strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);
            ////2010（Microsoft.ACE.OLEDB.12.0）
            //string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);

            string connectionString = string.Empty;
            System.IO.FileInfo fi = new System.IO.FileInfo(fileName);
            //MessageBox.Show(fi.Extension);
            DataSet data = new DataSet();
            try
            {
                if (fi.Extension == ".xls")
                    connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", fileName);
                if (fi.Extension == ".xlsx")
                    connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", fileName);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                ds = data;
                return false;
            }






            foreach (var sheetName in GetExcelSheetNames(connectionString))
            {
                combo.Items.Add(sheetName);
                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    Console.WriteLine(sheetName);
                    var dataTable = new System.Data.DataTable(sheetName);
                    string query = string.Format("SELECT * FROM [{0}]", sheetName);
                    con.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                    adapter.Fill(dataTable);
                    data.Tables.Add(dataTable);

                }
            }

            ds = data;

            return true;

        }

        static string[] GetExcelSheetNames(string connectionString)
        {
            OleDbConnection con = null;
            System.Data.DataTable dt = null;
            con = new OleDbConnection(connectionString);
            con.Open();
            dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (dt == null)
            {
                return null;
            }

            String[] excelSheetNames = new String[dt.Rows.Count];
            int i = 0;

            foreach (DataRow row in dt.Rows)
            {
                excelSheetNames[i] = row["TABLE_NAME"].ToString();
                
               
                i++;
                
            }

            return excelSheetNames;
        }

        #endregion

        private void btnAnalyzeFile_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtRawDataFile.Text.Trim()))
                return;

            if (!File.Exists(txtRawDataFile.Text.Trim()))
            {
                MessageBox.Show("U select raw data file is not exist,pls check...", "File Not Exist", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                txtRawDataFile.SelectAll();
                txtRawDataFile.Focus();
                return;
            }

           
            if (!DataSetParse(txtRawDataFile.Text.Trim(), out ds, this.comboSheetList))
                return;



            if (this.comboSheetList.Items.Count > 0)
            {
                this.comboSheetList.SelectedIndex = 0;
                datagridRawData.DataSource = ds.Tables[0];
            }


        }

        private void btnImportData_Click(object sender, EventArgs e)
        {

            if (this.comboSheetList.SelectedIndex == -1)
                return;

            if (!checkFormat(ds, comboSheetList))
                return;
            string sql = "";
            if (comboSheetList.Text.Trim() == "會議$")                
                sql = "SELECT COUNT(*) FROM t_meetingrawdata";
            if (comboSheetList.Text.Trim() == "Report$")
                sql = "SELECT COUNT(*) FROM t_reportrawdata";             
            int line = 0;
            queryCount(sql, out line);
            if (line > 0)
            {
                DialogResult dt = MessageBox.Show("There are " + line + " records in database ,r u sure to append the data?", "Append or Cancel?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dt == DialogResult.Cancel)
                    return;
            }


            if (comboSheetList.Text.Trim() == "會議$")
            {
                this.Enabled = false;
                updateRawData(ds, comboSheetList, p.WorkType.Meeting);
                this.Enabled = true;
            }

            if (comboSheetList.Text.Trim() == "Report$")
            {
                this.Enabled = false;
                updateRawData(ds, comboSheetList, p.WorkType.Report);
                this.Enabled = true;
            }
        }

        private void comboSheetList_SelectedIndexChanged(object sender, EventArgs e)
        {
            datagridRawData.DataSource = ds.Tables[comboSheetList.SelectedIndex];
        }

        private bool checkFormat(DataSet ds,ComboBox combo)
        {

            try
            {
                ds.Tables[combo.SelectedIndex].Rows[0]["部門代碼"].ToString();
            }
            catch (ArgumentException ex)
            {

                MessageBox.Show(ex.Message + ",pls check format...", "Format Fail...", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return false;
            }

            try
            {
                ds.Tables[combo.SelectedIndex].Rows[0]["課別代碼"].ToString();
            }
            catch (ArgumentException ex)
            {

                MessageBox.Show(ex.Message + ",pls check format...", "Format Fail...", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return false;
            }

            try
            {
                ds.Tables[combo.SelectedIndex].Rows[0]["工號"].ToString();
            }
            catch (ArgumentException ex)
            {

                MessageBox.Show(ex.Message + ",pls check format...", "Format Fail...", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return false;
            }

            try
            {
                ds.Tables[combo.SelectedIndex].Rows[0]["英文姓名"].ToString();
            }
            catch (ArgumentException ex)
            {

                MessageBox.Show(ex.Message + ",pls check format...", "Format Fail...", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return false;
            }


           
            return true;
        }

        private bool updateRawData(DataSet ds, ComboBox combo,p.WorkType worktype)
        {

            SQLiteConnection conn = new SQLiteConnection(p.dbConnectionString);
            using (SQLiteCommand cmd = new SQLiteCommand())
            {
                conn.Open();
                cmd.Connection = conn;
                Stopwatch sw = new Stopwatch();
                sw.Start();

                SQLiteTransaction trans = conn.BeginTransaction();
                cmd.Transaction = trans;

                try
                {
                    for (int i = 0; i < ds.Tables[combo.SelectedIndex].Rows.Count; i++)
                    {
                        //textBox1.Text = ds.Tables[combo.SelectedIndex].Rows.Count.ToString();
                        if (worktype == p.WorkType.Meeting)
                        {
                            cmd.CommandText =
                                @"INSERT INTO t_meetingrawdata (depcode,
seccode,
opid,
engname,
meetingtype,
workcontent,
workdetail,
worktype,
isinworkbook,
ismywork,
vanva,
singleworktime,
weeklyworkfreq,
weeklyworktime,
monthlyworktime,
caller,
callerdep,
callerlevel,
optimizemethod,
weeklysavetime,
description,
reviewdate,
reviewer) VALUES (@_depcode,
@_seccode,
@_opid,
@_engname,
@_meetingtype,
@_workcontent,
@_workdetail,
@_worktype,
@_isinworkbook,
@_ismywork,
@_vanva,
@_singleworktime,
@_weeklyworkfreq,
@_weeklyworktime,
@_monthlyworktime,
@_caller,
@_callerdep,
@_callerlevel,
@_optimizemethod,
@_weeklysavetime,
@_description,
@_reviewdate,
@_reviewer)";

                            cmd.Parameters.Add(new SQLiteParameter("@_depcode", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_seccode", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_opid", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_engname", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_meetingtype", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_workcontent", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_workdetail", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_worktype", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_isinworkbook", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_ismywork", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_vanva", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_singleworktime", DbType.Decimal));
                            cmd.Parameters.Add(new SQLiteParameter(@"_weeklyworkfreq", DbType.Decimal));
                            cmd.Parameters.Add(new SQLiteParameter(@"_weeklyworktime", DbType.Decimal));
                            cmd.Parameters.Add(new SQLiteParameter(@"_monthlyworktime", DbType.Decimal));
                            cmd.Parameters.Add(new SQLiteParameter(@"_caller", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_callerdep", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_callerlevel", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_optimizemethod", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_weeklysavetime", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_description", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_reviewdate", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_reviewer", DbType.String));
                            //
                            cmd.Parameters[0].Value = ds.Tables[combo.SelectedIndex].Rows[i]["部門代碼"].ToString();
                            cmd.Parameters[1].Value = ds.Tables[combo.SelectedIndex].Rows[i]["課別代碼"].ToString();
                            cmd.Parameters[2].Value = ds.Tables[combo.SelectedIndex].Rows[i]["工號"].ToString();
                            cmd.Parameters[3].Value = ds.Tables[combo.SelectedIndex].Rows[i]["英文姓名"].ToString();

                            string _temp = ds.Tables[combo.SelectedIndex].Rows[i]["会议类型"].ToString();
                            p.replaceString(ref _temp);
                            cmd.Parameters[4].Value = _temp;

                            _temp = ds.Tables[combo.SelectedIndex].Rows[i]["工作內容"].ToString();
                            p.replaceString(ref  _temp);
                            cmd.Parameters[5].Value = _temp;

                            _temp = ds.Tables[combo.SelectedIndex].Rows[i]["工作細目"].ToString();
                             p.replaceString(ref  _temp);
                             cmd.Parameters[6].Value = _temp;

                            cmd.Parameters[7].Value = ds.Tables[combo.SelectedIndex].Rows[i]["工作分类"].ToString();
                            cmd.Parameters[8].Value = ds.Tables[combo.SelectedIndex].Rows[i]["是否在岗位说明书"].ToString();
                            cmd.Parameters[9].Value = ds.Tables[combo.SelectedIndex].Rows[i]["本职/非本职"].ToString();
                            cmd.Parameters[10].Value = ds.Tables[combo.SelectedIndex].Rows[i]["VA/NVA  (部级评核)"].ToString();
                            cmd.Parameters[11].Value = decimal.Round(Convert.ToDecimal(ds.Tables[combo.SelectedIndex].Rows[i]["单次工作时间（分钟）"]) / 60, 4);
                            cmd.Parameters[12].Value = Convert.ToDecimal(ds.Tables[combo.SelectedIndex].Rows[i]["周工作频率（次）"]);
                            cmd.Parameters[13].Value = decimal.Round(Convert.ToDecimal(ds.Tables[combo.SelectedIndex].Rows[i]["周工時(分)"]) / 60, 4);
                            cmd.Parameters[14].Value = decimal.Round(Convert.ToDecimal(ds.Tables[combo.SelectedIndex].Rows[i]["月工时（分）"]) / 60, 4);
                            cmd.Parameters[15].Value = ds.Tables[combo.SelectedIndex].Rows[i]["召集者"].ToString();
                            cmd.Parameters[16].Value = ds.Tables[combo.SelectedIndex].Rows[i]["召集者部門"].ToString();
                            cmd.Parameters[17].Value = ds.Tables[combo.SelectedIndex].Rows[i]["會議要求者級別"].ToString();
                            cmd.Parameters[18].Value = ds.Tables[combo.SelectedIndex].Rows[i]["改善方法"].ToString();
                            cmd.Parameters[19].Value = ds.Tables[combo.SelectedIndex].Rows[i]["節省周工時"].ToString();
                            cmd.Parameters[20].Value = ds.Tables[combo.SelectedIndex].Rows[i]["描述"].ToString();
                            //
                            string _date = ds.Tables[combo.SelectedIndex].Rows[i]["review date"].ToString();
                            if (string.IsNullOrEmpty(_date))
                            {
                                cmd.Parameters[21].Value = "";
                            }
                            else
                            {
                                cmd.Parameters[21].Value = Convert.ToDateTime(_date).ToString("yyyy-MM-dd");
                            }

                            cmd.Parameters[22].Value = ds.Tables[combo.SelectedIndex].Rows[i]["reviewer"].ToString();

                        }

                        if (worktype == p.WorkType.Report)
                        {
                            cmd.CommandText =
                                @"INSERT INTO t_reportrawdata (depcode,
seccode,
opid,
engname,
reporttype,
workcontent,
workdetail,
worktype,
isinworkbook,
ismywork,
vanva,
singleworktime,
weeklyworkfreq,
weeklyworktime,
monthlyworktime,
reportobject,
reporttype2,
reportmethod,
optimizemethod,
weeklysavetime,
description,
reviewdate,
reviewer) VALUES (@_depcode,
@_seccode,
@_opid,
@_engname,
@_reporttype,
@_workcontent,
@_workdetail,
@_worktype,
@_isinworkbook,
@_ismywork,
@_vanva,
@_singleworktime,
@_weeklyworkfreq,
@_weeklyworktime,
@_monthlyworktime,
@_reportobject,
@_reporttype2,
@_reportmethod,
@_optimizemethod,
@_weeklysavetime,
@_description,
@_reviewdate,
@_reviewer)";

                            cmd.Parameters.Add(new SQLiteParameter("@_depcode", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_seccode", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_opid", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_engname", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_reporttype", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_workcontent", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_workdetail", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_worktype", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_isinworkbook", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_ismywork", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_vanva", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_singleworktime", DbType.Decimal));
                            cmd.Parameters.Add(new SQLiteParameter(@"_weeklyworkfreq", DbType.Decimal));
                            cmd.Parameters.Add(new SQLiteParameter(@"_weeklyworktime", DbType.Decimal));
                            cmd.Parameters.Add(new SQLiteParameter(@"_monthlyworktime", DbType.Decimal));
                            cmd.Parameters.Add(new SQLiteParameter(@"_reportobject", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_reporttype2", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_reportmethod", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_optimizemethod", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_weeklysavetime", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_description", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_reviewdate", DbType.String));
                            cmd.Parameters.Add(new SQLiteParameter(@"_reviewer", DbType.String));
                            //
                            cmd.Parameters[0].Value = ds.Tables[combo.SelectedIndex].Rows[i]["部門代碼"].ToString();
                            cmd.Parameters[1].Value = ds.Tables[combo.SelectedIndex].Rows[i]["課別代碼"].ToString();
                            cmd.Parameters[2].Value = ds.Tables[combo.SelectedIndex].Rows[i]["工號"].ToString();
                            cmd.Parameters[3].Value = ds.Tables[combo.SelectedIndex].Rows[i]["英文姓名"].ToString();
                            cmd.Parameters[4].Value = ds.Tables[combo.SelectedIndex].Rows[i]["报告类型"].ToString();
                            cmd.Parameters[5].Value = ds.Tables[combo.SelectedIndex].Rows[i]["工作內容"].ToString();
                            cmd.Parameters[6].Value = ds.Tables[combo.SelectedIndex].Rows[i]["工作細目"].ToString();
                            cmd.Parameters[7].Value = ds.Tables[combo.SelectedIndex].Rows[i]["工作分类"].ToString();
                            cmd.Parameters[8].Value = ds.Tables[combo.SelectedIndex].Rows[i]["是否在岗位说明书"].ToString();
                            cmd.Parameters[9].Value = ds.Tables[combo.SelectedIndex].Rows[i]["本职/非本职"].ToString();
                            cmd.Parameters[10].Value = ds.Tables[combo.SelectedIndex].Rows[i]["VA/NVA  (部级评核)"].ToString();
                            cmd.Parameters[11].Value = decimal.Round(Convert.ToDecimal(ds.Tables[combo.SelectedIndex].Rows[i]["单次工作时间（分钟）"]) / 60, 4);
                            cmd.Parameters[12].Value = Convert.ToDecimal(ds.Tables[combo.SelectedIndex].Rows[i]["周工作频率（次）"]);
                            cmd.Parameters[13].Value = decimal.Round(Convert.ToDecimal(ds.Tables[combo.SelectedIndex].Rows[i]["周工時(分)"]) / 60, 4);
                            cmd.Parameters[14].Value = decimal.Round(Convert.ToDecimal(ds.Tables[combo.SelectedIndex].Rows[i]["月工时（分）"]) / 60, 4);
                            cmd.Parameters[15].Value = ds.Tables[combo.SelectedIndex].Rows[i]["報告對象"].ToString();
                            cmd.Parameters[16].Value = ds.Tables[combo.SelectedIndex].Rows[i]["報表類型"].ToString();
                            cmd.Parameters[17].Value = ds.Tables[combo.SelectedIndex].Rows[i]["製作方式"].ToString();
                            cmd.Parameters[18].Value = ds.Tables[combo.SelectedIndex].Rows[i]["改善方法"].ToString();
                            cmd.Parameters[19].Value = ds.Tables[combo.SelectedIndex].Rows[i]["節省周工時"].ToString();
                            cmd.Parameters[20].Value = ds.Tables[combo.SelectedIndex].Rows[i]["描述"].ToString();
                            //
                            string _date = ds.Tables[combo.SelectedIndex].Rows[i]["review date"].ToString();
                            if (string.IsNullOrEmpty(_date))
                            {
                                cmd.Parameters[21].Value = "";
                            }
                            else
                            {
                                cmd.Parameters[21].Value = Convert.ToDateTime(_date).ToString("yyyy-MM-dd");
                            }
                            cmd.Parameters[22].Value = ds.Tables[combo.SelectedIndex].Rows[i]["reviewer"].ToString();
                           // textBox1.Text = i.ToString();
                        }
                        cmd.ExecuteNonQuery();
                         
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.Rollback();
                    MessageBox.Show(ex.Message  , "INSERT FAIL", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    conn.Close ();
                    return false;
                }
                conn.Close ();
                sw.Stop();
                MessageBox.Show(string.Format("Insert into database {0} records,used time(ms):{1}", ds.Tables[combo.SelectedIndex].Rows.Count, sw.ElapsedMilliseconds.ToString ()));

            }


            return true;
        }



        private bool queryCount(string sql, out int linecount)
        {

             SQLiteConnection conn = new SQLiteConnection(p.dbConnectionString);

             try
             {
                 conn.Open();
                 SQLiteCommand cmd = new SQLiteCommand(sql, conn);
                 object o = cmd.ExecuteScalar();
                 linecount = Convert.ToInt16(o);
             }
             catch (Exception ex)
             {

                 MessageBox.Show(ex.Message, "QUERY FAIL", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                 linecount = 0;
                 return false;
             }
             finally
             {
                 conn.Close();
             }
          
            return true;
        }


        private void loadTreeViewData(TreeView trview, p.WorkType worktype)
        {
            lastParentNode = "";
            lastChildNode = "";
            //bool _bFindNode = false;
            //bool _bFindChildNode = false;
            trview.Nodes.Clear();

            string sql = "";
            if (worktype == p.WorkType.Report)
                sql = "SELECT * FROM t_reportrawdata ORDER by reporttype ASC";
            if (worktype == p.WorkType.Meeting)
                sql = "SELECT * FROM t_meetingrawdata order by meetingtype ASC";
            
            //
            SQLiteConnection conn = new SQLiteConnection(p.dbConnectionString);
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            SQLiteDataReader re = cmd.ExecuteReader();
            if (re.HasRows)
            {
                while (re.Read())
                {

                    string _nodename = string.Empty;
                    string _childnodename = string.Empty;

                    if (worktype == p.WorkType.Meeting)
                        _nodename = re["meetingtype"].ToString();
                    if (worktype == p.WorkType.Report)
                        _nodename = re["reporttype"].ToString();
                    _childnodename = re["workdetail"].ToString();
                    //
                    TreeNode tr = new TreeNode(_nodename);
                    //tr.Text = _nodename;
                    ////---------------                   
                    if ( _nodename != lastParentNode)
                    {
                        decimal _totaltime = 0;
                        decimal _savetime = 0;
                        int _itemscount = 0;
                        int _tipscount = 0;

                        loadNodeItemsTimeTipsSavetimeInfo(worktype, true, _nodename, out _totaltime, out _savetime, out _itemscount, out _tipscount);
                        tr.Text = _nodename + ";Items:" + _itemscount + ",TotalTime(h):" + _totaltime + ",Tips:" + _tipscount + ",SaveTime(h):" +
                           _savetime + ",PCT(%):" + p.CalcPCT(_savetime, _totaltime);
                    }
                    else
                        tr.Text = _nodename;
                    ////----------------

                    if (!nodeIsInTreView(trview,_nodename  ))
                    {
                         trview.Nodes.Add(tr);
                         lastParentNode = _nodename;
                          
                    }
                    //
                    int nodeindex = getNodeIndex( trview,tr.Text ) ;
                    if (nodeindex  != -1)
                    {
                        tr = trview.Nodes[nodeindex];
                        //
                        TreeNode childnode = new TreeNode();
                        //if (_childnodename != lastChildNode)
                        //{
                        //    decimal _totaltime = 0;
                        //    decimal _savetime = 0;
                        //    int _itemscount = 0;
                        //    int _tipscount = 0;

                        //    loadNodeItemsTimeTipsSavetimeInfo(worktype, false, _childnodename, out _totaltime, out _savetime, out _itemscount, out _tipscount);
                        //    _childnodename = _childnodename + ",TotalTime(h):" + _totaltime + ",Tips:" + _tipscount + ",SaveTime(h):" +
                        //       _savetime + ",PCT(%):" + p.CalcPCT(_savetime, _totaltime);

                        //    if (_tipscount > 0)
                        //    {
                        //        childnode.BackColor = Color.Green;
                        //        childnode.ForeColor = Color.White;
                        //    }
                        //}
                        childnode.Text = _childnodename;
                        if (!childnodeIsInTreView(tr, _childnodename))
                        {
                          //Application.DoEvents();

                            if (worktype == p.WorkType.Report)
                                sql = "SELECT SUM(tips) FROM t_reporttips WHERE workdetail = '" + _childnodename + "'";
                            if (worktype == p.WorkType.Meeting)
                                sql = "SELECT SUM(tips) FROM t_meetingtips WHERE workdetail = '" + _childnodename + "'";


         
                            if (p.queryCount(sql) > 0)
                            {

                                childnode.BackColor = Color.Green;
                                childnode.ForeColor = Color.White;
                            }
                            else
                            {
                                childnode.BackColor = Color.White;
                                childnode.ForeColor = Color.Black;
                            }
                            tr.Nodes.Add(childnode);
                            lastChildNode = _childnodename;
                        }
                    }
                }
            }
            conn.Close();


            trview.ExpandAll();
         }




        private void loadNodeItemsTimeTipsSavetimeInfo(p.WorkType worktype,bool _bselectparentnode,string _nodename,out decimal _totaltime,out decimal _savetime,out int _itemscount ,out int _tipscount)
        {
            _totaltime = _savetime = 0;
           _itemscount = _tipscount = 0;
            string sql ="";
            if (worktype == p.WorkType.Meeting)
            {
                if (_bselectparentnode)
                {
                    sql = "SELECT SUM (weeklyworktime) FROM t_meetingrawdata WHERE meetingtype = '" + _nodename + "'";
                    _totaltime = p.querySum(sql);
                    _savetime = p.querySum("SELECT SUM(tipsavetime) FROM t_meetingtips WHERE meetingtype = '" + _nodename + "'");
                    _itemscount = p.queryCount("SELECT COUNT(*) FROM t_meetingrawdata WHERE meetingtype = '" + _nodename + "'");
                    _tipscount = Convert.ToInt16(p.querySum("SELECT COUNT(tips) FROM t_meetingtips WHERE meetingtype = '" + _nodename + "'"));
                }
                else
                {
                    sql = "SELECT SUM (weeklyworktime) FROM t_meetingrawdata WHERE workdetail = '" + _nodename + "'";
                    _totaltime = p.querySum(sql);
                    _savetime = p.querySum("SELECT SUM(tipsavetime) FROM t_meetingtips WHERE workdetail = '" + _nodename + "'");
                    //_itemscount = p.queryCount("SELECT COUNT(*) FROM t_meetingrawdata WHERE workdetail = '" + _nodename + "'");
                    _tipscount = Convert.ToInt16(p.querySum("SELECT COUNT(tips) FROM t_meetingtips WHERE workdetail = '" + _nodename + "'"));
                }
                
            }
            if (worktype == p.WorkType.Report)
            {
                if (_bselectparentnode)
                {
                    sql = "SELECT SUM (weeklyworktime) FROM t_reportrawdata WHERE reporttype = '" + _nodename + "'";
                    _totaltime = p.querySum(sql);
                    _savetime = p.querySum("SELECT SUM(tipsavetime) FROM t_reporttips WHERE reporttype = '" + _nodename + "'");
                    _itemscount = p.queryCount("SELECT COUNT(*) FROM t_reportrawdata WHERE reporttype = '" + _nodename + "'");
                    _tipscount = Convert.ToInt16(p.querySum("SELECT COUNT(tips) FROM t_reporttips WHERE reporttype = '" + _nodename + "'"));
                }
                else
                {
                    sql = "SELECT SUM (weeklyworktime) FROM t_reportrawdata WHERE workdetail = '" + _nodename + "'";
                    _totaltime = p.querySum(sql);
                    _savetime = p.querySum("SELECT SUM(tipsavetime) FROM t_reporttips WHERE workdetail = '" + _nodename + "'");
                    _itemscount = p.queryCount("SELECT COUNT(*) FROM t_reportrawdata WHERE workdetail = '" + _nodename + "'");
                    _tipscount = Convert.ToInt16(p.querySum("SELECT COUNT(tips) FROM t_reporttips WHERE workdetail = '" + _nodename + "'"));
                }
                
            }
        }




        private void tabMain_Selected(object sender, TabControlEventArgs e)
        {
            tsslStatus.Text = "";
            if (e.TabPage == tabReport )
            {
                loadTreeViewData(trviewReport, p.WorkType.Report );
                trviewReport.Sort();

                if (trviewReport.Nodes.Count > 0)
                    trviewReport.SelectedNode = trviewReport.Nodes[0];



            }
            if (e.TabPage == tabMeeting )
            {
                loadTreeViewData(trviewMeeting, p.WorkType.Meeting);
                trviewMeeting.Sort();

                if (trviewMeeting.Nodes.Count > 0)
                    trviewMeeting.SelectedNode = trviewMeeting.Nodes[0];
            }
            if (e.TabPage == tabRawData)
            {
                if ((p.queryCount("SELECT COUNT(*) FROM t_reportrawdata") > 0) && (p.queryCount("SELECT COUNT(*) FROM t_meetingrawdata") > 0))
                {
                    loadRawdata2DataGrid();
                }
                else
                {
                    tsslStatus.ForeColor = Color.Red;
                    tsslStatus.Text = "Raw data in database is incomplete,pls check...";
                }

            }
  
        }

        /// <summary>
        /// check node is in treeview
        /// </summary>
        /// <param name="node">node</param>
        /// <param name="trview">treeview</param>
        /// <returns>exist,return true;not exist,return false </returns>
        private bool nodeIsInTreView(TreeNode node,TreeView trview)
        {
            if (trview.Nodes.Count > 0)
            {
                for (int i = 0; i < trview.Nodes.Count; i++)
                {
                    if (trview.Nodes[i].Text.Trim().ToUpper() == node.Text.Trim().ToUpper())
                        return true;
                }
            }
            return false;
        }


        /// <summary>
        /// check node is in treeview
        /// </summary>
        /// <param name="node">node</param>
        /// <param name="trview">treeview</param>
        /// <returns>exist,return true;not exist,return false </returns>
        private bool nodeIsInTreView( TreeView trview,string keyname)
        {
            if (trview.Nodes.Count > 0)
            {
                for (int i = 0; i < trview.Nodes.Count; i++)
                {
                    if (trview.Nodes[i].Text.Trim().ToUpper().StartsWith (keyname.ToUpper ().Trim ()))
                        return true;
                }
            }
            return false;
        }



        /// <summary>
        /// get node in treeview index 
        /// </summary>
        /// <param name="node">node</param>
        /// <param name="trview">treeview</param>
        /// <returns></returns>
        private int getNodeIndex(TreeNode node, TreeView trview)
        {
             if (trview.Nodes.Count > 0)
            {
                for (int i = 0; i < trview.Nodes.Count; i++)
                {
                    if (trview.Nodes[i].Text.Trim().ToUpper() == node.Text.Trim().ToUpper())
                        return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// get node in treeview index 
        /// </summary>
        /// <param name="node">node</param>
        /// <param name="trview">treeview</param>
        /// <returns></returns>
        private int getNodeIndex( TreeView trview,string keyname)
        {
            if (trview.Nodes.Count > 0)
            {
                for (int i = 0; i < trview.Nodes.Count; i++)
                {
                    if (trview.Nodes[i].Text.Trim().ToUpper().StartsWith (keyname.Trim ().ToUpper ()))
                        return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// childnode is in the treeview node
        /// </summary>
        /// <param name="node"></param>
        /// <param name="childnode"></param>
        /// <returns></returns>
        private bool childnodeIsInTreView(TreeNode node,  string  childnode)
        {
            if (node.Nodes.Count > 0)
            {
                for (int i = 0; i < node.Nodes.Count; i++)
                {
                    if (node.Nodes[i].Text.Trim().ToUpper().StartsWith(childnode.Trim().ToUpper())) 
                        return true;
                }
            }
            return false;
        }

        private void trviewReport_AfterSelect(object sender, TreeViewEventArgs e)
        {
            this.Enabled = false;
            loadInfo2toolStrip(p.WorkType.Report, lstviewReport, trviewReport);
            this.Enabled = true ;
        }

        private void loadInfo2toolStrip(p.WorkType worktype,ListView listview,TreeView treview)
        {

            //report
            txtReportNewTipsSaveTime.Text = "";
            txtReportNewTips.Text = "";
            txtReportNewTipsOptimizePCT.Text = "";
            txtReportNewTipsOptimizePCTTotal.Text = "";
            txtReportLastUpdateTime.Text = "";
            txtReportTotalTime.Text = "";
            txtReportAlreadyHaveTips.Text = "";
            txtReportHaveTipsSaveTime.Text = "";
            txtReportHaveTipsOptimizePCT.Text = "";
            txtReportHaveTipsOptimizePCTTotal.Text = "";

            txtReportNewReviewer.Text = "";
            txtReportNewDescription.Text = "";
            this.comboReportTipsStatus.SelectedIndex = -1;
            this.comboReportOptimizeMethod.SelectedIndex = -1;

            //meeting
            txtMeetingNewTipsSaveTime.Text = "";
            txtMeetingNewTips.Text = "";
            txtMeetingNewTipsOptimizePCT.Text = "";
            txtMeetingNewTipsOptimizePCTTotal.Text = "";
            txtMeetingLastUpdateTime.Text = "";
            txtMeetingTotalTime.Text = "";
            txtMeetingAlreadyHaveTips.Text = "";
            txtMeetingHaveTipsSaveTime.Text = "";
            txtMeetingHaveTipsOptimizePCT.Text = "";
            txtMeetingHaveTipsOptimizePCTTotal.Text = "";

            txtMeetingNewReviewer.Text = "";
            txtMeetingNewDescription.Text = "";
            this.comboMeetingTipsStatus.SelectedIndex = -1;
            this.comboMeetingOptimizeMethod.SelectedIndex = -1;




            bool _bSelectParentNode = false;
            listview.Items.Clear();
            string workdetail = "";
            string sql = "";
            try
            {

                    //workdetail = treview.SelectedNode.Parent.Text.Substring(0, treview.SelectedNode.Parent.Text.IndexOf(';'));
                    //workdetail = treview.SelectedNode.Text.Substring(0, treview.SelectedNode.Text.IndexOf(';'));

                workdetail = treview.SelectedNode.Text;

                if (workdetail.Contains("Items") && workdetail.Contains("TotalTime"))
                {
                    _bSelectParentNode = true;
                    workdetail = workdetail.Substring(0, workdetail.IndexOf(';'));
                    if (worktype == p.WorkType.Report)
                    {
                        grbReportChildNode.Enabled = false;
                        grbReportParentNode.Text = workdetail;
                    }
                    if (worktype == p.WorkType.Meeting)
                    {
                        grbMeetingChildNode.Enabled = false;
                        grbMeetingParentNode.Text = workdetail;
                    }
                }
                else
                {
                    _bSelectParentNode = false;
                    if (worktype == p.WorkType.Report)
                    {
                        grbReportChildNode.Enabled = true;
                        grbReportParentNode.Text = treview.SelectedNode.Parent.Text.Substring(0, treview.SelectedNode.Parent.Text.IndexOf(';'));
                    }
                    if (worktype == p.WorkType.Meeting)
                    {
                        grbMeetingChildNode.Enabled = true;
                        grbMeetingParentNode.Text = treview.SelectedNode.Parent.Text.Substring(0, treview.SelectedNode.Parent.Text.IndexOf(';'));
                    }
                }

            }
            catch (Exception)
            {
                //_bSelectParentNode = true;

         
                   // workdetail = treview.SelectedNode.Parent .Text.Substring(0, treview.SelectedNode.Text.IndexOf(';'));
             
              

            }

            //-------------------------------------
            if (!string.IsNullOrEmpty(workdetail))
            {
                int icount = -1;
                decimal totaltime = 0;//父项或者子项
                decimal totalworktime = 0;// 整张table
                if (worktype == p.WorkType.Report)
                {
                    if (_bSelectParentNode)
                       sql = "SELECT COUNT (reporttype) FROM t_reportrawdata WHERE reporttype = '" + workdetail + "'";
                    else 
                        sql = "SELECT COUNT (workdetail) FROM t_reportrawdata WHERE workdetail = '" + workdetail + "'";
                    icount = p.queryCount(sql);
                    if (_bSelectParentNode)
                        sql = "SELECT SUM (weeklyworktime) FROM t_reportrawdata WHERE reporttype = '" + workdetail + "'";
                    else 
                        sql = "SELECT SUM (weeklyworktime) FROM t_reportrawdata WHERE workdetail = '" + workdetail + "'";
                    totaltime = p.querySum(sql);//
                    loadData2ListView(listview, p.WorkType.Report, workdetail, _bSelectParentNode);
                    if (string.IsNullOrEmpty(txtReportTotalTime.Text.Trim()))
                    {
                        sql = "SELECT SUM (weeklyworktime) FROM t_reportrawdata";
                        totalworktime =  p.querySum(sql);
                        txtReportTotalTime.Text = totalworktime.ToString();
                    }
                }

                if (worktype == p.WorkType.Meeting)
                {
                    if (_bSelectParentNode)
                        sql = "SELECT COUNT (meetingtype) FROM t_meetingrawdata WHERE meetingtype = '" + workdetail + "'";
                    else 
                        sql = "SELECT COUNT (workdetail) FROM t_meetingrawdata WHERE workdetail = '" + workdetail + "'";
                    icount = p.queryCount(sql);
                    if (_bSelectParentNode )
                        sql = "SELECT SUM (weeklyworktime) FROM t_meetingrawdata WHERE meetingtype = '" + workdetail + "'";
                    else 
                        sql = "SELECT SUM (weeklyworktime) FROM t_meetingrawdata WHERE workdetail = '" + workdetail + "'";
                    totaltime = p.querySum(sql);
                    loadData2ListView(listview, p.WorkType.Meeting, workdetail, _bSelectParentNode);
                    if (string.IsNullOrEmpty(txtMeetingTotalTime.Text.Trim()))
                    {
                        sql = "SELECT SUM (weeklyworktime) FROM t_meetingrawdata";
                        totalworktime = p.querySum(sql);
                        txtMeetingTotalTime.Text = totalworktime.ToString();
                    }

                }

                tsslStatus.ForeColor = Color.Blue;
                tsslStatus.Text = workdetail + " | itemscount:" + icount + " | weeklyworktime(h):" + totaltime;
                //---------------------------
                if (_bSelectParentNode) //
                {
                    if (worktype == p.WorkType.Report)
                    {
                        //
                        //                     
                        decimal _savetime = p.querySum("SELECT SUM(tipsavetime) FROM t_reporttips");
                        txtReportSummary.Text = "Items:" + p.queryCount("SELECT COUNT(*) FROM t_reportrawdata") + ",TotalTime(h):" + totalworktime + ",Tips:" + p.querySum("SELECT COUNT(tips) FROM t_reporttips") + ",SaveTime(h):" +
                           _savetime + ",PCT(%):" + p.CalcPCT(_savetime, totalworktime);

                        int _itemscount, _tipscount = 0;
                        loadNodeItemsTimeTipsSavetimeInfo(worktype, _bSelectParentNode, workdetail, out totaltime, out _savetime, out _itemscount, out _tipscount);
                        //_savetime = p.querySum("SELECT SUM(tipsavetime) FROM t_reporttips WHERE reporttype = '" + workdetail +"'");
                        txtReportParentType.Text = "Items:" + _itemscount + ",TotalTime(h):" + totaltime + ",Tips:" + _tipscount + ",SaveTime(h):" +
                           _savetime + ",PCT(%):" + p.CalcPCT(_savetime, totaltime);


                        grbReportChildNode.Text = workdetail;
                        txtReportParentTotalTime.Text = totaltime.ToString();
                        //tips count 
                        sql = "SELECT SUM(tips) FROM t_reporttips where reporttype = '" + workdetail + "'";
                        txtReportAlreadyHaveTips.Text = p.querySum(sql).ToString();
                        //tips save time
                        sql = "SELECT SUM(tipsavetime) FROM t_reporttips where reporttype = '" + workdetail + "'";
                        txtReportHaveTipsSaveTime.Text = p.querySum(sql).ToString();
                        txtReportHaveTipsOptimizePCT.Text = p.CalcPCT(Convert.ToDecimal(txtReportHaveTipsSaveTime.Text.Trim()), totaltime);
                        txtReportHaveTipsOptimizePCTTotal.Text = p.CalcPCT(Convert.ToDecimal(txtReportHaveTipsSaveTime.Text.Trim()), totalworktime);

                    }


                    if (worktype == p.WorkType.Meeting)
                    {
                        //
                        //
                        decimal _savetime = p.querySum("SELECT SUM(tipsavetime) FROM t_meetingtips");
                        txtMeetingSummary.Text = "Items:" + p.queryCount("SELECT COUNT(*) FROM t_meetingrawdata") + ",TotalTime(h):" + totalworktime + ",Tips:" + p.querySum("SELECT COUNT(tips) FROM t_meetingtips") + ",SaveTime(h):" +
                           _savetime + ",PCT(%):" + p.CalcPCT(_savetime, totalworktime);

                        int _itemscount, _tipscount = 0;
                        loadNodeItemsTimeTipsSavetimeInfo(worktype, _bSelectParentNode, workdetail, out totaltime, out _savetime, out _itemscount, out _tipscount);
                        //_savetime = p.querySum("SELECT SUM(tipsavetime) FROM t_reporttips WHERE reporttype = '" + workdetail +"'");
                        txtMeetingParentType.Text = "Items:" + _itemscount + ",TotalTime(h):" + totaltime + ",Tips:" + _tipscount + ",SaveTime(h):" +
                           _savetime + ",PCT(%):" + p.CalcPCT(_savetime, totaltime);


                        grbMeetingChildNode.Text = workdetail;
                        txtMeetingParentTotalTime.Text = totaltime.ToString();
                        //tips count 
                        sql = "SELECT SUM(tips) FROM t_meetingtips where meetingtype = '" + workdetail + "'";
                        txtMeetingAlreadyHaveTips.Text = p.querySum(sql).ToString();
                        //tips save time
                        sql = "SELECT SUM(tipsavetime) FROM t_meetingtips where meetingtype = '" + workdetail + "'";
                        txtMeetingHaveTipsSaveTime.Text = p.querySum(sql).ToString();
                        txtMeetingHaveTipsOptimizePCT.Text = p.CalcPCT(Convert.ToDecimal(txtMeetingHaveTipsSaveTime.Text.Trim()), totaltime);
                        txtMeetingHaveTipsOptimizePCTTotal.Text = p.CalcPCT(Convert.ToDecimal(txtMeetingHaveTipsSaveTime.Text.Trim()), totalworktime);


                    }

                }
                else //childnode
                {
                    if (worktype == p.WorkType.Report)
                    {
                        grbReportChildNode.Text = workdetail;
                        txtReportParentTotalTime.Text = totaltime.ToString();
                        //tips count 
                        sql = "SELECT SUM(tips) FROM t_reporttips where workdetail = '" + workdetail + "'";
                        txtReportAlreadyHaveTips.Text = p.querySum(sql).ToString();
                        //tips save time
                        sql = "SELECT SUM(tipsavetime) FROM t_reporttips where workdetail = '" + workdetail + "'";
                        txtReportHaveTipsSaveTime.Text = p.querySum(sql).ToString();
                        txtReportHaveTipsOptimizePCT.Text = p.CalcPCT(Convert.ToDecimal(txtReportHaveTipsSaveTime.Text.Trim()), totaltime);
                        txtReportHaveTipsOptimizePCTTotal.Text = p.CalcPCT(Convert.ToDecimal(txtReportHaveTipsSaveTime.Text.Trim()), totalworktime);
                        string lastdate = "";
                        p.queryData("SELECT * FROM t_reporttips WHERE workdetail = '" + workdetail + "'", "reviewdate", out lastdate);
                        txtReportLastUpdateTime.Text = lastdate;

                        //reviewer,description,method,
                        string _reviewer, _description, _method, _duedate, _status;

                        sql = "SELECT reviewer FROM t_reporttips WHERE workdetail = '" + workdetail + "'";
                        p.queryData(sql, "reviewer", out _reviewer);
                        txtReportOldReviwer.Text = _reviewer;

                        sql = "SELECT description FROM t_reporttips WHERE workdetail = '" + workdetail + "'";
                        p.queryData(sql, "description", out _description);
                        txtReportOldDescription.Text = _description;

                        sql = "SELECT optimizemethod FROM t_reporttips WHERE workdetail = '" + workdetail + "'";
                        p.queryData(sql, "optimizemethod", out _method);
                        txtReportOptimizeMethod.Text = _method;

                        sql = "SELECT duedate FROM t_reporttips WHERE workdetail = '" + workdetail + "'";
                        p.queryData(sql, "duedate", out _duedate);
                        txtReportOldDueDate.Text = _duedate;

                        sql = "SELECT status FROM t_reporttips WHERE workdetail = '" + workdetail + "'";
                        p.queryData(sql, "status", out _status);
                        txtReportTipsStatus.Text = _status;
                    }

                    if (worktype == p.WorkType.Meeting)
                    {
                        grbMeetingChildNode.Text = workdetail;
                        txtMeetingParentTotalTime.Text = totaltime.ToString();
                        //tips count 
                        sql = "SELECT SUM(tips) FROM t_meetingtips where workdetail = '" + workdetail + "'";
                        txtMeetingAlreadyHaveTips.Text = p.querySum(sql).ToString();
                        //tips save time
                        sql = "SELECT SUM(tipsavetime) FROM t_meetingtips where workdetail = '" + workdetail + "'";
                        txtMeetingHaveTipsSaveTime.Text = p.querySum(sql).ToString();
                        txtMeetingHaveTipsOptimizePCT.Text = p.CalcPCT(Convert.ToDecimal(txtMeetingHaveTipsSaveTime.Text.Trim()), totaltime);
                        txtMeetingHaveTipsOptimizePCTTotal.Text = p.CalcPCT(Convert.ToDecimal(txtMeetingHaveTipsSaveTime.Text.Trim()), totalworktime);
                        string lastdate = "";
                        p.queryData("SELECT * FROM t_meetingtips WHERE workdetail = '" + workdetail + "'", "reviewdate", out lastdate);
                        txtMeetingLastUpdateTime.Text = lastdate;

                        //reviewer,description,method
                        string _reviewer, _description, _method, _duedate, _status;

                        sql = "SELECT reviewer FROM t_meetingtips WHERE workdetail = '" + workdetail + "'";
                        p.queryData(sql, "reviewer", out _reviewer);
                        txtMeetingOldReviwer.Text = _reviewer;

                        sql = "SELECT description FROM t_meetingtips WHERE workdetail = '" + workdetail + "'";
                        p.queryData(sql, "description", out _description);
                        txtMeetingOldDescription.Text = _description;

                        sql = "SELECT optimizemethod FROM t_meetingtips WHERE workdetail = '" + workdetail + "'";
                        p.queryData(sql, "optimizemethod", out _method);
                        txtMeetingOptimizeMethod .Text  = _method;

                        sql = "SELECT duedate FROM t_meetingtips WHERE workdetail = '" + workdetail + "'";
                        p.queryData(sql, "duedate", out _duedate);
                        txtMeetingOldDueDate.Text = _duedate;

                        sql = "SELECT status FROM t_meetingtips WHERE workdetail = '" + workdetail + "'";
                        p.queryData(sql, "status", out _status);
                        txtMeetingTipsStatus.Text = _status;

                    }
                }
         
              
            }

            //MessageBox.Show(_bSelectParentNode.ToString());
                 
        }

        private void setListview(ListView listview, p.WorkType worktype)
        {
            listview.MultiSelect = false;
            listview.AutoArrange = true;
            listview.GridLines = true;
            listview.FullRowSelect = true;
            listview.Columns.Add("ID", 60, HorizontalAlignment.Center);            
            listview.Columns.Add("Sec.Code", 60, HorizontalAlignment.Center);
            listview.Columns.Add("OPID", 60, HorizontalAlignment.Center);
            listview.Columns.Add("Eng.Name", 80, HorizontalAlignment.Center);
            //if (worktype == p.WorkType.Meeting )
               // listview.Columns.Add("Meeting Type", 80, HorizontalAlignment.Center);
           // if (worktype == p.WorkType.Report)
               // listview.Columns.Add("Report Type", 80, HorizontalAlignment.Center);
            listview.Columns.Add("Work Content", 120, HorizontalAlignment.Center);
            listview.Columns.Add("Work Detail", 150, HorizontalAlignment.Center);
            listview.Columns.Add("单次工作时间(h)", 80, HorizontalAlignment.Center);
            listview.Columns.Add("周工作频率(次)", 80, HorizontalAlignment.Center);
            listview.Columns.Add("周工作时间(h)", 80, HorizontalAlignment.Center);
            listview.Columns.Add("月工作时间(h)", 80, HorizontalAlignment.Center);
        }

        private void loadData2ListView(ListView listview, p.WorkType worktype,string workdetail,bool selectparentnode)
        {
            listview.Items.Clear();
            listview.BeginUpdate();//数据更新，UI暂时挂起，直到EndUpdate绘制控件，可以有效避免闪烁并大大提高加载速度 
            //
            string sql = "";
            if (worktype == p.WorkType.Meeting)
            {
                if (selectparentnode)
                    sql = "SELECT * FROM t_meetingrawdata WHERE meetingtype = '" + workdetail + "' order by seccode ASC";
                else
                    sql = "SELECT * FROM t_meetingrawdata WHERE workdetail = '" + workdetail + "' order by seccode ASC";
            }
            if (worktype == p.WorkType .Report)
            {
                if (selectparentnode)
                    sql = "SELECT * FROM t_reportrawdata WHERE reporttype = '" + workdetail + "' order by seccode ASC";
                else
                    sql = "SELECT * FROM t_reportrawdata WHERE workdetail = '" + workdetail + "' order by seccode ASC";
            }
            //MessageBox.Show(sql);

            SQLiteConnection conn = new SQLiteConnection(p.dbConnectionString);
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            SQLiteDataReader re = cmd.ExecuteReader();
            ListViewItem lt = new ListViewItem();
            if (re.HasRows)
            {
                while (re.Read())
                {
                    p.SetListItemFont(lt, 9);
                    lt = listview.Items.Add(re["id"].ToString());
                    p.SetListItemFont(lt, 9);
                    string _seccode = re["seccode"].ToString();
                    lt.SubItems.Add(re["seccode"].ToString());
                    if (_seccode.ToUpper() == "KD1210" || _seccode.ToUpper() == "KD1220")
                        lt.ForeColor = Color.FromArgb(255, 51, 156, 150);
                    if (_seccode.ToUpper() == "KD1230")
                        lt.ForeColor = Color.FromArgb(255, 201, 137, 100);

                    lt.SubItems.Add(re["opid"].ToString());
                    lt.SubItems.Add(re["engname"].ToString());
                    //if (worktype == p.WorkType.Report)
                       // lt.SubItems.Add(re["reporttype"].ToString());
                    //if (worktype == p.WorkType.Meeting)
                       // lt.SubItems.Add(re["meetingtype"].ToString());
                    lt.SubItems.Add(re["workcontent"].ToString());
                    lt.SubItems.Add (re["workdetail"].ToString());
                    lt.SubItems.Add(re["singleworktime"].ToString());
                    lt.SubItems.Add(re["weeklyworkfreq"].ToString());
                    lt.SubItems.Add(re["weeklyworktime"].ToString());
                    lt.SubItems.Add(re["monthlyworktime"].ToString());
                }


            }
            conn.Close();

            lt = listview.Items.Add("Total");
            lt.SubItems.Add("");
            lt.SubItems.Add("");
            lt.SubItems.Add("");
            lt.SubItems.Add("");
            lt.SubItems.Add("");
            lt.SubItems.Add("");
            if (worktype == p.WorkType.Meeting )
            {
                if (selectparentnode )
                    sql = "SELECT SUM (weeklyworktime) FROM t_meetingrawdata WHERE meetingtype = '" + workdetail + "'";
                else
                    sql = "SELECT SUM (weeklyworktime) FROM t_meetingrawdata WHERE workdetail = '" + workdetail + "'";
            }

            if (worktype == p.WorkType.Report )
            {
                if (selectparentnode)
                    sql = "SELECT SUM (weeklyworktime) FROM t_reportrawdata WHERE reporttype = '" + workdetail + "'";
                else
                    sql = "SELECT SUM (weeklyworktime) FROM t_reportrawdata WHERE workdetail = '" + workdetail + "'";
            }
            //MessageBox.Show(sql);

            lt.SubItems.Add(p.querySum(sql).ToString());
            lt.ForeColor = Color.FromArgb(255, 68,140, 211);
            p.SetListItemFont(lt,9);
           

            listview.EndUpdate();//结束数据处理，UI界面一次性绘制。

            
            
        }

        private void trviewMeeting_AfterSelect(object sender, TreeViewEventArgs e)
        {

            loadInfo2toolStrip(p.WorkType.Meeting, lstviewMeeting, trviewMeeting);
        }


        private DataTable queryRawData(p.WorkType worktype)
        {
            DataTable dt = new DataTable();
            string sql = "";
            if (worktype == p.WorkType.Meeting)
                sql = "SELECT * FROM t_meetingrawdata";
            if (worktype == p.WorkType.Report)
                sql = "SELECT * FROM t_reportrawdata";

            SQLiteConnection conn = new SQLiteConnection(p.dbConnectionString);
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            SQLiteDataReader re = cmd.ExecuteReader();

            try
            {
                dt.Load(re);
            }
            catch (Exception)
            {
                
                //throw;
            }
            
            conn.Close();
            return dt;
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            loadRawdata2DataGrid();
        }


        private void loadRawdata2DataGrid()
        {
            string sql = "";
            int icount = -1;
            decimal totaltime = 0;
            p.WorkType worktype = (p.WorkType)Enum.Parse(typeof(p.WorkType), comboRawDataType.Text);
            //
            if (worktype == p.WorkType.Report)
                sql = "SELECT COUNT(*) FROM t_meetingrawdata";
            if (worktype == p.WorkType.Meeting)
                sql = "SELECT COUNT(*) FROM t_reportrawdata";
            icount = p.queryCount(sql);
            //
            if (worktype == p.WorkType.Report)
                sql = "SELECT SUM(weeklyworktime) FROM t_meetingrawdata";
            if (worktype == p.WorkType.Meeting)
                sql = "SELECT SUM(weeklyworktime) FROM t_reportrawdata";
            totaltime = p.querySum(sql);
            DataTable dt = queryRawData(worktype);
            datagridRawData.DataSource = dt;
            datagridRawData.ReadOnly = true;
            //AutoSizeColumn(datagridRawData);

            tsslStatus.ForeColor = Color.Blue;
            tsslStatus.Text = worktype.ToString() + ",there is " + icount + " records in database,total weekly worktime(h):" + totaltime;
        }


        /// <summary>
        /// 使DataGridView的列自适应宽度
        /// </summary>
        /// <param name="dgViewFiles"></param>
        private void AutoSizeColumn(DataGridView dgViewFiles)
        {
            int width = 0;
            //使列自使用宽度
            //对于DataGridView的每一个列都调整
            for (int i = 0; i < dgViewFiles.Columns.Count; i++)
            {
                //将每一列都调整为自动适应模式
                dgViewFiles.AutoResizeColumn(i, DataGridViewAutoSizeColumnMode.AllCells);
                //记录整个DataGridView的宽度
                width += dgViewFiles.Columns[i].Width;
            }
            //判断调整后的宽度与原来设定的宽度的关系，如果是调整后的宽度大于原来设定的宽度，
            //则将DataGridView的列自动调整模式设置为显示的列即可，
            //如果是小于原来设定的宽度，将模式改为填充。
            if (width > dgViewFiles.Size.Width)
            {
                dgViewFiles.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            }
            else
            {
                dgViewFiles.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            //冻结某列 从左开始 0，1，2
            dgViewFiles.Columns[1].Frozen = true;
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr = MessageBox.Show("Are u sure to exit?", "Exit or Not", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
                Environment.Exit(0);
            if (dr == DialogResult.No)
                e.Cancel = true;
        }

        private void btnSaveReport_Click(object sender, EventArgs e)
        {
            //check can't empty
            if (string.IsNullOrEmpty(txtReportNewTips.Text.Trim()))
            {
                MessageBox.Show("Tips qty. can't be empty.", "Number Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtReportNewTips.SelectAll();
                txtReportNewTips.Focus();
                return;
            }
            if (string.IsNullOrEmpty( txtReportNewTipsSaveTime .Text .Trim ()))
            {
                MessageBox.Show("Tips Save time can't be empty.", "Number Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtReportNewTipsSaveTime.SelectAll();
                txtReportNewTipsSaveTime.Focus();
                return;
            }

            if (string.IsNullOrEmpty(txtReportNewDescription.Text.Trim()))
            {
                MessageBox.Show("Tips description can't be empty.", "Number Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtReportNewDescription.SelectAll();
                txtReportNewDescription.Focus();
                return;
            }

            if (this.comboReportOptimizeMethod.SelectedIndex == -1)
            {
                MessageBox.Show("Pls select meeting optimize method", "DONT SELECT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.comboReportOptimizeMethod.Focus();
                return;
            }

            if (this.comboReportTipsStatus.SelectedIndex == -1)
            {
                MessageBox.Show("Pls select current tips status", "DONT SELECT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.comboReportTipsStatus.Focus();
                return;
            }





            //check range
            decimal totaltime = Convert.ToDecimal(txtReportParentTotalTime.Text.Trim());
            decimal lastsave = Convert.ToDecimal(txtReportHaveTipsSaveTime.Text.Trim());
            decimal newsave = Convert.ToDecimal(txtReportNewTipsSaveTime.Text.Trim());
            if ((lastsave + newsave) > totaltime)
            {
                MessageBox.Show("当前改善的时间(" + newsave + ")同已改善的时间(" + lastsave + ")总和已大于当前项目的总时间和(" + totaltime + "),请重新check", "Out Of Range", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtReportNewTipsSaveTime.SelectAll();
                txtReportNewTipsSaveTime.Focus();
                return;
            }
            //check datetime
            if (!string.IsNullOrEmpty (txtReportLastUpdateTime .Text.Trim ()))
            {
                if (checkDatetimeIsNow(txtReportLastUpdateTime.Text.Trim(), DateTime.Now.ToString("yyyy-MM-dd")))
                {
                    DialogResult dr = MessageBox.Show("当前更新时间等于或者小于上一次更新的时间,确认是否需要更新,更新点YES,不更新点NO", "Qestion", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == System.Windows.Forms.DialogResult.No)
                    {
                        return;
                    }
                }   
            }
            
            this.Enabled = false;
            int _tips = Convert.ToInt16(txtReportAlreadyHaveTips.Text.Trim()) + Convert.ToInt16(txtReportNewTips.Text.Trim());
            decimal _tipsavetime = Convert.ToDecimal(txtReportHaveTipsSaveTime.Text.Trim()) + Convert.ToDecimal(txtReportNewTipsSaveTime .Text.Trim() );
            string _reporttype = grbReportParentNode.Text;
            string _optimizemethod = this.comboReportOptimizeMethod.SelectedItem.ToString();
            string _description = txtReportNewDescription.Text.Trim();
            p.replaceString(ref  _description );
            string _reviewer = txtReportNewReviewer.Text.Trim();
            string _duedate = dtpReportDueDate.Value.ToString("yyyy-MM-dd");
            string _status = this.comboReportTipsStatus.SelectedItem.ToString ();


            string sql = @"REPLACE INTO t_reporttips  VALUES ('" +
                grbReportChildNode.Text + "','" +
                _reporttype  + "','" +
                _tips + "','" +
                _tipsavetime  + "','" +
                _optimizemethod +"','" +
                _description +"','" +
                _duedate +"','" +
                _status +"','"+
                _reviewer +"','"+
                DateTime.Now.ToString("yyyy-MM-dd") + "')";

            if (p.updateData2DB(sql))
            {
                MessageBox.Show("update date into database success...", "Update Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                loadInfo2toolStrip(p.WorkType.Report, lstviewReport, trviewReport);
                //OutputData2Text(p.WorkType.Report);
                loadTreeViewData(trviewReport, p.WorkType.Report);
             }
            this.Enabled = true;

            

        }

        /// <summary>
        /// 比较两个时间,如果现在的日期大于上次的日期, trun false,反之retun true
        /// </summary>
        /// <param name="lastdt"></param>
        /// <param name="currentdt"></param>
        /// <returns></returns>
        private bool checkDatetimeIsNow(string  lastdt,string  currentdt)
        {
            DateTime lastdate  = DateTime.ParseExact(lastdt , "yyyy-MM-dd", System.Globalization.CultureInfo.CurrentCulture);
            DateTime currentdate = DateTime.ParseExact(currentdt, "yyyy-MM-dd", System.Globalization.CultureInfo.CurrentCulture);

            if (currentdate > lastdate )
                return false;

            return true;
        }

        private void txtNewTips_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNumberInput(e);
        }


        private void onlyNumberInput(KeyPressEventArgs e)
        {
            if (!(char.IsNumber(e.KeyChar)) && e.KeyChar != (char)8 && e.KeyChar != (char)45)
                e.Handled = true;
            else
                e.Handled = false;
        }

        private void onlyDecimalInput(KeyPressEventArgs e)
        {
            if (!(char.IsNumber(e.KeyChar)) && e.KeyChar != (char)8 && e.KeyChar !=(char)46 && e.KeyChar !=(char) 45)
                e.Handled = true;
            else
                e.Handled = false;
        }

        private void txtNewTipsSaveTime_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyDecimalInput(e);
        }

        private void txtReportNewTipsSaveTime_TextChanged(object sender, EventArgs e)
        {
            try
            {
                txtReportNewTipsOptimizePCT.Text = p.CalcPCT(Convert.ToDecimal(txtReportNewTipsSaveTime.Text.Trim()), Convert.ToDecimal(txtReportParentTotalTime.Text.Trim()));
                txtReportNewTipsOptimizePCTTotal.Text = p.CalcPCT(Convert.ToDecimal(txtReportNewTipsSaveTime.Text.Trim()), Convert.ToDecimal(txtReportTotalTime.Text.Trim()));
            }
            catch (Exception)
            {

                //throw;
            }
            
        }

        private void btnSaveMeeting_Click(object sender, EventArgs e)
        {

           
            //check can't empty
            if (string.IsNullOrEmpty(txtMeetingNewTips.Text.Trim()))
            {
                MessageBox.Show("Tips qty. can't be empty.", "Number Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMeetingNewTips.SelectAll();
                txtMeetingNewTips.Focus();
                return;
            }
            if (string.IsNullOrEmpty(txtMeetingNewTipsSaveTime.Text.Trim()))
            {
                MessageBox.Show("Tips Save time can't be empty.", "Number Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMeetingNewTipsSaveTime.SelectAll();
                txtMeetingNewTipsSaveTime.Focus();
                return;
            }

            if (string.IsNullOrEmpty(txtMeetingNewDescription.Text.Trim()))
            {
                MessageBox.Show("Tips description can't be empty.", "Number Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMeetingNewDescription.SelectAll();
                txtMeetingNewDescription.Focus();
                return;
            }

            if (this.comboMeetingOptimizeMethod.SelectedIndex == -1)
            {
                MessageBox.Show("Pls select meeting optimize method", "DONT SELECT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.comboMeetingOptimizeMethod.Focus();
                return;
            }

            if (this.comboMeetingTipsStatus  .SelectedIndex == -1)
            {
                MessageBox.Show("Pls select current tips status", "DONT SELECT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.comboMeetingTipsStatus.Focus();
                return;
            }


            //check range
            decimal totaltime = Convert.ToDecimal(txtMeetingParentTotalTime.Text.Trim());
            decimal lastsave = Convert.ToDecimal(txtMeetingHaveTipsSaveTime.Text.Trim());
            decimal newsave = Convert.ToDecimal(txtMeetingNewTipsSaveTime.Text.Trim());
            if ((lastsave + newsave) > totaltime)
            {
                MessageBox.Show("当前改善的时间(" + newsave + ")同已改善的时间(" + lastsave + ")总和已大于当前项目的总时间和(" + totaltime + "),请重新check", "Out Of Range", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMeetingNewTipsSaveTime.SelectAll();
                txtMeetingNewTipsSaveTime.Focus();
                return;
            }
            //check datetime
            if (!string.IsNullOrEmpty(txtMeetingLastUpdateTime.Text.Trim()))
            {
                if (checkDatetimeIsNow(txtMeetingLastUpdateTime.Text.Trim(), DateTime.Now.ToString("yyyy-MM-dd")))
                {
                    DialogResult dr = MessageBox.Show("当前更新时间等于或者小于上一次更新的时间,确认是否需要更新,更新点YES,不更新点NO", "Qestion", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == System.Windows.Forms.DialogResult.No)
                    {
                        return;
                    }
                }
            }

            this.Enabled = false;
            int _tips = Convert.ToInt16(txtMeetingAlreadyHaveTips.Text.Trim()) + Convert.ToInt16(txtMeetingNewTips.Text.Trim());
            decimal _tipsavetime = Convert.ToDecimal(txtMeetingHaveTipsSaveTime.Text.Trim()) + Convert.ToDecimal(txtMeetingNewTipsSaveTime.Text.Trim());

            string _meetingtype = grbMeetingParentNode.Text;
            string _optimizemethod = this.comboMeetingOptimizeMethod.SelectedItem.ToString();
            string _description = txtMeetingNewDescription.Text.Trim();
            p.replaceString(ref _description);
            string _reviewer = txtMeetingNewReviewer.Text.Trim();
            string _duedate = dtpMeetingDueDate.Value.ToString("yyyy-MM-dd");
            string _status = this.comboMeetingTipsStatus.SelectedItem.ToString();


            string sql = @"REPLACE INTO t_meetingtips  VALUES ('" +
                grbMeetingChildNode.Text + "','" +
                _meetingtype + "','" +
                _tips + "','" +
                _tipsavetime + "','" +
                _optimizemethod + "','" +
                _description + "','" +
                _duedate + "','" +
                _status + "','" +
                _reviewer + "','" +
                DateTime.Now.ToString("yyyy-MM-dd") + "')";

            if (p.updateData2DB(sql))
            {
                MessageBox.Show("update date into database success...", "Update Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                loadInfo2toolStrip(p.WorkType.Meeting, lstviewMeeting, trviewMeeting);
                loadTreeViewData(trviewMeeting, p.WorkType.Meeting);
                trviewMeeting.Sort();
            }
            this.Enabled = true;
        }

        private void txtMeetingNewTips_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNumberInput(e);
        }

        private void txtMeetingNewTipsSaveTime_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyDecimalInput(e);
        }

        private void txtMeetingNewTipsSaveTime_TextChanged(object sender, EventArgs e)
        {
            try
            {
                txtMeetingNewTipsOptimizePCT.Text = p.CalcPCT(Convert.ToDecimal(txtMeetingNewTipsSaveTime.Text.Trim()), Convert.ToDecimal(txtMeetingParentTotalTime.Text.Trim()));
                txtMeetingNewTipsOptimizePCTTotal.Text = p.CalcPCT(Convert.ToDecimal(txtMeetingNewTipsSaveTime.Text.Trim()), Convert.ToDecimal(txtMeetingTotalTime.Text.Trim()));
            }
            catch (Exception)
            {

                //throw;
            }
        }


        private void OutputData2Text(p.WorkType worktype)
        {
            p.checkLogFile(worktype);

            //
            string filepath, _item, _depcode, _subtype, _workdetail, _type, _itemscount, _workingtime, _weeklyfreq, _weeklyworkingtime, _monthlyworkingtime, _optimizemethod, _tips, _savetime, _savepct, _updatedate, _description, _duedate, _status;
            filepath = _item = _depcode = _subtype = _workdetail = _type = _itemscount = _workingtime = _weeklyfreq = _weeklyworkingtime = _monthlyworkingtime = _optimizemethod = _tips = _savetime = _savepct = _updatedate = _description = _duedate = _status = "";
            //
            _type= worktype.ToString();
            if (worktype == p.WorkType.Report)
            {
                filepath = p.logReportFile;
                if (trviewReport.Nodes.Count > 0)
                {
                    string sql = "SELECT depcode FROM t_reportrawdata where id = '1'";
                    p.queryData(sql, "depcode", out _depcode);


                    for (int i = 0; i < trviewReport.Nodes.Count ; i++)
                    {
                        _item = (i + 1).ToString();//
                        //_subtype = trviewReport.Nodes[i].Text.Substring(0, trviewReport.Nodes[i].Text.IndexOf(';'));
                        _itemscount = trviewReport.Nodes[i].Nodes.Count.ToString();//

                        spliteParentInfo(trviewReport.Nodes[i].Text,
                            out _subtype, out _weeklyworkingtime, out _monthlyworkingtime,
                            out _tips,out  _savetime, out  _savepct);


                        // sql = "SELECT SUM(weeklyworktime) FROM t_reportrawdata WHERE reporttype = '" + _subtype + "'";
                        //_weeklyworkingtime = p.querySum(sql).ToString();

                        //sql = "SELECT SUM(monthlyworktime) FROM t_reportrawdata WHERE reporttype = '" + _subtype + "'";
                        //_monthlyworkingtime = p.querySum(sql).ToString();

                        //sql = "SELECT SUM(tips) FROM t_reporttips WHERE reporttype = '" + _subtype + "'";
                        //_tips = p.querySum(sql).ToString();
                        //sql = "SELECT SUM(tipsavetime) FROM t_reporttips WHERE reporttype = '" + _subtype + "'";
                        //_savetime = p.querySum(sql).ToString();                        
                        //_savepct = p.CalcPCT(Convert.ToDecimal(_savetime), Convert.ToDecimal(_weeklyworkingtime));

                        //sql = "SELECT optimizemethod FROM t_reporttips WHERE reporttype = '" + _subtype + "'";
                        //p.queryData(sql, "optimizemethod", out  _optimizemethod);

                        //sql = "SELECT duedate FROM t_reporttips WHERE reporttype = '" + _subtype + "'";
                        //p.queryData(sql, "duedate", out _duedate);

                        //sql = "SELECT description FROM t_reporttips WHERE reporttype = '" + _subtype + "'";
                        //p.queryData(sql, "description", out _description);

                        //sql = "SELECT status FROM t_reporttips WHERE reporttype = '" + _subtype + "'";
                        //p.queryData(sql, "description", out _status);
                        ////////

                        saveLog(filepath, _item, _depcode, _subtype, _workdetail, _type,
                            _itemscount, _workingtime, _weeklyfreq, _weeklyworkingtime, _monthlyworkingtime,
                            _optimizemethod, _tips, _savetime, _savepct, _updatedate, _description,
                            _duedate, _status);

                        tsslStatus.Text = "正在保存文件," + (i + 1) + "-0";
                        Application.DoEvents();
                        for (int j = 0; j < trviewReport.Nodes[i].Nodes.Count ; j++)
                        {
                            _item = _subtype = "";
                            _workdetail = trviewReport.Nodes[i].Nodes[j].Text;

                            sql = "SELECT COUNT(workdetail) FROM t_reportrawdata WHERE workdetail = '" + _workdetail + "'";
                            _itemscount = p.queryCount(sql).ToString();

                            sql = "SELECT SUM(weeklyworktime) FROM t_reportrawdata WHERE workdetail = '" + _workdetail + "'";
                            _weeklyworkingtime = p.querySum(sql).ToString();

                            sql = "SELECT SUM(monthlyworktime) FROM t_reportrawdata WHERE workdetail = '" + _workdetail + "'";
                            _monthlyworkingtime = p.querySum(sql).ToString();

                            sql = "SELECT SUM(tips) FROM t_reporttips WHERE workdetail = '" + _workdetail + "'";
                            _tips = p.querySum(sql).ToString();
                            sql = "SELECT SUM(tipsavetime) FROM t_reporttips WHERE workdetail = '" + _workdetail + "'";
                            _savetime = p.querySum(sql).ToString();
                            _savepct = p.CalcPCT(Convert.ToDecimal(_savetime), Convert.ToDecimal( _weeklyworkingtime));

                            sql = "SELECT optimizemethod FROM t_reporttips WHERE workdetail = '" + _workdetail + "'";
                            p.queryData(sql, "optimizemethod", out  _optimizemethod);

                            sql = "SELECT duedate FROM t_reporttips WHERE workdetail= '" + _workdetail + "'";
                            p.queryData(sql, "duedate", out _duedate);

                            sql = "SELECT description FROM t_reporttips WHERE workdetail = '" + _workdetail + "'";
                            p.queryData(sql, "description", out _description);

                            sql = "SELECT status FROM t_reporttips WHERE workdetail = '" + _workdetail + "'";
                            p.queryData(sql, "description", out _status);

                            p.replaceString(ref _workdetail);

                            saveLog(filepath, _item, _depcode, _subtype, _workdetail, _type,
                           _itemscount, _workingtime, _weeklyfreq, _weeklyworkingtime, _monthlyworkingtime,
                           _optimizemethod, _tips, _savetime, _savepct, _updatedate, _description,
                           _duedate, _status);
                            tsslStatus.Text = "正在保存文件," + (i + 1) + "-" + (j + 1);
                            Application.DoEvents();
                        }
                        
                    }
                }
            }

            if (worktype == p.WorkType.Meeting)
            {
                filepath = p.logMeetingFile;
                if (trviewMeeting.Nodes.Count > 0)
                {
                    string sql = "SELECT depcode FROM t_meetingrawdata where id = '1'";
                    p.queryData(sql, "depcode", out _depcode);

                    for (int i = 0; i < trviewMeeting.Nodes.Count; i++)
                    {
                        _item = (i + 1).ToString();//
                       // _subtype = trviewMeeting.Nodes[i].Text.Substring(0, trviewMeeting.Nodes[i].Text.IndexOf(';'));
                        _itemscount = trviewMeeting.Nodes[i].Nodes.Count.ToString();//

                        spliteParentInfo(trviewMeeting.Nodes[i].Text,
                            out _subtype, out _weeklyworkingtime, out _monthlyworkingtime,
                            out _tips, out  _savetime, out  _savepct);

                        //sql = "SELECT SUM(weeklyworktime) FROM t_meetingrawdata WHERE meetingtype = '" + _subtype + "'";
                        //_weeklyworkingtime = p.querySum(sql).ToString();

                        //sql = "SELECT SUM(monthlyworktime) FROM t_meetingrawdata WHERE meetingtype = '" + _subtype + "'";
                        //_monthlyworkingtime = p.querySum(sql).ToString();

                        //sql = "SELECT SUM(tips) FROM t_meetingtips WHERE meetingtype = '" + _subtype + "'";
                        //_tips = p.querySum(sql).ToString();
                        //sql = "SELECT SUM(tipsavetime) FROM t_meetingtips WHERE meetingtype = '" + _subtype + "'";
                        //_savetime = p.querySum(sql).ToString();
                        //_savepct = p.CalcPCT(Convert.ToDecimal(_savetime), Convert.ToDecimal(_workingtime));

                        //sql = "SELECT optimizemethod FROM t_meetingtips WHERE meetingtype = '" + _subtype + "'";
                        //p.queryData(sql, "optimizemethod", out  _optimizemethod);

                        //sql = "SELECT duedate FROM t_meetingtips WHERE meetingtype = '" + _subtype + "'";
                        //p.queryData(sql, "duedate", out _duedate);

                        //sql = "SELECT description FROM t_meetingtips WHERE meetingtype = '" + _subtype + "'";
                        //p.queryData(sql, "description", out _description);

                        //sql = "SELECT status FROM t_meetingtips WHERE meetingtype = '" + _subtype + "'";
                        //p.queryData(sql, "description", out _status);

                        saveLog(filepath, _item, _depcode, _subtype, _workdetail, _type,
                            _itemscount, _workingtime, _weeklyfreq, _weeklyworkingtime, _monthlyworkingtime,
                            _optimizemethod, _tips, _savetime, _savepct, _updatedate, _description,
                            _duedate, _status);

                        tsslStatus.Text = "正在保存文件," + (i + 1) + "-0";
                        Application.DoEvents();
                        for (int j = 0; j < trviewMeeting.Nodes[i].Nodes.Count; j++)
                        {
                            _item = _subtype = "";
                            _workdetail = trviewMeeting.Nodes[i].Nodes[j].Text;

                            sql = "SELECT COUNT(workdetail) FROM t_meetingrawdata WHERE workdetail = '" + _workdetail + "'";
                            _itemscount = p.queryCount(sql).ToString();

                            sql = "SELECT SUM(weeklyworktime) FROM t_meetingrawdata WHERE workdetail = '" + _workdetail + "'";
                            _weeklyworkingtime = p.querySum(sql).ToString();

                            sql = "SELECT SUM(monthlyworktime) FROM t_meetingrawdata WHERE workdetail = '" + _workdetail + "'";
                            _monthlyworkingtime = p.querySum(sql).ToString();

                            sql = "SELECT SUM(tips) FROM t_meetingtips WHERE workdetail = '" + _workdetail + "'";
                            _tips = p.querySum(sql).ToString();
                            sql = "SELECT SUM(tipsavetime) FROM t_meetingtips WHERE workdetail = '" + _workdetail + "'";
                            _savetime = p.querySum(sql).ToString();
                            _savepct = p.CalcPCT(Convert.ToDecimal(_savetime), Convert.ToDecimal(_weeklyworkingtime ));

                            sql = "SELECT optimizemethod FROM t_meetingtips WHERE workdetail = '" + _workdetail + "'";
                            p.queryData(sql, "optimizemethod", out  _optimizemethod);

                            sql = "SELECT duedate FROM t_meetingtips WHERE workdetail = '" + _workdetail + "'";
                            p.queryData(sql, "duedate", out _duedate);

                            sql = "SELECT description FROM t_meetingtips WHERE workdetail = '" + _workdetail + "'";
                            p.queryData(sql, "description", out _description);

                            sql = "SELECT status FROM t_meetingtips WHERE workdetail = '" + _workdetail + "'";
                            p.queryData(sql, "description", out _status);

                            saveLog(filepath, _item, _depcode, _subtype, _workdetail, _type,
                           _itemscount, _workingtime, _weeklyfreq, _weeklyworkingtime, _monthlyworkingtime,
                           _optimizemethod, _tips, _savetime, _savepct, _updatedate, _description,
                           _duedate, _status);

                            tsslStatus.Text = "正在保存文件," + (i + 1) + "-" + (j + 1);
                            Application.DoEvents();
                        }

                    }
                }
            }


            MessageBox.Show("Save OK,file is " + filepath);
            tsslStatus.Text = "";

        }





        private void saveLog(string filepath,
            string _item,string _depcode,string _subtype,
            string _workdetail,string _type,string _itemscount,
            string _workingtime,string _weeklyfreq,string _weeklyworkingtime,
            string _monthlyworkingtime,string _optimizeMethod,string _tips,
            string _savetime,string _savepct,string _updatedate,string _description,
            string _duedate,string _status)

        {
            string line = _item.PadRight(6) + "*" + _depcode.PadRight(6) + "*" +
                _subtype.PadRight(40) + "*" + _workdetail.PadRight(100) + "*" +
                _type.PadRight(10) + "*" +
                _itemscount +"*"+_workingtime +"*"+   _weeklyfreq.PadRight(15) + "*" +
                _weeklyworkingtime.PadRight(25) + "*" + _monthlyworkingtime.PadRight(30) + "*" +
                _optimizeMethod.PadRight(20) + "*" +
                _tips.PadRight(6) + "*" + _savetime.PadRight(10) + "*" +
                 _savepct.PadRight(15) + "*" + _updatedate.PadRight(15) + "*" +
                 _description.PadRight(100) + "*" + _duedate.PadRight(15) + "*" +
                 _status.PadRight(10);                
            StreamWriter sw = new StreamWriter (filepath,true);
            sw.WriteLine(line);
            sw.Close();
            
        }





        private void spliteParentInfo(string parentstring,
            out string _parentnode,out string _weeklyworkingtime,out string _monthlyworkingtime,
            out string _tips, out string _savetime, out string _savepct)
        {
            _parentnode = "";
            _weeklyworkingtime = _monthlyworkingtime  =_tips = _savetime  = _savepct ="0";

            try
            {
                _parentnode = parentstring.Substring(0, parentstring.IndexOf(';'));
                string tailstring = parentstring.Replace(_parentnode, "");
                string[] s = tailstring.Split(',');
                foreach (string  item in s )
                {
                    if (item.ToUpper().Contains("TOTALTIME")) 
                        _weeklyworkingtime = item.ToUpper().Replace("TOTALTIME(H):","");
                    if (item.ToUpper().Contains("TIPS"))
                        _tips = item.ToUpper().Replace("TIPS:", "");
                     if (item.ToUpper().Contains("SAVETIME"))
                         _savetime = item.ToUpper().Replace("SAVETIME(H):", "");
                    if (item.ToUpper().Contains("PCT"))
                        _savepct = item.ToUpper().Replace("PCT(%):", "");
                    
                }
                _monthlyworkingtime = (Convert.ToDecimal(_weeklyworkingtime) * 4).ToString();
            }
            catch (Exception)
            {
                
                //pthrow;
            }



        }




        private void btnOutputMeetinig_Click(object sender, EventArgs e)
        {
            this.btnOutputMeetinig.Enabled = false;
            OutputData2Text(p.WorkType.Meeting);            
            this.btnOutputMeetinig.Enabled = true;
        }

        private void btnOutputReports_Click(object sender, EventArgs e)
        {
            //this.Enabled = false;
            this.btnOutputReports.Enabled = false;
            
            trviewReport.Sort();


            this.btnOutputReports.Enabled = true;
            //savelogworktype = p.WorkType.Report;
            //backgroundWorker1.RunWorkerAsync();
            //this.Enabled = true;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            //OutputData2Text(savelogworktype, this.backgroundWorker1);
        }




        //private void OutputData2Excel()
        //{
        //    string filePath = System.Windows.Forms.Application.StartupPath + @"\" + "WCD_ATE_AFTE_NTF_" + DateTime.Now.Date.AddDays(-1).ToString("yyyyMMdd") + @".xls";
        //    Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
        //    appExcel.Visible = false;
        //    Workbook wBook = appExcel.Workbooks.Add(true);
        //    Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
        //    wSheet.Name = "ATE NTF";



        //    //设置禁止弹出保存和覆盖的询问提示框   
        //    appExcel.DisplayAlerts = false;
        //    appExcel.AlertBeforeOverwriting = false;
        //    //保存工作簿   
        //    wBook.Save();
        //    //保存excel文件   
        //    appExcel.Save(filePath);
        //    appExcel.SaveWorkspace(filePath);
        //    appExcel.Quit();
        //    appExcel = null;


        //}



    }
}
