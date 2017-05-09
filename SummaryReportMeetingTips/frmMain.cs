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

            loadRawdata2DataGrid();


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
                            cmd.Parameters[4].Value = ds.Tables[combo.SelectedIndex].Rows[i]["会议类型"].ToString();
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
                    tr.Text = _nodename;
                    if (!nodeIsInTreView(tr, trview))
                        trview.Nodes.Add(tr);
                    //
                    int nodeindex = getNodeIndex(tr, trview) ;
                    if (nodeindex  != -1)
                    {
                        tr = trview.Nodes[nodeindex];
                        if (!childnodeIsInTreView (tr ,_childnodename))
                            tr.Nodes.Add(_childnodename);
                    }
                }
            }
            conn.Close();
         }
        

        private void tabMain_Selected(object sender, TabControlEventArgs e)
        {
            tsslStatus.Text = "";
            if (e.TabPage == tabReport )
            {
                loadTreeViewData(trviewReport, p.WorkType.Report );
                trviewReport.Sort();
                
            }
            if (e.TabPage == tabMeeting )
            {
                loadTreeViewData(trviewMeeting, p.WorkType.Meeting);
                trviewMeeting.Sort();
            }
            if (e.TabPage == tabRawData)
            {
                loadRawdata2DataGrid();
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
                    if (node.Nodes[i].Text.Trim().ToUpper() == childnode.Trim().ToUpper())
                        return true;
                }
            }
            return false;
        }

        private void trviewReport_AfterSelect(object sender, TreeViewEventArgs e)
        {
            loadInfo2toolStrip(p.WorkType.Report, lstviewReport, trviewReport);
        }






        private void loadInfo2toolStrip(p.WorkType worktype,ListView listview,TreeView treview)
        {
            bool _bSelectParentNode = false;
            listview.Items.Clear();
            string workdetail = "";
            string sql = "";
            try
            {
                workdetail = treview.SelectedNode.Parent.Text;
                workdetail = treview.SelectedNode.Text;
            }
            catch (Exception)
            {
                _bSelectParentNode = true;
                workdetail = treview.SelectedNode.Text;
            }

            //-------------------------------------
            if (!string.IsNullOrEmpty(workdetail))
            {
                int icount = -1;
                decimal totaltime = 0;
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
                    totaltime = p.querySum(sql);
                    loadData2ListView(listview, p.WorkType.Report, workdetail, _bSelectParentNode);
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
                }


                tsslStatus.ForeColor = Color.Blue;
                tsslStatus.Text = workdetail + " | itemscount:" + icount + " | weeklyworktime(h):" + totaltime;
            }





      
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
            listview.Columns.Add("Eng.Name", 90, HorizontalAlignment.Center);
            //if (worktype == p.WorkType.Meeting )
               // listview.Columns.Add("Meeting Type", 80, HorizontalAlignment.Center);
           // if (worktype == p.WorkType.Report)
               // listview.Columns.Add("Report Type", 80, HorizontalAlignment.Center);
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
            if (worktype == p.WorkType.Meeting && selectparentnode)
                sql = "SELECT * FROM t_meetingrawdata WHERE meetingtype = '" + workdetail + "' order by seccode ASC";
            else
                sql = "SELECT * FROM t_meetingrawdata WHERE workdetail = '" + workdetail + "' order by seccode ASC";

            if (worktype == p.WorkType .Report && selectparentnode)
                sql = "SELECT * FROM t_reportrawdata WHERE reporttype = '" + workdetail + "' order by seccode ASC";
            else
                sql = "SELECT * FROM t_reportrawdata WHERE workdetail = '" + workdetail + "' order by seccode ASC";

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
            if (worktype == p.WorkType.Meeting && selectparentnode )
                 sql = "SELECT SUM (weeklyworktime) FROM t_meetingrawdata WHERE meetingtype = '" + workdetail + "'";
            else
                sql = "SELECT SUM (weeklyworktime) FROM t_meetingrawdata WHERE workdetail = '" + workdetail + "'";

            if (worktype == p.WorkType.Report && selectparentnode )
                 sql = "SELECT SUM (weeklyworktime) FROM t_reportrawdata WHERE reporttype = '" + workdetail + "'";
            else
                sql = "SELECT SUM (weeklyworktime) FROM t_reportrawdata WHERE workdetail = '" + workdetail + "'";

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
    }
}
