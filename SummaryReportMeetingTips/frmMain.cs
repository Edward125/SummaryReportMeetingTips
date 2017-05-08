﻿using System;
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

            if (comboSheetList.Text.Trim() == "會議$")
            {
                string sql = "SELECT COUNT(*) FROM t_meetingrawdata";
                int line = 0;
                queryCount(sql, out line);
                if (line > 0)
                {
                    DialogResult dt = MessageBox.Show("There are " + line + " records in database ,r u sure to append the data?", "Append or Cancel?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dt == DialogResult.Cancel)
                        return;
                }

                updateRawData(ds, comboSheetList, p.WorkType.Meeting);
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
                            cmd.Parameters.Add(new SQLiteParameter(@"_weeklyworkfreq", DbType.Int16));
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
                            cmd.Parameters[12].Value = Convert.ToInt16(ds.Tables[combo.SelectedIndex].Rows[i]["周工作频率（次）"]);
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
                        }


                        cmd.ExecuteNonQuery();
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.Rollback();
                    MessageBox.Show(ex.Message, "INSERT FAIL", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    conn.Close ();
                    return false;
                }
                conn.Close ();
                sw.Stop();
                MessageBox.Show(string.Format("Insert into database {0} records,used time:{1}", ds.Tables[combo.SelectedIndex].Rows.Count, sw.Elapsed.ToString()));

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



    }
}
