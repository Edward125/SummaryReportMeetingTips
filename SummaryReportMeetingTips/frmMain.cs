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

        private void btnImportData_Click(object sender, EventArgs e)
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
    }
}
