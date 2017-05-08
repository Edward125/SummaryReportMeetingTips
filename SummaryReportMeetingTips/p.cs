using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.SQLite;

namespace SummaryReportMeetingTips
{
    public  class p
    {

        #region 参数定义

        public static string appFolder = Application.StartupPath + @"\SummaryRepotMeeting";
        public static string appDataDB = appFolder + @"\DB.sqlite";
        public static string dbConnectionString = "Data Source=" + @appDataDB;

        #endregion

        /// <summary>
        /// check folder,if not exist,create it
        /// </summary>
        /// <returns></returns>
        public static bool CheckFolder()
        {
            if (!Directory.Exists(appFolder))
            {
                try
                {
                    Directory.CreateDirectory(appFolder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Create directory fail,detail:" + ex.Message, "Create Directory Fail...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                     return false;
                }
            }
            return true;
        }

        /// <summary>
        /// init database
        /// </summary>
        /// <returns></returns>
        public static bool InitDataDB()
        {
            if (!File.Exists(appDataDB))
            {
                try
                {
                    SQLiteConnection.CreateFile(appDataDB);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Create Data base fail,detail:" + ex.Message, "Create DB Fail...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }


            return true;
        }

        // string sql = "create table highscores (name varchar(20), score int)";
        /// <summary>
        /// create table 
        /// </summary>
        /// <param name="sql">sql</param>
        /// <returns>create ok,return true;create ng,return false</returns>
        public static bool createTable(string sql)
        {
            SQLiteConnection conn = new SQLiteConnection(dbConnectionString);
            try
            {
                conn.Open();
            }
            catch (Exception ex)
            {

                MessageBox.Show("Connect to database fail," + ex.Message);
                return false;
            }

            try
            {
                SQLiteCommand command = new SQLiteCommand(sql, conn);
                command.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show("Create TABLE fail," + ex.Message);
                conn.Close();
                return false;

            }

            return true;
        }

        /// <summary>
        /// update data to sqlite
        /// </summary>
        /// <param name="sql">sql</param>
        /// <returns>success,return true;fail,return false</returns>
        public static bool updateData2DB(string sql)
        {
            SQLiteConnection conn = new SQLiteConnection(dbConnectionString);

            try
            {
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand(sql, conn);
                //cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
                conn.Close();

            }
            catch (Exception ex)
            {
                conn.Close();
                MessageBox.Show("Execute sql: " + sql + " fail," + ex.Message);
                return false;
            }
            return true;
        }


        public static bool createRawDataTable()
        {

            string sql = "CREATE TABLE IF NOT EXIST t_rawdata(depcode varchar(6),seccode varchar(6),opid varchar(9),engname varchar(30),reportmeetingtype varchar(20),workcontent varchar(255),workdetail varchar(255),worktype varchar(20),isinworkbook varchar(3),ismywork varchar(3),singleworktime decimal,weeklyworkfre int,weeklyworktime decimal,monthlyworktime decimal,";


            return true;
        }


    }
}
