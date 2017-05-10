using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.SQLite;
using System.Drawing;
using System.Text.RegularExpressions;

namespace SummaryReportMeetingTips
{
    public  class p
    {

        #region 参数定义

        public static string appFolder = Application.StartupPath + @"\SummaryRepotMeeting";
        public static string appDataDB = appFolder + @"\DB.sqlite";
        public static string dbConnectionString = "Data Source=" + @appDataDB;


        public enum WorkType
        {
            Report,
            Meeting
        }

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
                    if (!createDefaultTable())
                    {
                        //File.Delete(appDataDB);
                        return false;
                    }


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


        public static bool createMeetingRawDataTable()
        {

            string sql = @"CREATE TABLE IF NOT EXISTS t_meetingrawdata(
id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT ,
depcode varchar(6),
seccode varchar(6),
opid varchar(9),
engname varchar(30),
meetingtype varchar(20),
workcontent varchar(255),
workdetail varchar(255),
worktype varchar(20),
isinworkbook varchar(3),
ismywork varchar(3),
vanva varchar(3),
singleworktime decimal(10,4) NULL,
weeklyworkfreq decimal(10,4) NULL,
weeklyworktime decimal(10,4) NULL,
monthlyworktime decimal(10,4) NULL,
caller varchar(10),
callerdep varchar(10),
callerlevel varchar(10),
optimizemethod varchar(20),
weeklysavetime decimal(10,4),
description varchar(255),
reviewdate varchar(20),
reviewer varchar(30))";
            if (!createTable(sql))
                return false;
            return true;
        }

        public static bool createReportRawDataTable()
        {

            string sql = @"CREATE TABLE IF NOT EXISTS t_reportrawdata(
id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT ,
depcode varchar(6),
seccode varchar(6),
opid varchar(9),
engname varchar(30),
reporttype varchar(20),
workcontent varchar(255),
workdetail varchar(255),
worktype varchar(20),
isinworkbook varchar(3),
ismywork varchar(3),
vanva varchar(3),
singleworktime decimal(10,4) NULL,
weeklyworkfreq decimal(10,4) NULL,
weeklyworktime decimal(10,4) NULL,
monthlyworktime decimal(10,4) NULL,
reportobject varchar(10),
reporttype2 varchar(10),
reportmethod varchar(10),
optimizemethod varchar(20),
weeklysavetime decimal(10,4),
description varchar(255),
reviewdate varchar(20),
reviewer varchar(30))";
            if (!createTable(sql))
                return false;
            return true;
        }

        public static bool createReportTipsDataTable()
        {
            string sql = @"CREATE TABLE IF NOT EXISTS t_reporttips(
workdetail varchar(255) NOT NULL PRIMARY KEY,
reporttype varchar(20),
tips int,
tipsavetime decimal(10,4),
reviewdate varchar(20))";

            if (!createTable(sql))
                return false;
            return true;     
        }

        public static bool createMeetingTipsDataTable()
        {

            string sql = @"CREATE TABLE IF NOT EXISTS t_meetingtips(
workdetail varchar(255) NOT NULL PRIMARY KEY,
meetingtype varchar(20),
tips int,
tipsavetime decimal(10,4),
reviewdate varchar(20))";

            if (!createTable(sql))
                return false;
            return true;   
        }


        public static string replaceString(string str)
        {
            if (str.Contains(@"/"))
                str = str.Replace(@"/", "_");
            if (str.Contains("\r\n"))
                str = str.Replace("\r\n", " ");
            if (str.Contains("\r"))
                str = str.Replace("\r", " ");
            if (str.Contains("\n"))
                str = str.Replace("\n", " ");
            if (str.Contains("&"))
                str = str.Replace ("&", " and ");
            return str;
        }


        public static bool createDefaultTable()
        {
            if (!createMeetingRawDataTable())
                return false;

            if (!createReportRawDataTable())
                return false;

            if (!createMeetingTipsDataTable())
                return false;

            if (!createReportTipsDataTable())
                return false;

            return true;
        }

        /// <summary>
        /// calc percentage 
        /// </summary>
        /// <param name="member">分子</param>
        /// <param name="denominator">分母</param>
        /// <returns></returns>
        public static string CalcPCT(decimal member, decimal denominator)
        {
            try
            {
                return string.Format("{0:0.00%}", member / denominator);
            }
            catch (Exception)
            {

                return "0.00%";
            }



        }


        /// <summary>
        /// 判断是否为工作日
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static bool IsWorkDay(DateTime dt)
        {
            //先从日期表中，查找不是上班时间，如果不是直接返回 false ，如果是，直接返回 true。
            //如果在日期表中，找不到，则查找定义的日历，依据日历定义的周末时间来定义是否为工作日。
            //获取日历中不上班的标准周末时间,判断是不是上班时间
            if (dt.DayOfWeek == DayOfWeek.Sunday || dt.DayOfWeek == DayOfWeek.Saturday)
                return false;
            else
                return true;
        }



        /// <summary>
        /// check the string if it is decimal
        /// </summary>
        /// <param name="str">string </param>
        /// <returns>Hex,return true;not hex,return false</returns>
        public static bool IsDecimal(string str)
        {
            return Regex.IsMatch(str, @"^[0-9,.]*$");
        }

        /// <summary>
        /// check the string if it is int
        /// </summary>
        /// <param name="str">string </param>
        /// <returns>Hex,return true;not hex,return false</returns>
        public static bool IsInt(string str)
        {
            return Regex.IsMatch(str, @"^[0-9]*$");
        }


        /// <summary>
        /// 设置ListItem的字体大小,颜色
        /// </summary>
        /// <param name="li">需要设置的那一项</param>
        /// <param name="fontSize">字体大小,如9</param>
        public static void SetListItemFont(ListViewItem li, int fontSize)//Color fontColor)
        {
            System.Drawing.Font myFont;
            string strName = "Calibri";
            FontStyle myFontStyle;
            int sngSize;
            sngSize = fontSize;
            //int intColorR = 255;
            //int intColorG = 0;
            //int intColorB = 0;
            myFontStyle = FontStyle.Bold;
            Color myColor;
            myColor = Color.Red;
            //myColor = fontColo

            FontFamily myFontFamily;
            myFontFamily = new FontFamily(strName);
            myFont = new Font(myFontFamily, sngSize, myFontStyle, GraphicsUnit.Point);
            li.Font = myFont;
        }


        /// <summary>
        /// 设置ListItem的字体大小,颜色
        /// </summary>
        /// <param name="li">需要设置的那一项</param>
        /// <param name="fontSize">字体大小,如9</param>
        public static void SetListItemFont(ListViewItem li, int fontSize, Color fontColor)
        {
            System.Drawing.Font myFont;
            string strName = "Calibri";
            FontStyle myFontStyle;
            int sngSize;
            sngSize = fontSize;
            //int intColorR = 255;
            //int intColorG = 0;
            //int intColorB = 0;
            myFontStyle = FontStyle.Bold;
            Color myColor;
            myColor = fontColor;
            //myColor = fontColo

            FontFamily myFontFamily;
            myFontFamily = new FontFamily(strName);
            myFont = new Font(myFontFamily, sngSize, myFontStyle, GraphicsUnit.Point);
            li.Font = myFont;
        }


        /// <summary>
        /// 查询记录条数
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
       public static int  queryCount(string sql)
        {

            SQLiteConnection conn = new SQLiteConnection(dbConnectionString);
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            var i = cmd.ExecuteScalar();
            conn.Close();
            return Convert.ToInt16(i);
        }

       /// <summary>
       /// 查询记录条数
       /// </summary>
       /// <param name="sql"></param>
       /// <returns></returns>
       public static decimal querySum(string sql)
       {

           SQLiteConnection conn = new SQLiteConnection(dbConnectionString);
           conn.Open();
           SQLiteCommand cmd = new SQLiteCommand(sql, conn);
           var i = cmd.ExecuteScalar();
           conn.Close();
           try
           {
               return decimal.Round(Convert.ToDecimal(i), 4);
           }
           catch (Exception)
           {

               return 0;
           }
          
       }
    }
}
