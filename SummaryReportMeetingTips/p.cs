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




    }
}
