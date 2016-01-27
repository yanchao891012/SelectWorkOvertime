using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SelectWorkOvertime
{
    public class Tools
    {
        /// <summary>
        /// 读取TXT文档
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public string ReadTxt(string path)
        {
            string lines = "";
            StreamReader streamReader = new StreamReader(path, Encoding.Default);
            string line;
            while ((line = streamReader.ReadLine()) != null)
            {
                lines += line.ToString() + " ";
            }
            return lines;
        }
        /// <summary>
        /// 调用远程接口
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public string IsHoliday(string date)
        {
            string url = @"http://www.easybots.cn/api/holiday.php?d=";
            url = url + date;
            HttpWebRequest httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Timeout = 20000;
            httpRequest.Method = "GET";
            HttpWebResponse httpResponse = (HttpWebResponse)httpRequest.GetResponse();
            StreamReader sr = new StreamReader(httpResponse.GetResponseStream(), System.Text.Encoding.GetEncoding("gb2312"));
            string result = sr.ReadToEnd();
            result = result.Replace("\r", "").Replace("\n", "").Replace("\t", "");
            int status = (int)httpResponse.StatusCode;
            sr.Close();
            return result;
        }
        /// <summary>
        /// 读取Excel到Dataset
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public DataSet ExcelToDS(string Path, string fileName)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "select * from [" + fileName + "$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            return ds;
        }
        /// <summary>
        /// 获取日期的类型，是工作日，周末或节假日
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public string getDateType(string dateTime)
        {
            string date = Convert.ToDateTime(dateTime.Split(' ')[0].ToString()).ToString("yyyyMMdd");//获得到日期
            string isHoliday = IsHoliday(date);
            string numHoliday = isHoliday.Substring(isHoliday.Length - 3, 1);
            if (numHoliday == "1" || numHoliday == "2")//判断是不是节假日{2}，或者周末{1}
            {
                return numHoliday;
            }
            else
                return "0";//返回工作日{0}
        }
        /// <summary>
        /// 获取小时值
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public string getTimeHour(string dateTime)
        {
            string timeSFM = dateTime.Split(' ')[1].ToString();//时分秒
            string timeHour = timeSFM.Split(':')[0].ToString();
            return timeHour;
        }
        /// <summary>
        /// 计算小时差值 正常情况
        /// </summary>
        /// <param name="dateStart"></param>
        /// <param name="dateEnd"></param>
        /// <returns></returns>
        public string CalTimesNormal(DateTime dateStart, DateTime dateEnd)
        {
            string numTime;
            TimeSpan ts = dateEnd.Subtract(dateStart);
            if (ts.Minutes >= 30)
            {
                numTime = (ts.Hours + 0.5).ToString();
            }
            else
            {
                numTime = ts.Hours.ToString();
            }
            return numTime;
        }
        /// <summary>
        /// 计算小时差值 上班时间在12点之前的加班要减去1小时
        /// </summary>
        /// <param name="dateStart"></param>
        /// <param name="dateEnd"></param>
        /// <returns></returns>
        public string CalTimesAllDay(DateTime dateStart, DateTime dateEnd)
        {
            string numTime;
            TimeSpan ts = dateEnd.Subtract(dateStart);
            if (ts.Minutes >= 30)
            {
                numTime = (ts.Hours + 0.5 - 1).ToString();
            }
            else
            {
                numTime = (ts.Hours - 1).ToString();
            }
            return numTime;
        }
        /// <summary>
        /// 获取上一天
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public string getDateBefore(string dateTime)
        {
            string dateBefore = Convert.ToDateTime(dateTime).AddDays(-1).ToString();
            return dateBefore;
        }
        /// <summary>
        /// 比较两个时间是否相同
        /// </summary>
        /// <param name="dateTime1"></param>
        /// <param name="dateTime2"></param>
        /// <returns></returns>
        public int getCompareDate(string dateTime1, string dateTime2)
        {
            DateTime dt1 = Convert.ToDateTime(dateTime1.Split(' ')[0].ToString());
            DateTime dt2 = Convert.ToDateTime(dateTime2.Split(' ')[0].ToString());
            if (dt1 == dt2)
                return 1;
            return 0;
        }
    }
}
