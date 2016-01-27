using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SelectWorkOvertime
{
    /// <summary>
    /// 计算在不同的规则下加班的开始时间
    /// </summary>
    public class OverTimeStart
    {
        Tools tools = new Tools();

        /// <summary>
        /// 周一到周五 18:00以后下班，加班不超过当天
        /// </summary>
        /// <param name="dateTime">当天下班打卡时间</param>
        /// <returns></returns>
        public string OverTimeStart1(string dateTime)
        {
            DateTime dateTimeStart = Convert.ToDateTime(dateTime.Split(' ')[0].ToString() + " 18:00:00");
            return dateTimeStart.ToString("yyyy-MM-dd HH:mm");
        }
        /// <summary>
        /// 周一到周五 18:00以后下班，加班到第二天
        /// </summary>
        /// <param name="dateTime">当天下班打卡时间</param>
        /// <returns></returns>
        public string OverTimeStart2(string dateTime)
        {
            DateTime dateTimeStart = Convert.ToDateTime(tools.getDateBefore(dateTime).Split(' ')[0].ToString() + " 18:00:00");
            return dateTimeStart.ToString("yyyy-MM-dd HH:mm");
        }
        /// <summary>
        /// 周末或者假日 上班时间为9:00 之前
        /// </summary>
        /// <param name="dateTime">当天上班打卡时间</param>
        /// <returns></returns>
        public string OverTimeStart3(string dateTime)
        {
            DateTime dateTimeStart = Convert.ToDateTime(dateTime.Split(' ')[0].ToString() + " 9:00:00");
            return dateTimeStart.ToString("yyyy-MM-dd HH:mm");
        }
        /// <summary>
        /// 周末或者假日 上班时间为9:00 之后 
        /// </summary>
        /// <param name="dateTime">当天上班打卡时间</param>
        /// <returns></returns>
        public string OverTimeStart4(string dateTime)
        {
            DateTime dateTimeStart = Convert.ToDateTime(dateTime);
            return dateTimeStart.ToString("yyyy-MM-dd HH:mm");
        }
        /// <summary>
        /// 周末或者假日 打卡时间在12点到13点之间
        /// </summary>
        /// <param name="dateTime">当天上班打卡时间</param>
        /// <returns></returns>
        public string OverTimeStart5(string dateTime)
        {
            DateTime dateTimeStart = Convert.ToDateTime(dateTime.Split(' ')[0].ToString() + " 13:00:00");
            return dateTimeStart.ToString("yyyy-MM-dd HH:mm");
        }
    }
}
