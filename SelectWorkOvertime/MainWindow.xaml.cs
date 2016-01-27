using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel;

namespace SelectWorkOvertime
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        List<OverTimeModel> otmList = new List<OverTimeModel>();
        Tools tools = new Tools();
        OverTimeStart overTimeStart = new OverTimeStart();
        string fileName = "";
        DataTable dataTable = null;
        string[] nameList = null;

        BackgroundWorker bgWait;

        public MainWindow()
        {
            InitializeComponent();
        }
        public event PropertyChangedEventHandler PropertyChanged;
        private void INotifyPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }
        /// <summary>
        /// 打开打卡Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLoadCard_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLS文件(.xls)|*.xls";
            openFileDialog.ShowDialog();
            txtCardInfo.Text = openFileDialog.FileName;
            if (!string.IsNullOrEmpty(txtCardInfo.Text))
            {
                fileName = System.IO.Path.GetFileNameWithoutExtension(txtCardInfo.Text);
                dataTable = tools.ExcelToDS(txtCardInfo.Text, fileName).Tables[0];
            }
        }
        /// <summary>
        /// 打开姓名TXT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLoadName_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "TXT文件(.txt)|*.txt";
            openFileDialog.ShowDialog();
            if (!string.IsNullOrEmpty(openFileDialog.FileName))
            {
                txtNameInfo.Text = tools.ReadTxt(openFileDialog.FileName);
            }
        }
        /// <summary>
        /// 查询数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSearchInfo_Click(object sender, RoutedEventArgs e)
        {
            nameList = txtNameInfo.Text.Split(' ');
            DGShow.ItemsSource = null;
            if (nameList.Count() <= 0 || dataTable == null)
            {
                MessageBox.Show("请选择要查询的文件和人名");
                return;
            }

            bgWait = new BackgroundWorker();
            bgWait.WorkerReportsProgress = true;
            bgWait.DoWork += BgWait_DoWork;
            bgWait.RunWorkerCompleted += BgWait_RunWorkerCompleted;
            bgWait.RunWorkerAsync();
        }

        private void BgWait_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            DpWait.Visibility = Visibility.Collapsed;
            //throw new NotImplementedException();
        }

        private void BgWait_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(
                () =>
                {
                    DpWait.Visibility = Visibility.Visible;
                }));

            List<OverTimeModel> list = GetDepentInfo(nameList, dataTable);
            List<OverTimeModel> overTimeInfoList = GetOverTimeInfo(list);
            var gropResult = from p in overTimeInfoList
                             group p by new { p.Name, p.OvertimeStart } into g
                             select new OverTimeModel { Name = g.Key.Name, OvertimeStart = g.Key.OvertimeStart, OvertimeType = g.Max(p => p.OvertimeType), OvertimeEnd = g.Max(p => p.OvertimeEnd), OvertimeHours = g.Max(p => p.OvertimeHours), Remark = g.Max(p => p.Remark) };
           
            this.DGShow.Dispatcher.Invoke(new Action(() => { this.DGShow.ItemsSource = gropResult.ToList(); }));
            otmList = gropResult.ToList();
            if (gropResult.Count() <= 0)
            {
                MessageBox.Show("没有查询到加班数据");
            }
            //throw new NotImplementedException();
        }

        /// <summary>
        /// 从原始的打卡信息中取出姓名和打卡时间
        /// </summary>
        /// <param name="list"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public List<OverTimeModel> GetDepentInfo(string[] list, DataTable dt)
        {
            int count = list.Length;
            List<OverTimeModel> overTimeList = new List<OverTimeModel>();
            for (int i = 0; i < count; i++)
            {
                var Result = (from table in dataTable.AsEnumerable()
                              where table["NAME"].ToString() == nameList[i].ToString()
                              select new OverTimeModel()
                              {
                                  Name = table["NAME"].ToString(),
                                  OvertimeEnd = Convert.ToDateTime(table["CHECKTIME"]).ToString("yyyy-MM-dd HH:mm")
                              }).ToList();
                overTimeList.AddRange(Result);
            }
            return overTimeList;
        }
        /// <summary>
        /// 通过判断，进行相对应查询
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public List<OverTimeModel> GetOverTimeInfo(List<OverTimeModel> list)
        {
            List<OverTimeModel> overTimeList = new List<OverTimeModel>();
            OverTimeModel overTimeModel;

            for (int i = 0; i < list.Count; i++)
            {
                overTimeModel = new OverTimeModel();
                string dateType = tools.getDateType(list[i].OvertimeEnd);
                string timeHour = tools.getTimeHour(list[i].OvertimeEnd);
                int timeHourInt = Int32.Parse(timeHour);
                if (dateType == "0")//工作日
                {
                    if (timeHourInt >= 19)
                    {
                        overTimeModel = OverTimeRules(1, list, i);
                    }
                    if (timeHourInt < 7)
                    {
                        overTimeModel = OverTimeRules(2, list, i);
                    }
                }
                else//周末，假日
                {
                    if (timeHourInt < 7)
                    {
                        if (tools.getDateType(tools.getDateBefore(list[i].OvertimeEnd)) == "0")
                        {//特殊情况 获取当前的日期是周末，但是出于7点之前，为周五加班
                            overTimeModel = OverTimeRules(2, list, i);
                        }
                        else
                        {
                            if (tools.getCompareDate(tools.getDateBefore(list[i].OvertimeEnd).Split(' ')[0], list[i - 1].OvertimeEnd.Split(' ')[0]) == 1)
                                overTimeModel = OverTimeRules(5, list, i);
                        }
                    }
                    if (timeHourInt >= 8 && timeHourInt <= 9)
                    {
                        if (tools.getCompareDate(list[i].OvertimeEnd, list[i + 1].OvertimeEnd) == 1)
                            overTimeModel = OverTimeRules(3, list, i);
                    }
                    if (timeHourInt > 9)
                    {
                        if (tools.getCompareDate(list[i].OvertimeEnd, list[i + 1].OvertimeEnd) == 1)
                            overTimeModel = OverTimeRules(4, list, i);
                    }
                    if (dateType == "1")
                        overTimeModel.OvertimeType = "周末加班";
                    if (dateType == "2")
                        overTimeModel.OvertimeType = "节假日加班";
                }
                if (overTimeModel.Name != null)
                {
                    overTimeList.Add(overTimeModel);
                }

            }
            return overTimeList;
        }
        /// <summary>
        /// 根据设定规则赋值
        /// </summary>
        /// <param name="N"></param>
        /// <param name="list"></param>
        /// <param name="overTimeModel"></param>
        /// <returns></returns>
        public OverTimeModel OverTimeRules(int N, List<OverTimeModel> list, int i)
        {
            OverTimeModel ovm = new OverTimeModel();
            if (N == 1)//工作日加班，不超过12点
            {
                ovm.Name = list[i].Name;
                ovm.OvertimeStart = overTimeStart.OverTimeStart1(list[i].OvertimeEnd);
                ovm.OvertimeEnd = list[i].OvertimeEnd;
                ovm.OvertimeType = "夜间加班";
                ovm.OvertimeHours = tools.CalTimesNormal(Convert.ToDateTime(ovm.OvertimeStart), Convert.ToDateTime(ovm.OvertimeEnd));
            }
            if (N == 2)//工作日加班，超过12点
            {
                ovm.Name = list[i].Name;
                ovm.OvertimeStart = overTimeStart.OverTimeStart2(list[i].OvertimeEnd);
                ovm.OvertimeEnd = list[i].OvertimeEnd;
                ovm.OvertimeType = "夜间加班";
                ovm.OvertimeHours = tools.CalTimesNormal(Convert.ToDateTime(ovm.OvertimeStart), Convert.ToDateTime(ovm.OvertimeEnd));
            }
            if (N == 3)//周末或节假日 正常 在9点之前上班
            {
                ovm.Name = list[i].Name;
                ovm.OvertimeStart = overTimeStart.OverTimeStart3(list[i].OvertimeEnd);
                ovm.OvertimeEnd = list[i + 1].OvertimeEnd;//取下一条数据作为下班打卡
                ovm.OvertimeType = "周末或节假日加班";
                ovm.OvertimeHours = tools.CalTimesAllDay(Convert.ToDateTime(ovm.OvertimeStart), Convert.ToDateTime(ovm.OvertimeEnd));
            }
            if (N == 4)//周末或节假日 正常 在9点之后上班
            {
                ovm.Name = list[i].Name;
                ovm.OvertimeStart = overTimeStart.OverTimeStart4(list[i].OvertimeEnd);
                ovm.OvertimeEnd = list[i + 1].OvertimeEnd;//取下一条数据作为下班打卡
                ovm.OvertimeType = "周末或节假日加班";
                int overtimeStartHour = Int32.Parse(tools.getTimeHour(ovm.OvertimeStart));
                if (overtimeStartHour <= 12)//12点之前上班
                {
                    ovm.OvertimeHours = tools.CalTimesAllDay(Convert.ToDateTime(ovm.OvertimeStart), Convert.ToDateTime(ovm.OvertimeEnd));
                }
                else if (overtimeStartHour >= 13)//13点以后上班
                {
                    ovm.OvertimeHours = tools.CalTimesNormal(Convert.ToDateTime(ovm.OvertimeStart), Convert.ToDateTime(ovm.OvertimeEnd));
                }
                else//12点到13点之间上班
                {
                    ovm.OvertimeHours = tools.CalTimesNormal(Convert.ToDateTime(overTimeStart.OverTimeStart5(list[i].OvertimeEnd)), Convert.ToDateTime(ovm.OvertimeEnd));
                }
            }
            if (N == 5)//周末或节假日 超过12点 
            {
                ovm.Name = list[i].Name;
                ovm.OvertimeStart = overTimeStart.OverTimeStart4(list[i - 1].OvertimeEnd);//取上一条数据作为上班打卡
                ovm.OvertimeEnd = list[i].OvertimeEnd;
                ovm.OvertimeType = "周末或节假日加班";
                ovm.OvertimeHours = tools.CalTimesNormal(Convert.ToDateTime(ovm.OvertimeStart), Convert.ToDateTime(ovm.OvertimeEnd));
            }
            
            this.txtRemark.Dispatcher.Invoke(new Action(() => { ovm.Remark = this.txtRemark.Text; }));
            return ovm;
        }

        private Brush foregound=Brushes.Red;

        public Brush Foregound
        {
            get { return foregound; }
            set
            {
                foregound = value;
                INotifyPropertyChanged("Foregound");
            }
        }
        /// <summary>
        /// 跳转导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExprot_Click(object sender, RoutedEventArgs e)
        {
            if(otmList.Count>0)
            {
                ExportWindow ew = new ExportWindow(otmList);
                ew.WindowStartupLocation = WindowStartupLocation.CenterScreen;                
                ew.Owner = this;
                ew.ShowInTaskbar = false;
                ew.ShowDialog();
            }
            else
            {
                MessageBox.Show("没有需要导出的数据");
            }
        }


    }
}
