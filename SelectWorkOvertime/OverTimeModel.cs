using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Collections;

namespace SelectWorkOvertime
{
    public class OverTimeModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private void INotifyPropertyChanged(string name)
        {
            if(PropertyChanged!=null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }

        private string name;//姓名
        private string overtimeType;//加班类型
        private string overtimeStart;//加班时间起
        private string overtimeEnd;//加班时间止
        private string overtimeHours;//小时数
        private string remark;//备注

        public string Name//姓名
        {
            get
            {
                return name;
            }

            set
            {
                name = value;
                INotifyPropertyChanged("Name");
            }
        }        

        public string OvertimeType//加班类型
        {
            get
            {
                return overtimeType;
            }

            set
            {
                overtimeType = value;
                INotifyPropertyChanged("OvertimeType");
            }
        }

        public string OvertimeStart//加班时间起
        {
            get
            {
                return overtimeStart;
            }

            set
            {
                overtimeStart = value;
                INotifyPropertyChanged("OvertimeStart");
            }
        }

        public string OvertimeEnd//加班时间止
        {
            get
            {
                return overtimeEnd;
            }

            set
            {
                overtimeEnd = value;
                INotifyPropertyChanged("OvertimeEnd");
            }
        }

        public string OvertimeHours//小时数
        {
            get
            {
                return overtimeHours;
            }

            set
            {
                overtimeHours = value;
                INotifyPropertyChanged("OvertimeHours");
            }
        }

        public string Remark//备注
        {
            get
            {
                return remark;
            }

            set
            {
                remark = value;
                INotifyPropertyChanged("Remark");
            }
        }
    }
}
