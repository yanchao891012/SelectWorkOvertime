using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Media;

namespace SelectWorkOvertime
{
    public class StringConverToColor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if(value !=null && !string.IsNullOrEmpty(value.ToString()))
            {
                double valueInt = Double.Parse(value.ToString());
                if (valueInt < 0)
                    return Brushes.Red;
                if (valueInt > 0 && valueInt <= 0.5)
                    return Brushes.Yellow;
                if (valueInt >= 12)
                    return Brushes.Green;                
            }
            return Brushes.Black;
            //throw new NotImplementedException();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
