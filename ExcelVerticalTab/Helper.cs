using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelVerticalTab
{
    public static class Helper
    {
        public static EventArgs<T> CreateEventArgs<T>(T value)
        {
            return new EventArgs<T>(value);
        } 
    }
}
