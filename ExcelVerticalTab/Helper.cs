using System;
using System.Collections.Concurrent;
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

        public static TValue GetValueOrDefault<TKey, TValue>(this ConcurrentDictionary<TKey, TValue> dict, TKey key, TValue defaultValue = default(TValue))
        {
            TValue value;
            return dict.TryGetValue(key, out value)
                ? value
                : defaultValue;
        }
    }
}
