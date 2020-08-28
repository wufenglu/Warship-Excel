using System;
using System.Collections.Generic;
using System.Linq;

namespace Warship.Utility
{
    /// <summary>
    /// 扩展类
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// 去空格
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string TrimStartEnd(this string value)
        {
            if (string.IsNullOrEmpty(value)) {
                return value;
            }
            return value.TrimStart().TrimEnd();
        }

        /// <summary>
        /// 判断集合是否为空
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public static bool IsNullOrEmpty<T>(this List<T> list)
        {
            if (list == null || list.Count == 0)
            {
                return true;
            }
            return false;
        }
    }
}
