using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Utility
{
    /// <summary>
    /// 属性帮助
    /// </summary>
    public class PropertyHelper
    {
        /// <summary>
        /// 实体属性字典
        /// </summary>
        public static Dictionary<string, PropertyInfo[]> EntityPropertysDic { get; set; }

        /// <summary>
        /// 锁对象
        /// </summary>
        private static readonly object obj = new object();

        /// <summary>
        /// 构造函数
        /// </summary>
        static PropertyHelper()
        {
            EntityPropertysDic = new Dictionary<string, PropertyInfo[]>();
        }

        /// <summary>
        /// 获取实体属性
        /// </summary>
        /// <returns></returns>
        public static PropertyInfo[] GetPropertys<T>()
        {
            Type type = typeof(T);

            string key = type.FullName;
            if (EntityPropertysDic.Keys.Contains(key))
            {
                return EntityPropertysDic[key];
            }
            lock (obj)
            {
                if (EntityPropertysDic.Keys.Contains(key))
                {
                    return EntityPropertysDic[key];
                }

                EntityPropertysDic[key] = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
                return EntityPropertysDic[key];
            }
        }

        /// <summary>
        /// 获取实体属性
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="name"></param>
        public static PropertyInfo GetPropertyInfo<T>(string name)
        {
            PropertyInfo[] list = GetPropertys<T>();
            return list.Where(w => w.Name == name).FirstOrDefault();
        }
    }
}
