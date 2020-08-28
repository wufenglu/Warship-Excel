using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Utility
{
    /// <summary>
    /// 扩展
    /// </summary>
    public static class ObjectExtension
    {
        /// <summary>
        /// Set字典
        /// </summary>
        public static Dictionary<string, object> SetDic { get; set; }

        /// <summary>
        /// Get字典
        /// </summary>
        public static Dictionary<string, object> GetDic { get; set; }

        /// <summary>
        /// 锁对象
        /// </summary>
        private static readonly object obj = new object();

        /// <summary>
        /// 构造函数
        /// </summary>
        static ObjectExtension()
        {
            SetDic = new Dictionary<string, object>();
            GetDic = new Dictionary<string, object>();
        }

        /// <summary>
        /// 获取属性值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="prop"></param>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static object GetValueExt<T>(this PropertyInfo prop, T obj)
        {
            return GetPropertyValue(obj, prop.Name);
        }

        /// <summary>
        /// 获取属性值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="entity"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static object GetPropertyValue<T>(this T entity, string name)
        {
            Type type = typeof(T);
            string key = type.FullName + "." + name;
            if (GetDic.Keys.Contains(key))
            {
                var funRe = GetDic[key] as Func<T, object>;
                return funRe(entity);
            }

            PropertyInfo prop = type.GetProperty(name);
            var entityParam = Expression.Parameter(type);
            var propExpress = Expression.Property(entityParam, prop);
            var bodyExpression = Expression.Convert(propExpress, typeof(object));
            var fun = Expression.Lambda<Func<T, object>>(bodyExpression, entityParam).Compile();
            GetDic[key] = fun;

            return fun(entity);
        }

        /// <summary>
        /// 设置属性值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="prop"></param>
        /// <param name="obj"></param>
        /// <param name="value"></param>
        /// <param name="index"></param>
        public static void SetValueExt<T>(this PropertyInfo prop, T obj, object value, object[] index)
        {
            SetPropertyValue(obj, prop.Name, value);
        }

        /// <summary>
        /// 设置属性值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="entity"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static void SetPropertyValue<T>(this T entity, string name, object value)
        {
            Type type = typeof(T);
            string key = type.FullName + "." + name;
            if (SetDic.Keys.Contains(key))
            {
                var fun = SetDic[key] as Action<T, object>;
                fun(entity, value);
                return;
            }

            lock (obj)
            {
                if (SetDic.Keys.Contains(key))
                {
                    var fun = SetDic[key] as Action<T, object>;
                    fun(entity, value);
                    return;
                }

                PropertyInfo p = type.GetProperty(name);
                var param_obj = Expression.Parameter(type);//实体参数值
                var param_value = Expression.Parameter(typeof(object));//属性参数值
                var setMethod = p.GetSetMethod(true);//获取设置值方法
                if (setMethod != null)
                {
                    var body_val = Expression.Convert(param_value, p.PropertyType);//属性值body
                    var body = Expression.Call(param_obj, p.GetSetMethod(true), body_val);
                    var fun = Expression.Lambda<Action<T, object>>(body, param_obj, param_value).Compile();
                    fun(entity, value);
                    SetDic[key] = fun;
                }
            }
        }
    }
}
