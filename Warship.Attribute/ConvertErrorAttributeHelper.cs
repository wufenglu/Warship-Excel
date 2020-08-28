using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Warship.Attribute.Attributes;
using Warship.Attribute.Model;

namespace Warship.Attribute
{
    /// <summary>
    /// 转换失败特性帮助类
    /// </summary>
    public static class ConvertErrorAttributeHelper<TEntity> where TEntity : class, new()
    {
        /// <summary>
        /// 获取实体转换错误集合
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string,string> GetConvertErrors()
        {
            //结果集
            Dictionary<string, string> result = new Dictionary<string, string>();
            
            //遍历属性
            foreach (PropertyInfo info in typeof(TEntity).GetProperties())
            {
                //获取头部特性
                ConvertErrorAttribute attribute = info.GetCustomAttribute<ConvertErrorAttribute>();
                if (attribute != null)
                {
                    //添加至结果集
                    result.Add(info.Name, attribute.ErrorMessage);
                }
            }

            //返回结果集
            return result;
        }

        /// <summary>
        /// 获取属性转换异常信息
        /// </summary>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static string GetPropertyConvertErrorInfo(string propertyName)
        {
            //获取属性对应的头部实体
            Dictionary<string, string> dic = GetConvertErrors();
            if (dic.Count > 0 && dic.Keys.Contains(propertyName))
            {
                return dic[propertyName];
            }
            return null;
        }
    }
}
