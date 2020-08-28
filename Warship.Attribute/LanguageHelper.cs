using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Attribute
{
    /// <summary>
    /// 多语言帮助
    /// </summary>
    public class LanguageHelper
    {
        /// <summary>
        /// 获取多语言的值
        /// </summary>
        /// <param name="languageKey"></param>
        /// <param name="assembly"></param>
        /// <returns></returns>
        public static string GetLanguageValue(string languageKey, string assembly = null)
        {
            string assemblyStr = "Mysoft.Map6.Platform";
            string classStr = "Mysoft.Map6.Platform.LangExtensions";
            string methodStr = "Translate";
            if (string.IsNullOrEmpty(assembly) == false)
            {
                string[] sp = assembly.Split(',');
                assemblyStr = sp[0];
                classStr = sp[1];
                methodStr = sp[2];
            }
            MethodInfo methodInfo = Assembly.Load(assemblyStr).GetType(classStr).GetMethod(methodStr);
            return methodInfo.Invoke(null, new object[] { languageKey }).ToString();
        }
    }
}
