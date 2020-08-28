using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Attribute;

namespace Warship.Excel.Model.Column
{
    /// <summary>
    /// 列转换错误
    /// </summary>
    public class ColumnConvertErrorModel
    {
        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 多语言Key
        /// </summary>
        public string LanguageKey { get; set; }

        /// <summary>
        /// 多语言调用程序集
        /// </summary>
        public string LanguageAssembly
        {
            set
            {
                if (string.IsNullOrEmpty(LanguageKey) == false)
                {
                    ErrorMessage = LanguageHelper.GetLanguageValue(LanguageKey, value);
                }
            }
        }
    }
}
