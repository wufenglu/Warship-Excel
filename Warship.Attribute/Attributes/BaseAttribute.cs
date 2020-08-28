
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Attribute.Attributes
{
    /// <summary>
    /// 特性基类
    /// </summary>
    public abstract class BaseAttribute : System.Attribute
    {
        /// <summary>
        /// 定义字段
        /// </summary>
        private string _languageKey;

        /// <summary>
        /// 多语言Key
        /// </summary>
        public string LanguageKey
        {
            get { return this._languageKey; }
            set
            {
                this._languageKey = value;
            }
        }

        /// <summary>
        /// 多语言调用程序集
        /// </summary>
        public string LanguageAssembly {
            set {
                if (string.IsNullOrEmpty(_languageKey) == false)
                {
                    ErrorMessage = LanguageHelper.GetLanguageValue(_languageKey, value);
                }
            }
        }

        /// <summary>
        /// 提示文本
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}
