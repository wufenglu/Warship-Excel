
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Attribute.Attributes
{
    /// <summary>
    /// 转换错误特性
    /// </summary>
    public sealed class SheetAttribute : System.Attribute
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
        public string LanguageAssembly
        {
            set
            {
                SheetName = LanguageHelper.GetLanguageValue(_languageKey, value);
            }
        }

        /// <summary>
        /// Sheet名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 主实体属性（用于建立关联关系）
        /// </summary>
        public string MasterEntityProperty { get; set; }

        /// <summary>
        /// 从实体属性（用于建立关联关系）
        /// </summary>
        public string SlaveEntityProperty { get; set; }

        /// <summary>
        /// 起始行
        /// </summary>
        public int StartRowIndex { get; set; }

        /// <summary>
        /// 起始列
        /// </summary>
        public int StartColumnIndex { get; set; }
    }
}
