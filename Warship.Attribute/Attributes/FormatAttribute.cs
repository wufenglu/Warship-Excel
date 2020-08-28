using Warship.Attribute.Enum;
using Warship.Attribute.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Attribute.Attributes
{
    /// <summary>
    /// 类型验证属性
    /// </summary>
    public sealed class FormatAttribute : BaseAttribute
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="rormatEnum"></param>
        public FormatAttribute(FormatEnum rormatEnum)
        {
            FormatEnum = rormatEnum;
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="regex">正则表达式</param>
        public FormatAttribute(string regex)
        {
            Regex = regex;
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        public FormatAttribute()
        {
        }

        /// <summary>
        /// 正则表达式
        /// </summary>
        public string Regex { get; set; }

        /// <summary>
        /// 验证类型
        /// </summary>
        public FormatEnum FormatEnum { get; set; }
    }
}
