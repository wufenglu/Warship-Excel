
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Attribute.Attributes
{
    /// <summary>
    /// 长度校验属性
    /// </summary>
    public sealed class LengthAttribute : BaseAttribute
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public LengthAttribute()
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="length"></param>
        public LengthAttribute(int length)
        {
            Length = length;
        }

        /// <summary>
        /// 长度
        /// </summary>
        public int Length { get; set; }
    }
}
