using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Attribute.Attributes
{
    /// <summary>
    /// 范围校验属性
    /// </summary>
    public sealed class RangeAttribute : BaseAttribute
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public RangeAttribute()
        {
        }

        /// <summary>
        /// 范围校验属性初始化
        /// </summary>
        /// <param name="minimum">最小值</param>
        /// <param name="maximum">最大值</param>
        public RangeAttribute(int minimum, int maximum) {
            Minimum = minimum;
            Maximum = maximum;
        }

        /// <summary>
        /// 范围校验属性初始化
        /// </summary>
        /// <param name="minimum">最小值</param>
        /// <param name="maximum">最大值</param>
        public RangeAttribute(double minimum, double maximum)
        {
            Minimum = minimum;
            Maximum = maximum;
        }

        /// <summary>
        /// 范围校验属性初始化
        /// </summary>
        /// <param name="minimum">最小值</param>
        /// <param name="maximum">最大值</param>
        public RangeAttribute(decimal minimum, decimal maximum)
        {
            Minimum = minimum;
            Maximum = maximum;
        }

        /// <summary>
        /// 最大值
        /// </summary>
        public object Maximum { get; set; }

        /// <summary>
        /// 最小值
        /// </summary>
        public object Minimum { get; set; }        
    }
}
