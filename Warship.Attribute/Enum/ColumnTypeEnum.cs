using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Attribute.Enum
{
    /// <summary>
    /// 列类型
    /// </summary>
    public enum ColumnTypeEnum
    {
        /// <summary>
        /// 文本
        /// </summary>
        Text = 0,
        /// <summary>
        /// 浮点
        /// </summary>
        Decimal = 1,
        /// <summary>
        /// 选项
        /// </summary>
        Option = 2,

        /// <summary>
        /// 日期
        /// </summary>
        Date = 3,

        /// <summary>
        /// 日期时间
        /// </summary>
        DateTime = 4,
    }
}
