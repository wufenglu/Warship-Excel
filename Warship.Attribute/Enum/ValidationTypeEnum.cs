using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Attribute.Enum
{
    /// <summary>
    /// 验证类型
    /// </summary>
    public enum ValidationTypeEnum
    {
        /// <summary>
        /// 必填
        /// </summary>
        Required,
        /// <summary>
        /// 长度
        /// </summary>
        Length,
        /// <summary>
        /// 范围
        /// </summary>
        Range,
        /// <summary>
        /// 格式
        /// </summary>
        Format
    }
}
