using Warship.Attribute.Attributes;
using Warship.Attribute.Enum;

namespace Warship.Excel.Model.Column
{
    /// <summary>
    /// 列验证信息
    /// </summary>
    public class ColumnValidationModel
    {
        /// <summary>
        /// 验证类型
        /// </summary>
        public ValidationTypeEnum ValidationTypeEnum { get; set; }

        /// <summary>
        /// 必填
        /// </summary>
        public RequiredAttribute RequiredAttribute { get; set; }

        /// <summary>
        /// 长度
        /// </summary>
        public LengthAttribute LengthAttribute { get; set; }

        /// <summary>
        /// 范围
        /// </summary>
        public RangeAttribute RangeAttribute { get; set; }

        /// <summary>
        /// 格式
        /// </summary>
        public FormatAttribute FormatAttribute { get; set; }
    }
}
