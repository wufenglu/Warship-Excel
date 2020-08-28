using Warship.Attribute.Enum;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Attribute.Model
{
    /// <summary>
    /// Excel头部信息
    /// </summary>
    public class ExcelHeadDTO
    {
        /// <summary>
        /// 属性名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 熟悉类型
        /// </summary>
        public Type PropertyType { set; get; }

        /// <summary>
        /// 头部名称
        /// </summary>
        public string HeadName { get; set; }

        /// <summary>
        /// 列序号
        /// </summary>
        public int ColumnIndex { get; set; }
        /// <summary>
        /// 排序字段，按照该字段的降序排列
        /// </summary>
        public int SortValue { set; get; }

        /// <summary>
        /// 整列背景色
        /// </summary>
        public short BackgroundColor { set; get; }

        /// <summary>
        /// 头部背景色
        /// </summary>
        public short HeaderBackgroundColor { set; get; }

        /// <summary>
        /// 是否验证头部
        /// </summary>
        public bool IsValidationHead { get; set; }

        /// <summary>
        /// 是否设置头部颜色：必填情况下才启用
        /// </summary>
        public bool IsSetHeadColor { get; set; }

        /// <summary>
        /// 是否锁定列
        /// </summary>
        public bool IsLocked { get; set; }

        /// <summary>
        /// 是否隐藏
        /// </summary>
        public bool IsHiddenColumn { get; set; }

        /// <summary>
        /// 列宽
        /// </summary>
        public int ColumnWidth { get; set; }

        /// <summary>
        /// 列类型
        /// </summary>
        public ColumnTypeEnum ColumnType { get; set; }

        /// <summary>
        /// 格式
        /// </summary>
        public string Format { get; set; }
    }
}
