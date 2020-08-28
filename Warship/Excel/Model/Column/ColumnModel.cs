using Warship.Attribute.Enum;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Attribute;
using Warship.Attribute.Model;

namespace Warship.Excel.Model.Column
{
    /// <summary>
    /// 列实体
    /// </summary>
    public class ColumnModel
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public ColumnModel() {
            ColumnValidations = new List<ColumnValidationModel>();
        }

        /// <summary>
        /// 属性名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 列序号
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 排序位置
        /// </summary>
        public int SortValue { set; get; }
        /// <summary>
        /// 列名
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// 列值
        /// </summary>
        public string ColumnValue { get; set; }

        /// <summary>
        /// 列验证集合
        /// </summary>
        public List<ColumnValidationModel> ColumnValidations { get; set; }

        /// <summary>
        /// 列错误信息
        /// </summary>
        public ColumnConvertErrorModel ColumnConvertError { get; set; }

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
        /// 背景色
        /// </summary>
        public short BackgroundColor { set; get; }

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
                    ColumnName = LanguageHelper.GetLanguageValue(LanguageKey, value);
                }
            }
        }
    }
}
