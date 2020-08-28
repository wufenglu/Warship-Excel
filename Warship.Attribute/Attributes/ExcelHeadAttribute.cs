
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Attribute.Enum;

namespace Warship.Attribute.Attributes
{
    /// <summary>
    /// 类型验证属性
    /// </summary>
    public sealed class ExcelHeadAttribute : System.Attribute
    {
        /// <summary>
        /// 构造函数，
        /// </summary>
        public ExcelHeadAttribute()
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="headName">头名称</param>
        /// <param name="isLocked">是否锁定</param>
        public ExcelHeadAttribute(string headName, bool isLocked = false)
        {
            HeadName = headName;
            IsLocked = isLocked;
        }

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
                HeadName = LanguageHelper.GetLanguageValue(_languageKey, value);
            }
        }

        /// <summary>
        /// 表头名称
        /// </summary>
        public string HeadName { get; set; }

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
        /// 列序号
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 按照该字段的值降序排序
        /// </summary>
        public int SortValue { set; get; }

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
        /// 整列颜色
        /// </summary>
        public short BackgroundColor { set; get; }

        /// <summary>
        /// 头部背景色
        /// </summary>
        public short HeaderBackgroundColor { set; get; }
    }
}
