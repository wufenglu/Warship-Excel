using Warship.Attribute.Enum;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Model
{
    /// <summary>
    /// 列错误信息
    /// </summary>
    public class ColumnErrorMessage
    {
        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
        /// <summary>
        /// 列表头名称
        /// </summary>
        public string ColumnName { get; set; }
        /// <summary>
        /// 属性名称
        /// </summary>
        public string PropertyName { get; set; }
    }
}
