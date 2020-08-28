using Warship.Attribute.Attributes;
using Warship.Excel.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Demo
{

    /// <summary>
    /// 合同
    /// </summary>
    [Serializable]
    public class DynamicColumnDTO : ExcelRowModel
    {

        /// <summary>
        /// 合同名称
        /// </summary>
        [ExcelHead("合同名称", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        [Required(ErrorMessage = "合同名称必填")]
        [Length(100, ErrorMessage = "长度不能超过100")]
        public string Name { get; set; }

        /// <summary>
        /// 合同编码
        /// </summary>
        [ExcelHead("合同编码", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        [Required(ErrorMessage = "合同编码必填")]
        [Length(100, ErrorMessage = "长度不能超过100")]
        public string Code { get; set; }
        
    }
}
