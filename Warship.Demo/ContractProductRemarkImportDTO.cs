using Warship.Attribute.Attributes;
using Warship.Excel.Model;
using System;

namespace Warship.Demo
{

    /// <summary>
    /// 合同明细详细说明
    /// </summary>
    [Serializable]
    public class ContractProductRemarkImportDTO : ExcelRowModel
    {
        /// <summary>
        /// 材料名称
        /// </summary>
        [ExcelHead("材料名称", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        [Required(ErrorMessage = "材料名称必填")]
        [Length(100, ErrorMessage = "长度不能超过100")]
        public string Name { get; set; }        

        [ExcelHead("说明")]
        public string Remark { get; set; }
    }
}
