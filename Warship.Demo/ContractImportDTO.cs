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
    public class ContractImportDTO : ExcelRowModel
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

        /// <summary>
        /// 材料
        /// </summary>
        [Sheet(SheetName = "合同材料",MasterEntityProperty = "Code",SlaveEntityProperty = "ContractCode")]
        [ExcelHead("合同材料", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        public List<ContractProductImportDTO> Products { get; set; }

        /// <summary>
        /// 合同金额
        /// </summary>
        [ExcelHead("合同金额", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        [Required(ErrorMessage = "合同金额必填")]
        [Range(Maximum =100,Minimum = 0, ErrorMessage = "合同编码必填")]
        public decimal? HtAmount { get; set; }
    }
}
