using Warship.Attribute.Attributes;
using Warship.Excel.Model;
using System;
using System.Collections.Generic;
using Warship.Attribute.Enum;

namespace Warship.Demo
{

    /// <summary>
    /// 合同明细详细
    /// </summary>
    [Serializable]
    public class ContractProductImportDTO : ExcelRowModel
    {
        /// <summary>
        /// 合同编码
        /// </summary>
        [ExcelHead("合同编码", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        [Required(ErrorMessage = "合同编码必填")]
        [Length(100, ErrorMessage = "合同编码超过100")]
        public string ContractCode { get; set; }

        /// <summary>
        /// 材料名称
        /// </summary>
        [ExcelHead("材料名称", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        [Required(ErrorMessage = "材料名称必填")]
        [Length(100, ErrorMessage = "长度不能超过100")]
        public string Name { get; set; }

        /// <summary>
        /// 数量
        /// </summary>
        [ExcelHead("数量")]
        [Required(ErrorMessage = "数量必填")]
        [Range(0, 999999999, ErrorMessage = "范围必须在1-999999999之间")]
        [ConvertError(ErrorMessage = "转换失败，请输入数字")]
        public decimal? Count { get; set; }

        ///<summary>
        /// 单价（含税）
        ///</summary>
        [ExcelHead("单价（含税）",ColumnType = ColumnTypeEnum.Decimal)]
        [Required(ErrorMessage = "单价(含税)必填")]
        [Range(0, 10000, ErrorMessage = "范围必须在1-10000之间")]
        [ConvertError(ErrorMessage = "转换失败，请输入数字")]
        public decimal? Price { get; set; }

        /// <summary>
        /// 日期
        /// </summary>
        [ExcelHead("日期", ColumnType = ColumnTypeEnum.Date, Format = "yyyy-MM-dd")]
        public DateTime? Date { get; set; }

        /// <summary>
        /// 类型
        /// </summary>
        [ExcelHead("类型", ColumnType = ColumnTypeEnum.Option)]
        public int Type { get; set; }

        [ExcelHead("邮箱")]
        [Format(FormatEnum.Email,ErrorMessage ="格式错误")]
        public string Email { get; set; }

        /// <summary>
        /// 材料说明
        /// </summary>
        [Sheet(SheetName = "合同材料说明", MasterEntityProperty = "Name", SlaveEntityProperty = "Name")]
        [ExcelHead("合同材料说明", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        public List<ContractProductRemarkImportDTO> Remarks { get; set; }
    }
}
