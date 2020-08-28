using Warship.Attribute.Attributes;
using Warship.Attribute.Enum;
using Warship.Excel.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static NPOI.HSSF.Util.HSSFColor;

namespace Warship.Demo.DemoDTO
{
    /// <summary>
    /// 采购合同批量调价
    /// </summary>
    public class ContractAdjustPriceExeclDTO : ExcelRowModel
    {
        /// <summary>
        /// 合同材料明细GUID
        /// </summary>
        [ExcelHead("合同明细主键", IsHiddenColumn = true)]
        public virtual Guid ContractProductDetailsGUID { set; get; }

        /// <summary>
        /// 材料编码
        /// </summary>        
        [ExcelHead(HeadName = "材料编码")]
        public virtual string ProductCode { set; get; }

        /// <summary>
        /// 材料名称
        /// </summary>
        [ExcelHead(HeadName = "材料名称")]
        public virtual string ProductName { set; get; }

        /// <summary>
        /// 品牌
        /// </summary>
        [ExcelHead(HeadName = "品牌")]
        public virtual string ProductBrand { set; get; }

        /// <summary>
        /// 型号
        /// </summary>
        [ExcelHead(HeadName = "型号")]
        public virtual string ProductModel { set; get; }

        /// <summary>
        /// 指标属性
        /// </summary>
        [ExcelHead(HeadName = "指标属性")]
        public string ProductAttribute { set; get; }

        /// <summary>
        /// 单位
        /// </summary>
        [ExcelHead(HeadName = "单位")]
        public virtual string ProductUnit { set; get; }

        /// <summary>
        /// 协议单价（含税）
        /// </summary>
        [ExcelHead(HeadName = "协议单价(含税)", ColumnType = ColumnTypeEnum.Decimal)]
        public virtual decimal? TacticCgAgreementPrice { set; get; }

        /// <summary>
        /// 协议运输单价（含税）
        /// </summary>
        [ExcelHead(HeadName = "协议运输单价(含税)", ColumnType = ColumnTypeEnum.Decimal)]
        public virtual decimal? FreightPrice { set; get; }

        /// <summary>
        /// 协议安装单价（含税）
        /// </summary>
        [ExcelHead(HeadName = "协议安装单价(含税)", ColumnType = ColumnTypeEnum.Decimal)]
        public virtual decimal? InstallPrice { set; get; }

        /// <summary>
        /// 数量
        /// </summary>
        [ExcelHead(HeadName = "数量", ColumnType = ColumnTypeEnum.Decimal, Format = "0.00")]
        [Range(-999999999d, 999999999d, ErrorMessage = "超出范围")]
        [Required(ErrorMessage = "必填")]
        public virtual decimal? Count { set; get; }

        /// <summary>
        /// 材料单价
        /// </summary>
        [ExcelHead(HeadName = "单价(含税)", ColumnType = ColumnTypeEnum.Decimal)]
        [Range(0, 999999999, ErrorMessage = "超出范围")]
        public virtual decimal? Price { set; get; }
        /// <summary>
        /// 税率
        /// </summary>
        [ExcelHead(HeadName = "税率(%)", ColumnType = ColumnTypeEnum.Decimal)]
        [Range(0, 999999999, ErrorMessage = "超出范围")]
        [Required(ErrorMessage = "必填")]
        public virtual decimal? TaxRate { set; get; }
        /// <summary>
        /// 不含税价格
        /// </summary>
        [ExcelHead(HeadName = "单价(不含税)", ColumnType = ColumnTypeEnum.Decimal)]
        [Range(0, 999999999, ErrorMessage = "超出范围")]
        public virtual decimal? NoTaxPrice { set; get; }

        /// <summary>
        /// 计划到货日期
        /// </summary>
        [ExcelHead(HeadName = "计划到货日期", ColumnType = ColumnTypeEnum.Date, Format = "yyyy-MM-dd")]
        public virtual DateTime? ExpectedArrivalDate { set; get; }

        /// <summary>
        /// 使用部位
        /// </summary>
        [ExcelHead(HeadName = "使用部位")]
        public virtual string UsePart { set; get; }
        /// <summary>
        /// 合同材料明细备注
        /// </summary>
        [ExcelHead(HeadName = "合同备注")]
        [Length(500, ErrorMessage ="超长")]
        public virtual string ContractProductRemark { set; get; }

    }
}
