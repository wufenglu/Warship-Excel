using Warship.Attribute.Attributes;
using Warship.Excel.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Excel.Model.Column;

namespace Warship.Demo.DemoDTO
{

    /// <summary>
    /// 材料
    /// </summary>
    [Serializable]
    public class ProductDTO : ExcelRowModel
    {

        /// <summary>
        /// 材料名称
        /// </summary>
        [ExcelHead("材料名称", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        [Required(ErrorMessage = "材料名称必填")]
        [Length(100, ErrorMessage = "长度不能超过100")]
        public string Name { get; set; }

        /// <summary>
        /// 材料分类编码
        /// </summary>
        [ExcelHead("材料分类编码", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        [Required(ErrorMessage = "材料分类编码必填")]
        [Length(100, ErrorMessage = "长度不能超过100")]
        public string Code { get; set; }

        /// <summary>
        /// 材料图片
        /// </summary>
        [ExcelHead("材料图片")]
        public List<ColumnFile> ProductPictures { set; get; }
    }
}
