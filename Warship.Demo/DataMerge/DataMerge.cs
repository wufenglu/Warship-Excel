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
    /// DataMerge
    /// </summary>
    [Serializable]
    [Sheet(SheetName = "RDC+更新服务")]
    public class DataMerge : ExcelRowModel
    {
        /// <summary>
        /// 客户ID
        /// </summary>
        [ExcelHead("客户ID")]
        public Guid ID { get; set; }

        /// <summary>
        /// 客户名称
        /// </summary>
        [ExcelHead("客户名称")]
        public string Name {get;set;}

        /// <summary>
        /// 当前版本
        /// </summary>
        [ExcelHead("当前版本")]
        public string Version { get; set; }

        /// <summary>
        /// 当前版本
        /// </summary>
        [ExcelHead("工作流版本")]
        public string WfVersion { get; set; }

        /// <summary>
        /// 材料图片
        /// </summary>
        [ExcelHead("区域")]
        public string Area { set; get; }

        /// <summary>
        /// 一线
        /// </summary>
        [ExcelHead("一线")]
        public string Admin { set; get; }

        /// <summary>
        /// 一线邮箱
        /// </summary>
        [ExcelHead("一线邮箱")]
        public string AdminEmail { set; get; }
    }
}
