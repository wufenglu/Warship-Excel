using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Model
{
    /// <summary>
    /// 行样式设置
    /// </summary>
    public class RowStyleSet
    {
        /// <summary>
        /// 整行背景颜色
        /// </summary>
        public short? RowBackgroundColor { set; get; }

        /// <summary>
        /// 整行字体颜色
        /// </summary>
        public short? RowFontColor { set; get; }

        /// <summary>
        /// 单元格背景颜色
        /// </summary>
        public Dictionary<int, short> CellBackgroundColorDic { get; set; }

        /// <summary>
        /// 单元格字体颜色
        /// </summary>
        public Dictionary<int, short> CellFontColorDic { get; set; }
    }
}