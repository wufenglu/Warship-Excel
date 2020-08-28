using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Utility
{
    /// <summary>
    /// execl单元格样式帮助类
    /// </summary>
    internal static class CellStyleHelper
    {
        /// <summary>
        /// 用户有没有设置样式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public static bool IsNullOrEmpty(this ICellStyle cellStyle)
        {
            if (cellStyle == null)
            {
                return true;
            }
            //如果当前单元格四周都没有边框，则我们认为该单元格没有样式
            bool isNoStyle = cellStyle.BorderBottom == BorderStyle.None
                            && cellStyle.BorderLeft == BorderStyle.None
                            && cellStyle.BorderRight == BorderStyle.None
                            && cellStyle.BorderTop == BorderStyle.None;

            return isNoStyle;
        }


    }
}
